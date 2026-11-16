# Dịch vụ sửa chính tả ProtonX
import json
import logging
import os
import re
import threading
from typing import List

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

import torch
from fastapi import FastAPI
from pydantic import BaseModel
from transformers import AutoModelForSeq2SeqLM, AutoTokenizer

try:
    from .logger import get_logger
except ImportError:
    from logger import get_logger

logging.basicConfig(level=logging.INFO)
logger = get_logger("spellfix_service")

MODEL_PATH = os.getenv("SPELLFIX_MODEL", "protonx-models/protonx-legal-tc")
NUM_BEAMS = int(os.getenv("SPELLFIX_NUM_BEAMS", "4"))
MAX_TOKENS = int(os.getenv("SPELLFIX_MAX_TOKENS", "160"))
BATCH_SIZE = int(os.getenv("SPELLFIX_BATCH_SIZE", "16"))
DTYPE = {"float32": torch.float32, "float16": torch.float16, "bfloat16": torch.bfloat16}[
    os.getenv("SPELLFIX_DTYPE", "float32")]
MIN_LEN_RATIO = float(os.getenv("SPELLFIX_MIN_LEN_RATIO", "0.8"))
MAX_LEN_RATIO = float(os.getenv("SPELLFIX_MAX_LEN_RATIO", "1.25"))

_SENTENCE_SPLIT_RE = re.compile(r'(?<=[.;:!?\n])\s+')

_model = None
_tokenizer = None
_device = None
_lock = threading.Lock()


# Dữ liệu đầu vào
class CorrectRequest(BaseModel):
    texts: List[str]


# Dữ liệu trả về
class CorrectResponse(BaseModel):
    texts: List[str]


# Nạp tokenizer
def _load_tokenizer():
    try:
        return AutoTokenizer.from_pretrained(MODEL_PATH)
    except Exception as e:
        logger.warning(f"AutoTokenizer lỗi ({type(e).__name__}: {e}) → thử nạp thẳng tokenizer.json")

    tok_file = os.path.join(MODEL_PATH, "tokenizer.json")
    if not os.path.isfile(tok_file):
        files = sorted(os.listdir(MODEL_PATH)) if os.path.isdir(MODEL_PATH) else []
        raise RuntimeError(f"Không nạp được tokenizer và không có {tok_file} (thư mục có: {files})")
    try:
        from transformers import PreTrainedTokenizerFast
    except ImportError:
        from transformers import TokenizersBackend as PreTrainedTokenizerFast

    special = {"eos_token": "</s>", "pad_token": "<pad>", "unk_token": "<unk>"}
    map_file = os.path.join(MODEL_PATH, "special_tokens_map.json")
    if os.path.isfile(map_file):
        with open(map_file, encoding="utf-8") as f:
            for key, value in json.load(f).items():
                if key in special:
                    special[key] = value["content"] if isinstance(value, dict) else value
    tokenizer = PreTrainedTokenizerFast(tokenizer_file=tok_file, **special)
    logger.info(f"Đã nạp tokenizer từ {tok_file} ({special})")
    return tokenizer


# Nạp model một lần
def get_model():
    global _model, _tokenizer, _device
    with _lock:
        if _model is None:
            _device = "cuda" if torch.cuda.is_available() else "cpu"
            gpu = os.getenv("CUDA_VISIBLE_DEVICES", "all")
            logger.info(f"Loading spellfix model {MODEL_PATH} on {_device} (CUDA_VISIBLE_DEVICES={gpu}, {DTYPE})")
            _tokenizer = _load_tokenizer()
            try:
                _model = AutoModelForSeq2SeqLM.from_pretrained(MODEL_PATH, dtype=DTYPE).to(_device).eval()
            except torch.cuda.OutOfMemoryError:
                logger.error(f"Hết VRAM trên GPU {gpu} khi nạp model → chuyển sang CPU (float32, chậm hơn)")
                torch.cuda.empty_cache()
                _device = "cpu"
                _model = AutoModelForSeq2SeqLM.from_pretrained(MODEL_PATH, dtype=torch.float32).eval()
            if _device == "cuda":
                logger.info(f"Spellfix model đã nạp, VRAM dùng {torch.cuda.memory_allocated() / 1024**2:.0f} MiB")
    return _model, _tokenizer, _device


# Đếm số token
def _n_tokens(tokenizer, text: str) -> int:
    return len(tokenizer(text, add_special_tokens=True)["input_ids"])


# Cắt text thành đoạn theo số token
def _chunk(tokenizer, text: str, limit: int) -> List[str]:
    if _n_tokens(tokenizer, text) <= limit:
        return [text]
    pieces = []
    for sentence in _SENTENCE_SPLIT_RE.split(text):
        if _n_tokens(tokenizer, sentence) <= limit:
            pieces.append(sentence)
            continue
        words, current = sentence.split(), []
        for word in words:
            if current and _n_tokens(tokenizer, " ".join(current + [word])) > limit:
                pieces.append(" ".join(current))
                current = []
            current.append(word)
        if current:
            pieces.append(" ".join(current))
    chunks, current = [], ""
    for piece in pieces:
        candidate = f"{current} {piece}".strip()
        if current and _n_tokens(tokenizer, candidate) > limit:
            chunks.append(current)
            current = piece
        else:
            current = candidate
    if current:
        chunks.append(current)
    return chunks


# Kiểm tra bản sửa có nhận được không
def _accept(original: str, corrected: str) -> str:
    if not corrected:
        return original
    ratio = len(corrected) / max(len(original), 1)
    if MIN_LEN_RATIO <= ratio <= MAX_LEN_RATIO:
        return corrected
    logger.warning(f"REJECTED (length ratio {ratio:.2f} outside {MIN_LEN_RATIO}-{MAX_LEN_RATIO}), "
                   f"keep original: {original!r} | model output: {corrected!r}")
    return original


# Sửa chính tả danh sách text
def correct_texts(texts: List[str]) -> List[str]:
    model, tokenizer, device = get_model()
    limit = MAX_TOKENS - 10
    chunk_map, flat = [], []
    for i, text in enumerate(texts):
        stripped = (text or "").strip()
        if not stripped:
            continue
        for part in _chunk(tokenizer, stripped, limit):
            chunk_map.append(i)
            flat.append(part)

    corrected = []
    for start in range(0, len(flat), BATCH_SIZE):
        batch = flat[start:start + BATCH_SIZE]
        enc = tokenizer(batch, return_tensors="pt", padding=True, truncation=True,
                        max_length=MAX_TOKENS).to(device)
        with torch.no_grad():
            out = model.generate(**enc, num_beams=NUM_BEAMS, max_new_tokens=MAX_TOKENS, early_stopping=True)
        decoded = tokenizer.batch_decode(out, skip_special_tokens=True)
        corrected.extend(_accept(o, c.strip()) for o, c in zip(batch, decoded))

    results = [[] for _ in texts]
    for i, part in zip(chunk_map, corrected):
        results[i].append(part)
    return [" ".join(parts) if parts else (texts[i] or "") for i, parts in enumerate(results)]


app = FastAPI(title="ProtonX spellfix service")


# Nạp model khi khởi động
@app.on_event("startup")
def _warmup():
    get_model()


# API sửa chính tả
@app.post("/spellfix/correct", response_model=CorrectResponse)
def correct(req: CorrectRequest):
    import time
    t0 = time.time()
    texts = correct_texts(req.texts)
    changed = sum(1 for a, b in zip(req.texts, texts) if a != b)
    logger.info(f"request: {len(req.texts)} texts, {changed} changed, {time.time() - t0:.2f}s")
    return CorrectResponse(texts=texts)


# Kiểm tra dịch vụ
@app.get("/health")
def health():
    return {"status": "ok", "model": MODEL_PATH, "device": _device}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8006)
