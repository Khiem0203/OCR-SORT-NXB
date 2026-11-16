from fastapi import FastAPI, HTTPException, UploadFile, File, Form, Request
from fastapi.responses import JSONResponse
from starlette.datastructures import UploadFile as StarletteUploadFile
import os
import time
import uuid
import shutil
import logging
import asyncio
import datetime
from typing import Optional, Union
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv
from .file_source import resolve_sources, resolve_audio, resolve_bia, AUDIO_EXT, BIA_EXT, ALLOWED_ROOTS, ALLOW_URL
from .logger import get_logger

load_dotenv()


from . import jobs
from .models import SubmitResponse, StatusResponse


logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = get_logger("api")

TEMP_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "temp_processing")
DEFAULT_TIMEOUT = int(os.getenv("SORT_API_TIMEOUT", "600"))
TEMP_MAX_AGE_HOURS = float(os.getenv("SORT_TEMP_MAX_AGE_HOURS", "24"))

if not ALLOWED_ROOTS:
    logger.warning("SORT_ALLOWED_ROOTS đang trống: API được phép đọc mọi file .pdf/.docx/.doc máy chủ truy cập được")
if ALLOW_URL:
    logger.warning("SORT_ALLOW_URL=1: API sẽ tải file từ URL bất kỳ được gửi lên")


app = FastAPI(
    title="OCR Sort Book Service",
    description="Nhận diện tên sách (8 trang đầu / Word) và METADATA trang cuối từ file PDF",
    version="2.0.0"
)

# Ghi log mỗi request
@app.middleware("http")
async def _log_requests(request: Request, call_next):
    req_id = uuid.uuid4().hex[:8]
    client = request.client.host if request.client else "-"
    t0 = time.time()
    try:
        response = await call_next(request)
    except Exception:
        logger.exception(f"[{req_id}] {client} {request.method} {request.url.path} → 500 sau {time.time() - t0:.2f}s")
        raise
    level = logging.ERROR if response.status_code >= 500 else logging.WARNING if response.status_code >= 400 else logging.INFO
    logger.log(level, f"[{req_id}] {client} {request.method} {request.url.path} → {response.status_code} "
                      f"({time.time() - t0:.2f}s)")
    return response


# Kiểm tra cổng các dịch vụ model lúc khởi động
@app.on_event("startup")
def _check_services():
    import socket
    from urllib.parse import urlparse
    log = get_logger("connection")
    services = {
        "Chandra OCR": os.getenv("CHANDRA_URL", "http://localhost:8101/v1/chat/completions"),
        "Qwen LLM":    os.getenv("QWEN_LLM_URL", "http://localhost:8113/v1/chat/completions"),
        "Spellfix":    os.getenv("SPELLFIX_API_URL", "http://localhost:8006/spellfix/correct"),
    }
    log.info(f"API khởi động, {jobs.SORT_WORKERS} job chạy cùng lúc")
    for name, url in services.items():
        u = urlparse(url)
        host, port = u.hostname or "localhost", u.port or (443 if u.scheme == "https" else 80)
        try:
            socket.create_connection((host, port), timeout=3).close()
            log.info(f"  [OK]   {name:<12} {host}:{port}")
        except OSError as e:
            log.error(f"  [FAIL] {name:<12} {host}:{port} — {e}")


app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# API sắp xếp sách

# Bỏ trường file rỗng
def _as_upload(value) -> Optional[StarletteUploadFile]:
    if isinstance(value, StarletteUploadFile) and value.filename:
        return value
    return None


# Xóa thư mục tạm cũ
def _cleanup_old_temp():
    if TEMP_MAX_AGE_HOURS <= 0 or not os.path.isdir(TEMP_DIR):
        return
    cutoff = time.time() - TEMP_MAX_AGE_HOURS * 3600
    for name in os.listdir(TEMP_DIR):
        path = os.path.join(TEMP_DIR, name)
        try:
            if os.path.isdir(path) and os.path.getmtime(path) < cutoff:
                shutil.rmtree(path, ignore_errors=True)
                logger.info(f"Dọn temp cũ: {path}")
        except OSError:
            pass


# Lưu file upload
def _save_upload(upload: StarletteUploadFile, folder: str, allowed_ext: tuple) -> str:
    name = os.path.basename(upload.filename or "")
    if not name.lower().endswith(allowed_ext):
        raise HTTPException(status_code=400, detail=f"File '{name}' không hợp lệ, chỉ nhận {', '.join(allowed_ext)}")
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, name)
    with open(path, "wb") as f:
        shutil.copyfileobj(upload.file, f)
    return path


# Tạo job và chờ kết quả
async def _run_sort(mode: str, pdf, word, wait: bool, timeout: int,
                    pdf_path: str = "", word_path: str = "", folder_path: str = "",
                    audio=None, audio_path: str = "", bia=None, bia_path: str = ""):
    pdf, word, audio, bia = _as_upload(pdf), _as_upload(word), _as_upload(audio), _as_upload(bia)
    loop = asyncio.get_event_loop()
    loop.run_in_executor(None, _cleanup_old_temp)
    pdf_src, docx_src, audio_src, bia_src = "", "", "", ""
    try:
        bia_src = "" if bia else resolve_bia(bia_path)
        if mode == "sound":
            audio_src = "" if audio else resolve_audio(audio_path=audio_path, folder_path=folder_path)
        else:
            pdf_src, docx_src = resolve_sources(
                pdf_path="" if pdf else pdf_path,
                word_path="" if word else word_path,
                folder_path=folder_path,
            )
    except ValueError as e:
        logger.warning(f"/api/sort ({mode}) đầu vào không hợp lệ: {e}")
        raise HTTPException(status_code=400, detail=str(e))
    if mode == "sound":
        if audio is None and not audio_src:
            raise HTTPException(status_code=400, detail="Cần gửi file âm thanh (audio) hoặc đường dẫn (audio_path/folder_path)")
    elif pdf is None and not pdf_src and word is None and not docx_src and not (mode == "title" and (bia or bia_src)):
        raise HTTPException(status_code=400, detail="Cần gửi file (pdf/word) hoặc đường dẫn (pdf_path/word_path/folder_path)"
                                                    + ("; hoặc file bìa (bia/bia_path)" if mode == "title" else ""))

    job_id = uuid.uuid4().hex
    job_folder = os.path.join(TEMP_DIR, job_id)
    os.makedirs(job_folder, exist_ok=True)
    try:
        pdf_local = await loop.run_in_executor(None, _save_upload, pdf, job_folder, (".pdf",)) if pdf else ""
        docx_local = await loop.run_in_executor(None, _save_upload, word, job_folder, (".docx", ".doc")) if word else ""
        audio_local = await loop.run_in_executor(None, _save_upload, audio, job_folder, AUDIO_EXT) if audio else ""
        bia_local = await loop.run_in_executor(None, _save_upload, bia, os.path.join(job_folder, "bia_goc"), BIA_EXT) if bia else ""
    except HTTPException:
        shutil.rmtree(job_folder, ignore_errors=True)
        raise
    job_log = get_logger("api", job_id)
    job_log.info(f"Nhận job mode={mode} pdf={pdf_local or pdf_src or '-'} word={docx_local or docx_src or '-'} "
                 f"audio={audio_local or audio_src or '-'} bia={bia_local or bia_src or '-'} wait={wait} timeout={timeout}")

    future = jobs.submit(job_id, dict(job_id=job_id, job_folder=job_folder, pdf_path=pdf_local, docx_path=docx_local,
                                      mode=mode, pdf_src=pdf_src, docx_src=docx_src,
                                      audio_path=audio_local, audio_src=audio_src, bia_path=bia_local, bia_src=bia_src))

    pending = SubmitResponse(
        error_code=202,
        error_message="Đã nhận yêu cầu, tra kết quả bằng /api/request-status/{request_id}",
        request_id=job_id,
    )
    if not wait:
        job_log.info(f"Trả request_id={job_id} (wait=false)")
        return JSONResponse(status_code=202, content=pending.model_dump())

    t0 = time.time()
    try:
        output = await asyncio.wait_for(asyncio.shield(asyncio.wrap_future(future)), timeout)
    except asyncio.TimeoutError:
        job_log.warning(f"Quá {timeout}s chưa xong, trả request_id={job_id}")
        return JSONResponse(status_code=202, content=pending.model_dump())
    except Exception as e:
        job_log.error(f"Job lỗi sau {time.time() - t0:.2f}s: {e}")
        raise HTTPException(status_code=500, detail=f"Lỗi xử lý: {e}")
    output["request_id"] = job_id
    job_log.log(logging.ERROR if output.get("error") else logging.INFO,
                f"Kết quả sau {time.time() - t0:.2f}s: status={output.get('status')} "
                f"ten_sach={output.get('result', {}).get('ten_sach', '')!r} error={output.get('error') or '-'}")
    return JSONResponse(status_code=500 if output.get("error") else 200, content=output)


_PDF_PATH_DESC = r"Đường dẫn PDF ở thư mục ngoài (vd D:\sach\a.pdf hoặc /mnt/d/sach/a.pdf); URL http(s) cần SORT_ALLOW_URL=1"
_WORD_PATH_DESC = "Đường dẫn Word (.docx/.doc) ở thư mục ngoài hoặc URL; bỏ trống thì tự tìm Word cùng thư mục với PDF"
_FOLDER_DESC = "Thư mục chứa PDF (+ Word) của cuốn sách"
_BIA_DESC = "File bìa (ảnh .jpg/.png/.tif... hoặc PDF bìa) upload: Word / PDF không ra tên sách thì OCR bìa rồi tìm lại"
_BIA_PATH_DESC = "Đường dẫn file bìa ở thư mục ngoài hoặc URL (cần SORT_ALLOW_URL=1)"


# Lấy tên sách, tháng, năm và thông tin trang cuối
@app.post("/api/sort")
async def sort_book(
    pdf: Union[UploadFile, str, None] = File(None, description="File PDF sách (upload)"),
    word: Union[UploadFile, str, None] = File(None, description="File Word (.docx/.doc) upload, có thì đọc thay OCR 8 trang đầu"),
    pdf_path: str = Form("", description=_PDF_PATH_DESC),
    word_path: str = Form("", description=_WORD_PATH_DESC),
    folder_path: str = Form("", description=_FOLDER_DESC),
    bia: Union[UploadFile, str, None] = File(None, description=_BIA_DESC),
    bia_path: str = Form("", description=_BIA_PATH_DESC),
    wait: bool = Form(True, description="True: chờ và trả kết quả; False: trả request_id ngay"),
    timeout: int = Form(DEFAULT_TIMEOUT, description="Số giây chờ tối đa khi wait=True"),
):
    return await _run_sort("all", pdf, word, wait, timeout, pdf_path, word_path, folder_path, bia=bia, bia_path=bia_path)


# Chỉ lấy tên sách
@app.post("/api/sort/title")
async def sort_book_title(
    pdf: Union[UploadFile, str, None] = File(None, description="File PDF sách (upload)"),
    word: Union[UploadFile, str, None] = File(None, description="File Word (.docx/.doc) upload"),
    pdf_path: str = Form("", description=_PDF_PATH_DESC),
    word_path: str = Form("", description=_WORD_PATH_DESC),
    folder_path: str = Form("", description=_FOLDER_DESC),
    bia: Union[UploadFile, str, None] = File(None, description=_BIA_DESC + "; gửi riêng file bìa cũng được"),
    bia_path: str = Form("", description=_BIA_PATH_DESC),
    wait: bool = Form(True),
    timeout: int = Form(DEFAULT_TIMEOUT),
):
    return await _run_sort("title", pdf, word, wait, timeout, pdf_path, word_path, folder_path, bia=bia, bia_path=bia_path)


# Chỉ lấy tháng, năm lưu chiểu
@app.post("/api/sort/metadata")
async def sort_book_metadata(
    pdf: Union[UploadFile, str, None] = File(None, description="File PDF sách (upload)"),
    word: Union[UploadFile, str, None] = File(None, description="File Word (.docx/.doc) upload, có thì đọc trang cuối Word trước"),
    pdf_path: str = Form("", description=_PDF_PATH_DESC),
    word_path: str = Form("", description=_WORD_PATH_DESC),
    folder_path: str = Form("", description=_FOLDER_DESC),
    wait: bool = Form(True),
    timeout: int = Form(DEFAULT_TIMEOUT),
):
    return await _run_sort("metadata", pdf, word, wait, timeout, pdf_path, word_path, folder_path)


# Chỉ lấy thông tin xuất bản ở trang cuối
@app.post("/api/sort/info")
async def sort_book_info(
    pdf: Union[UploadFile, str, None] = File(None, description="File PDF sách (upload)"),
    word: Union[UploadFile, str, None] = File(None, description="File Word (.docx/.doc) upload, có thì đọc trang cuối Word trước"),
    pdf_path: str = Form("", description=_PDF_PATH_DESC),
    word_path: str = Form("", description=_WORD_PATH_DESC),
    folder_path: str = Form("", description=_FOLDER_DESC),
    wait: bool = Form(True),
    timeout: int = Form(DEFAULT_TIMEOUT),
):
    return await _run_sort("info", pdf, word, wait, timeout, pdf_path, word_path, folder_path)


# Lấy metadata file âm thanh (header + thẻ, chưa dùng AI)
@app.post("/api/sort/sound")
async def sort_sound(
    audio: Union[UploadFile, str, None] = File(None, description=f"File âm thanh upload ({', '.join(AUDIO_EXT)})"),
    audio_path: str = Form("", description="Đường dẫn file âm thanh ở thư mục ngoài hoặc URL (cần SORT_ALLOW_URL=1)"),
    folder_path: str = Form("", description="Thư mục chứa file âm thanh, lấy file đầu tiên theo tên"),
    wait: bool = Form(True),
    timeout: int = Form(DEFAULT_TIMEOUT),
):
    return await _run_sort("sound", None, None, wait, timeout, folder_path=folder_path,
                           audio=audio, audio_path=audio_path)


# Xem trạng thái job
@app.get("/api/request-status/{request_id}", response_model=StatusResponse)
async def get_request_status(request_id: str):
    st = jobs.status(request_id)
    if st is None:
        raise HTTPException(status_code=404, detail="Không có request_id này (đã quá hạn giữ hoặc API đã khởi động lại)")
    return StatusResponse(request_id=request_id, status=st[0], result=st[1])


# Số job đang chờ / đang chạy
@app.get("/api/jobs/status")
async def jobs_status():
    return jobs.thong_ke()


# Kiểm tra dịch vụ
@app.get("/health")
async def health_check():
    return {"status": "healthy", "timestamp": datetime.datetime.now().isoformat()}
