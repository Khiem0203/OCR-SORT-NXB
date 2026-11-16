# Lấy file đầu vào và lọc file
import os
import re
import shutil
import time
import unicodedata

import requests
from dotenv import load_dotenv

from .logger import get_logger

load_dotenv()

WINDOWS_MOUNT = os.getenv("SORT_WINDOWS_MOUNT", "/mnt")
ALLOWED_ROOTS = [r for r in os.getenv("SORT_ALLOWED_ROOTS", "").split(",") if r.strip()]
ALLOW_URL = os.getenv("SORT_ALLOW_URL", "0").lower() in ("1", "true", "t", "yes")

_WIN_DRIVE_RE = re.compile(r"^([A-Za-z]):[\\/]*(.*)$")


# Kiểm tra có phải URL không
def is_url(path: str) -> bool:
    return path.lower().startswith(("http://", "https://"))


# Đổi đường dẫn Windows sang WSL
def to_local_path(path: str) -> str:
    path = path.strip().strip('"')
    m = _WIN_DRIVE_RE.match(path)
    if m and os.name != "nt":
        rest = m.group(2).replace("\\", "/")
        path = f"{WINDOWS_MOUNT}/{m.group(1).lower()}/{rest}"
    return os.path.normpath(path)


# Kiểm tra đường dẫn được phép đọc
def _check_allowed(path: str):
    if not ALLOWED_ROOTS:
        return
    real = os.path.realpath(path)
    roots = [os.path.realpath(to_local_path(r)) for r in ALLOWED_ROOTS]
    if not any(real == r or real.startswith(r.rstrip(os.sep) + os.sep) for r in roots):
        raise ValueError(f"Đường dẫn ngoài thư mục cho phép: {path}")


# Kiểm tra file local hợp lệ
def _local_file(path: str, ext) -> str:
    local = to_local_path(path)
    _check_allowed(local)
    if not local.lower().endswith(ext):
        raise ValueError(f"File '{path}' không phải {' / '.join(ext) if isinstance(ext, tuple) else ext}")
    if not os.path.isfile(local):
        raise ValueError(f"Không tìm thấy file: {path} ({local})")
    return local


_BIA_RE = re.compile(r"(?<![a-z0-9])bia(?![a-z0-9])")
_PHAN_SO_RE = re.compile(r"-\s*\d+\s*$")
_ANH_EXT = (".jpg", ".jpeg")


# Bỏ dấu tiếng Việt
def _khong_dau(s: str) -> str:
    s = s.replace("đ", "d").replace("Đ", "D")
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").lower()


# Lý do bỏ qua file, rỗng nếu dùng được
def skip_reason(name: str) -> str:
    stem, ext = os.path.splitext(os.path.basename(name))
    ext = ext.lower()
    if ext in _ANH_EXT:
        return "ảnh jpg"
    if _BIA_RE.search(_khong_dau(stem)):
        return "tên có chữ Bìa"
    if ext == ".pdf" and _PHAN_SO_RE.search(stem):
        return "PDF tên kết thúc '- số'"
    return ""


WORD_EXT = (".docx", ".doc")


# File Word dùng được
def _is_docx(name: str) -> bool:
    return name.lower().endswith(WORD_EXT) and not name.startswith("~$") and not skip_reason(name)


# Thứ tự ưu tiên Word: .docx trước
def _word_rank(name: str) -> tuple:
    return (not name.lower().endswith(".docx"), name)


# File PDF dùng được
def _is_pdf(name: str) -> bool:
    return name.lower().endswith(".pdf") and not skip_reason(name)


# Tìm Word đi kèm PDF
def _find_docx_for(pdf_local: str) -> str:
    folder = os.path.dirname(pdf_local)
    stem = os.path.splitext(os.path.basename(pdf_local))[0].lower()
    docx = sorted((n for n in os.listdir(folder) if _is_docx(n)), key=_word_rank)
    same = [n for n in docx if os.path.splitext(n)[0].lower() == stem]
    pick = same or docx
    return os.path.join(folder, pick[0]) if pick else ""


# Kiểm tra và trả đường dẫn PDF, Word
def resolve_sources(pdf_path: str = "", word_path: str = "", folder_path: str = "") -> tuple[str, str]:
    pdf_src, docx_src = "", ""
    if folder_path:
        folder = to_local_path(folder_path)
        _check_allowed(folder)
        if not os.path.isdir(folder):
            raise ValueError(f"Không tìm thấy thư mục: {folder_path} ({folder})")
        names = sorted(os.listdir(folder))
        pdfs = [n for n in names if _is_pdf(n)]
        skipped = [n for n in names if n.lower().endswith((".pdf",) + WORD_EXT + _ANH_EXT) and skip_reason(n)]
        if skipped:
            get_logger("input").info(f"Bỏ qua trong {folder}: " + "; ".join(f"{n} ({skip_reason(n)})" for n in skipped))
        if pdfs and not pdf_path:
            pdf_src = os.path.join(folder, pdfs[0])
        docx = sorted((n for n in names if _is_docx(n)), key=_word_rank)
        if docx and not word_path and not pdfs:
            docx_src = os.path.join(folder, docx[0])

    for path in (pdf_path, word_path):
        if path and is_url(path) and not ALLOW_URL:
            raise ValueError(f"Không nhận URL ({path}); đặt SORT_ALLOW_URL=1 nếu muốn tải file qua http(s)")
    if pdf_path:
        pdf_src = pdf_path if is_url(pdf_path) else _local_file(pdf_path, ".pdf")
    if word_path:
        docx_src = word_path if is_url(word_path) else _local_file(word_path, WORD_EXT)
    elif pdf_src and not docx_src and not is_url(pdf_src):
        docx_src = _find_docx_for(pdf_src)
    get_logger("input").info(
        f"Đầu vào pdf_path={pdf_path!r} word_path={word_path!r} folder_path={folder_path!r} "
        f"→ pdf={pdf_src or '-'} word={docx_src or '-'}")
    return pdf_src, docx_src


AUDIO_EXT = (".mp3", ".m4a", ".m4b", ".aac", ".wav", ".flac", ".ogg", ".oga", ".opus", ".wma", ".amr", ".aiff", ".aif")


# File âm thanh dùng được
def _is_audio(name: str) -> bool:
    return name.lower().endswith(AUDIO_EXT) and not name.startswith("~$")


# Kiểm tra và trả đường dẫn file âm thanh (thư mục thì lấy file đầu tiên)
def resolve_audio(audio_path: str = "", folder_path: str = "") -> str:
    if audio_path:
        if is_url(audio_path):
            if not ALLOW_URL:
                raise ValueError(f"Không nhận URL ({audio_path}); đặt SORT_ALLOW_URL=1 nếu muốn tải file qua http(s)")
            src = audio_path
        else:
            src = _local_file(audio_path, AUDIO_EXT)
    elif folder_path:
        folder = to_local_path(folder_path)
        _check_allowed(folder)
        if not os.path.isdir(folder):
            raise ValueError(f"Không tìm thấy thư mục: {folder_path} ({folder})")
        names = sorted(n for n in os.listdir(folder) if _is_audio(n))
        if not names:
            raise ValueError(f"Thư mục {folder_path} không có file âm thanh ({', '.join(AUDIO_EXT)})")
        if len(names) > 1:
            get_logger("input").info(f"{folder} có {len(names)} file âm thanh → lấy {names[0]}")
        src = os.path.join(folder, names[0])
    else:
        src = ""
    get_logger("input").info(f"Đầu vào audio_path={audio_path!r} folder_path={folder_path!r} → audio={src or '-'}")
    return src


BIA_EXT = (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp", ".pdf")


# File bìa (ảnh hoặc PDF bìa): OCR khi Word / PDF ruột không ra tên sách
def resolve_bia(bia_path: str = "") -> str:
    if not bia_path:
        return ""
    if is_url(bia_path):
        if not ALLOW_URL:
            raise ValueError(f"Không nhận URL ({bia_path}); đặt SORT_ALLOW_URL=1 nếu muốn tải file qua http(s)")
        return bia_path
    return _local_file(bia_path, BIA_EXT)


# Copy hoặc tải file vào thư mục job
def fetch_into(src: str, folder: str, job_id: str = "") -> str:
    log = get_logger("input", job_id or None)
    os.makedirs(folder, exist_ok=True)
    t0 = time.time()
    try:
        if is_url(src):
            name = os.path.basename(src.split("?", 1)[0]) or "download"
            dest = os.path.join(folder, name)
            with requests.get(src, stream=True, timeout=300) as resp:
                resp.raise_for_status()
                with open(dest, "wb") as f:
                    for chunk in resp.iter_content(1024 * 1024):
                        f.write(chunk)
            action = "Tải"
        else:
            dest = os.path.join(folder, os.path.basename(src))
            shutil.copy2(src, dest)
            action = "Copy"
    except (requests.ConnectionError, requests.Timeout) as e:
        get_logger("connection", job_id or None).error(f"Tải {src}: không kết nối được: {e}")
        raise
    except Exception as e:
        log.error(f"Lấy file {src} lỗi: {e}")
        raise
    log.info(f"{action} {src} → {dest} ({os.path.getsize(dest) / 1024 / 1024:.1f} MB, {time.time() - t0:.2f}s)")
    return dest
