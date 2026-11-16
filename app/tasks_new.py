import os
import shutil
import gc
import json
import time
import threading
import datetime

from dotenv import load_dotenv

from .logger import get_logger, LOG_DIR

load_dotenv()


KEEP_TEMP = os.getenv("SORT_KEEP_TEMP", "0").lower() in ("1", "true", "t", "yes")
# Kết quả mỗi job ghi thêm 1 dòng vào results/sort_YYYY-MM-DD.jsonl (thay bảng sort_book_results)
RESULT_DIR = os.getenv("SORT_RESULT_DIR", os.path.join(os.path.dirname(LOG_DIR), "results"))
_result_lock = threading.Lock()


# Các luồng chạy theo mode
FLOW_TASKS = {
    "all":      ["name", "date", "metadata"],
    "title":    ["name"],
    "metadata": ["date"],
    "info":     ["metadata"],
    "sound":    ["sound"],
}


# Bước chuẩn bị: lấy file vào thư mục job
def sort_prepare(job_id: str, job_folder: str, pdf_path: str = "", docx_path: str = "",
                 mode: str = "all", pdf_src: str = "", docx_src: str = "",
                 audio_path: str = "", audio_src: str = "", bia_path: str = "", bia_src: str = "") -> dict:
    from .file_source import fetch_into
    from . import flow_sort_date, flow_sort_metadata, flow_sort_name

    log = get_logger("task", job_id)
    t0 = time.time()
    log.info(f"Prepare bắt đầu: pdf={pdf_path or pdf_src or '-'} word={docx_path or docx_src or '-'} "
             f"audio={audio_path or audio_src or '-'} bia={bia_path or bia_src or '-'} mode={mode}")
    context = {
        "flow_type":  "sort",
        "job_id":     job_id,
        "job_folder": job_folder,
        "pdf_path":   pdf_path,
        "docx_path":  docx_path,
        "pdf_src":    pdf_src,
        "docx_src":   docx_src,
        "audio_path": audio_path,
        "audio_src":  audio_src,
        "bia_paths":  [bia_path] if bia_path else [],
        "mode":       mode,
        "page_count": None,
        "text_source": "",
        "result": {} if mode == "sound" else {**flow_sort_name.EMPTY_RESULT, **flow_sort_date.EMPTY_RESULT,
                                              **flow_sort_metadata.EMPTY_RESULT},
    }
    try:
        if pdf_src and not pdf_path:
            context["pdf_path"] = fetch_into(pdf_src, job_folder, job_id)
        if docx_src and not docx_path:
            context["docx_path"] = fetch_into(docx_src, job_folder, job_id)
        if audio_src and not audio_path:
            context["audio_path"] = fetch_into(audio_src, job_folder, job_id)
        if bia_src and not bia_path:
            context["bia_paths"] = [fetch_into(bia_src, os.path.join(job_folder, "bia_goc"), job_id)]
        log.info(f"Prepare xong {time.time() - t0:.2f}s")
    except Exception as e:
        log.error(f"Prepare lỗi sau {time.time() - t0:.2f}s: {e}", exc_info=True)
        context["error"] = f"Lỗi chuẩn bị file: {e}"
    return context


# Chạy một luồng, lỗi thì ghi vào context để các bước sau bỏ qua
def _run_flow(context: dict, ten: str, buoc, loi: str) -> dict:
    log = get_logger("task", context.get("job_id"))
    if context.get("error"):
        log.warning(f"{ten} bỏ qua do lỗi trước: {context['error']}")
        return context
    t0 = time.time()
    try:
        context.setdefault("result", {}).update(buoc(context))
        log.info(f"{ten} xong {time.time() - t0:.2f}s")
    except Exception as e:
        log.error(f"{ten} lỗi sau {time.time() - t0:.2f}s: {e}", exc_info=True)
        context["error"] = f"{loi}: {e}"
    return context


# Luồng tên sách
def sort_name(context: dict) -> dict:
    from . import flow_sort_name

    def buoc(ctx):
        flow_sort_name.name_prepare(ctx)
        flow_sort_name.name_ocr(ctx)
        r = flow_sort_name.name_result(ctx)
        get_logger("task", ctx.get("job_id")).info(f"Tên sách ({r['text_source'] or '-'}): {r['ten_sach']!r}")
        return r
    return _run_flow(context, "Luồng tên sách", buoc, "Lỗi trích xuất tên sách")


# Luồng tháng, năm lưu chiểu
def sort_date(context: dict) -> dict:
    from . import flow_sort_date

    def buoc(ctx):
        flow_sort_date.date_prepare(ctx)
        flow_sort_date.date_ocr(ctx)
        return flow_sort_date.date_result(ctx)
    return _run_flow(context, "Luồng ngày lưu chiểu", buoc, "Lỗi trích xuất ngày lưu chiểu")


# Luồng thông tin trang cuối
def sort_metadata(context: dict) -> dict:
    from . import flow_sort_metadata

    def buoc(ctx):
        flow_sort_metadata.info_prepare(ctx)
        flow_sort_metadata.info_ocr(ctx)
        return flow_sort_metadata.info_result(ctx)
    return _run_flow(context, "Luồng thông tin trang cuối", buoc, "Lỗi trích xuất thông tin trang cuối")


# Luồng âm thanh
def sort_sound(context: dict) -> dict:
    from . import flow_sort_sound

    def buoc(ctx):
        flow_sort_sound.sound_prepare(ctx)
        r = flow_sort_sound.sound_result(ctx)
        r["audio_file"] = ctx.get("audio_src") or os.path.basename(ctx.get("audio_path", ""))
        return r
    return _run_flow(context, "Luồng âm thanh", buoc, "Lỗi đọc file âm thanh")


FLOWS = {"name": sort_name, "date": sort_date, "metadata": sort_metadata, "sound": sort_sound}


# Ghi kết quả ra file jsonl
def _ghi_ket_qua(row: dict):
    os.makedirs(RESULT_DIR, exist_ok=True)
    path = os.path.join(RESULT_DIR, f"sort_{datetime.date.today():%Y-%m-%d}.jsonl")
    with _result_lock, open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(row, ensure_ascii=False) + "\n")
    return path


# Bước cuối: ghi kết quả, dọn file tạm
def sort_finish(context: dict) -> dict:
    job_id     = context.get("job_id")
    job_folder = context.get("job_folder")
    error      = context.get("error", "")
    result     = context.get("result", {})
    status     = "failed" if error else "completed"
    if context.get("mode") != "sound":
        result.update(text_source=context.get("text_source", ""), page_count=context.get("page_count"),
                      pdf_file=context.get("pdf_src") or os.path.basename(context.get("pdf_path", "")),
                      docx_file=context.get("docx_src") or os.path.basename(context.get("docx_path", "")))

    log = get_logger("database", job_id)
    try:
        path = _ghi_ket_qua({
            "job_id": job_id, "created_at": datetime.datetime.now().isoformat(timespec="seconds"),
            "mode": context.get("mode", "all"), "status": status, "error": error,
            "pdf_file": os.path.basename(context.get("pdf_path", "")),
            "docx_file": os.path.basename(context.get("docx_path", "")),
            "audio_file": os.path.basename(context.get("audio_path", "")),
            "source_text": context.get("source_text", ""), "result": result,
        })
        log.info(f"Ghi kết quả ({status}) vào {path}, ten_sach={result.get('ten_sach', '')!r}")
    except Exception as e:
        log.error(f"Ghi kết quả lỗi: {e}", exc_info=True)

    task_log = get_logger("task", job_id)
    if job_folder and os.path.exists(job_folder) and not KEEP_TEMP:
        shutil.rmtree(job_folder, ignore_errors=True)
        task_log.info(f"Đã xóa temp {job_folder}")
    task_log.info(f"Kết thúc job: {status}{' - ' + error if error else ''}")
    gc.collect()
    return {"job_id": job_id, "status": status, "error": error, "result": result}


# Chạy cả job: chuẩn bị -> các luồng theo mode -> ghi kết quả
def run_job(**kwargs) -> dict:
    context = sort_prepare(**kwargs)
    for name in FLOW_TASKS[context["mode"]]:
        context = FLOWS[name](context)
    return sort_finish(context)
