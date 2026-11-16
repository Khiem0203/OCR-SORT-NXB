"""Hàng đợi job chạy ngay trong tiến trình API (thay Celery + RabbitMQ + Postgres)."""
import os
import threading
import time
from concurrent.futures import Future, ThreadPoolExecutor

from .logger import get_logger
from .tasks_new import run_job

SORT_WORKERS = int(os.getenv("SORT_WORKERS", "4"))                     # số job chạy cùng lúc
JOB_KEEP_HOURS = float(os.getenv("SORT_JOB_KEEP_HOURS", "24"))          # giữ trạng thái job đã xong

_pool = ThreadPoolExecutor(max_workers=max(SORT_WORKERS, 1), thread_name_prefix="sort")
_jobs: dict[str, dict] = {}
_lock = threading.Lock()


# Xoá trạng thái job xong quá lâu
def _don_job_cu():
    han = time.time() - JOB_KEEP_HOURS * 3600
    with _lock:
        for k in [k for k, v in _jobs.items() if v.get("done_at") and v["done_at"] < han]:
            del _jobs[k]


def _chay(job_id: str, kwargs: dict) -> dict:
    with _lock:
        _jobs[job_id]["state"] = "STARTED"
    try:
        out = run_job(**kwargs)
        state, info = "SUCCESS", out
    except Exception as e:
        get_logger("task", job_id).error(f"Job lỗi: {e}", exc_info=True)
        state, info = "FAILURE", f"{type(e).__name__}: {e}"
    with _lock:
        _jobs[job_id].update(state=state, result=info, done_at=time.time())
    if state == "FAILURE":
        raise RuntimeError(info)
    return info


# Đưa job vào hàng đợi
def submit(job_id: str, kwargs: dict) -> Future:
    _don_job_cu()
    with _lock:
        _jobs[job_id] = {"state": "PENDING", "result": None, "done_at": None}
    return _pool.submit(_chay, job_id, kwargs)


# Trạng thái job: PENDING / STARTED / SUCCESS / FAILURE, None nếu không có
def status(job_id: str):
    with _lock:
        j = _jobs.get(job_id)
        return None if j is None else (j["state"], j["result"])


# Số job đang chờ / đang chạy
def thong_ke() -> dict:
    with _lock:
        states = [j["state"] for j in _jobs.values()]
    return {"workers": SORT_WORKERS, "pending": states.count("PENDING"), "running": states.count("STARTED"),
            "success": states.count("SUCCESS"), "failure": states.count("FAILURE")}
