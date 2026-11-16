# Ghi log theo nhóm, mỗi ngày một file
import glob
import logging
import os
import threading
import time
from datetime import datetime

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOG_DIR = os.getenv("LOG_DIR", os.path.join(BASE_DIR, "logs"))
LOG_KEEP_DAYS = int(os.getenv("LOG_KEEP_DAYS", "30"))
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()

_FORMAT = logging.Formatter("%(asctime)s [%(levelname)s] [pid %(process)d] %(message)s")
_lock = threading.Lock()


# Mỗi ngày một file log
class DailyFileHandler(logging.FileHandler):
    def __init__(self, folder: str, prefix: str):
        self.folder, self.prefix = folder, prefix
        self.day = datetime.now().strftime("%Y-%m-%d")
        super().__init__(self._path(self.day), mode="a", encoding="utf-8", delay=True)

    def _path(self, day: str) -> str:
        return os.path.join(self.folder, f"{self.prefix}_{day}.log")

    def emit(self, record):
        day = datetime.now().strftime("%Y-%m-%d")
        if day != self.day:
            self.acquire()
            try:
                if day != self.day:
                    self.close()
                    self.day = day
                    self.baseFilename = self._path(day)
                    self._cleanup()
            finally:
                self.release()
        super().emit(record)

    def _cleanup(self):
        if LOG_KEEP_DAYS <= 0:
            return
        cutoff = time.time() - LOG_KEEP_DAYS * 86400
        for old in glob.glob(os.path.join(self.folder, f"{self.prefix}_*.log")):
            try:
                if os.path.getmtime(old) < cutoff:
                    os.remove(old)
            except OSError:
                pass


# Thêm job_id vào log
class JobAdapter(logging.LoggerAdapter):
    def process(self, msg, kwargs):
        return f"[{self.extra['job_id']}] {msg}", kwargs


# Lấy logger theo nhóm
def get_logger(category: str, job_id: str | None = None):
    logger = logging.getLogger(f"sort.{category}")
    with _lock:
        if not getattr(logger, "_sort_configured", False):
            folder = os.path.join(LOG_DIR, category)
            try:
                os.makedirs(folder, exist_ok=True)
                handler = DailyFileHandler(folder, category)
                handler.setFormatter(_FORMAT)
                handler._cleanup()
                logger.addHandler(handler)
            except OSError as e:
                logging.getLogger(__name__).warning(f"Không tạo được log file cho '{category}': {e}")
            logger.setLevel(LOG_LEVEL)
            logger.propagate = True
            logger._sort_configured = True
    return JobAdapter(logger, {"job_id": job_id}) if job_id else logger
