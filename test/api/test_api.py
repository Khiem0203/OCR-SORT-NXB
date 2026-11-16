# Test API và hàng đợi job (app/jobs.py), giả lập các luồng nên không cần Chandra / Qwen / Spellfix
import glob
import json
import os
import shutil
import sys
import tempfile
import time
import unittest
from unittest import mock

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, ROOT)

# Ghi kết quả, log vào thư mục tạm
TMP = tempfile.mkdtemp(prefix="test_api_")
os.environ["SORT_RESULT_DIR"] = os.path.join(TMP, "results")
os.environ["LOG_DIR"] = os.path.join(TMP, "logs")

from fastapi.testclient import TestClient
from app import jobs, main, tasks_new

PDF = ("a.pdf", b"%PDF-1.4 x", "application/pdf")


# Luồng giả: chờ delay giây rồi ghi key = val vào kết quả, loi=True thì báo lỗi
def luong_gia(key="", val="", delay=0.0, loi=False):
    def f(ctx):
        time.sleep(delay)
        if loi:
            ctx["error"] = "Lỗi giả"
        else:
            ctx["result"][key] = val
        return ctx
    return f


# Hỏi request-status đến khi xong
def cho_xong(c, rid, giay=10):
    for _ in range(int(giay * 10)):
        st = c.get(f"/api/request-status/{rid}").json()
        if st["status"] in ("SUCCESS", "FAILURE"):
            return st
        time.sleep(0.1)
    return st


class Api(unittest.TestCase):
    def setUp(self):
        self.c = TestClient(main.app)

    def test_ten_sach_cho_ket_qua(self):
        with mock.patch.dict(tasks_new.FLOWS, name=luong_gia("ten_sach", "Sách A")):
            r = self.c.post("/api/sort/title", files={"pdf": PDF})
        self.assertEqual(r.status_code, 200, r.text)
        b = r.json()
        self.assertEqual((b["status"], b["result"]["ten_sach"]), ("completed", "Sách A"))
        st = self.c.get(f"/api/request-status/{b['request_id']}").json()
        self.assertEqual((st["status"], st["result"]["result"]["ten_sach"]), ("SUCCESS", "Sách A"))
        # Kết quả ghi vào results/*.jsonl, thư mục tạm của job đã xoá
        rows = []
        for p in glob.glob(os.path.join(TMP, "results", "*.jsonl")):
            with open(p, encoding="utf-8") as f:
                rows += [json.loads(l) for l in f]
        self.assertIn(b["request_id"], [x["job_id"] for x in rows])
        self.assertFalse(os.path.exists(os.path.join(main.TEMP_DIR, b["request_id"])))

    def test_mode_all_chay_du_3_luong(self):
        with mock.patch.dict(tasks_new.FLOWS, name=luong_gia("ten_sach", "A"), date=luong_gia("luu_chieu_nam", 2020),
                             metadata=luong_gia("thong_tin", {"x": 1})):
            b = self.c.post("/api/sort", files={"pdf": PDF}).json()
        self.assertEqual((b["result"]["ten_sach"], b["result"]["luu_chieu_nam"], b["result"]["thong_tin"]),
                         ("A", 2020, {"x": 1}))

    def test_wait_false_tra_request_id(self):
        with mock.patch.dict(tasks_new.FLOWS, name=luong_gia("ten_sach", "B", delay=0.5)):
            r = self.c.post("/api/sort/title", files={"pdf": PDF}, data={"wait": "false"})
            self.assertEqual(r.status_code, 202)
            rid = r.json()["request_id"]
            self.assertIn(self.c.get(f"/api/request-status/{rid}").json()["status"], ("PENDING", "STARTED"))
            st = cho_xong(self.c, rid)
        self.assertEqual(st["result"]["result"]["ten_sach"], "B")

    def test_qua_timeout_tra_202_job_chay_tiep(self):
        with mock.patch.dict(tasks_new.FLOWS, name=luong_gia("ten_sach", "C", delay=1.5)):
            r = self.c.post("/api/sort/title", files={"pdf": PDF}, data={"timeout": "1"})
            self.assertEqual(r.status_code, 202)
            st = cho_xong(self.c, r.json()["request_id"])
        self.assertEqual(st["status"], "SUCCESS")

    def test_hai_file_chay_song_song(self):
        if jobs.SORT_WORKERS < 2:
            self.skipTest("SORT_WORKERS < 2")
        with mock.patch.dict(tasks_new.FLOWS, name=luong_gia("ten_sach", "D", delay=1.0)):
            t0 = time.time()
            rids = [self.c.post("/api/sort/title", files={"pdf": PDF}, data={"wait": "false"}).json()["request_id"]
                    for _ in range(2)]
            sts = [cho_xong(self.c, rid) for rid in rids]
        self.assertTrue(all(s["status"] == "SUCCESS" for s in sts))
        self.assertLess(time.time() - t0, 1.8)

    def test_luong_loi_tra_500(self):
        with mock.patch.dict(tasks_new.FLOWS, name=luong_gia(loi=True)):
            r = self.c.post("/api/sort/title", files={"pdf": PDF})
        self.assertEqual((r.status_code, r.json()["status"]), (500, "failed"))

    def test_job_sap_bao_failure(self):
        with mock.patch.dict(tasks_new.FLOWS, name=luong_gia("ten_sach", "E")), \
             mock.patch.object(tasks_new, "sort_finish", side_effect=RuntimeError("hỏng")):
            r = self.c.post("/api/sort/title", files={"pdf": PDF}, data={"wait": "false"})
            st = cho_xong(self.c, r.json()["request_id"])
        # sort_finish bị giả lập nên không xoá temp, test tự xoá
        shutil.rmtree(os.path.join(main.TEMP_DIR, r.json()["request_id"]), ignore_errors=True)
        self.assertEqual(st["status"], "FAILURE")

    def test_thieu_file_400_id_la_404(self):
        self.assertEqual(self.c.post("/api/sort/title").status_code, 400)
        self.assertEqual(self.c.get("/api/request-status/khong-co").status_code, 404)
        self.assertEqual(self.c.get("/api/jobs/status").json()["workers"], jobs.SORT_WORKERS)
        self.assertEqual(self.c.get("/health").json()["status"], "healthy")

    def test_khong_dung_celery_postgres(self):
        self.assertNotIn("celery", sys.modules)
        self.assertNotIn("psycopg2", sys.modules)


if __name__ == "__main__":
    unittest.main()
