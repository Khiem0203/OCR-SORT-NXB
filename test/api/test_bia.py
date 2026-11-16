# Test luồng tên sách có file bìa: Word / PDF không ra tên -> OCR bìa -> Qwen (giả lập Chandra, Qwen, Spellfix)
import os
import shutil
import sys
import tempfile
import unittest
from unittest import mock

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, ROOT)

TMP = tempfile.mkdtemp(prefix="test_bia_")
os.environ["SORT_RESULT_DIR"] = os.path.join(TMP, "results")
os.environ["LOG_DIR"] = os.path.join(TMP, "logs")

import cv2
import numpy as np
from docx import Document
from fastapi.testclient import TestClient
from app import flow_sort_name, main


# Chandra giả: ảnh bìa (nền xám 200) ra chữ có tên sách, trang khác ra chữ thường
def chandra_gia(img, *a, **k):
    return "NHÀ XUẤT BẢN X\nCHỮ BÌA: SỔ TAY BÌA" if img is not None and int(img.mean()) == 200 else "Mục lục trang 1"


# Qwen giả: chỉ tìm ra tên khi có chữ bìa
def qwen_gia(text, job_id=""):
    return "SỔ TAY BÌA" if "CHỮ BÌA" in text else ""


def tao_file(ten, du_lieu=None):
    p = os.path.join(TMP, ten)
    if isinstance(du_lieu, bytes):  # file hỏng: ghi nguyên nội dung
        open(p, "wb").write(du_lieu)
    elif ten.endswith((".jpg", ".png")):
        cv2.imwrite(p, np.full((300, 200, 3), 200, np.uint8))
    elif ten.endswith(".pdf"):
        from PIL import Image
        Image.new("RGB", (200, 300), (200, 200, 200)).save(p, "PDF")
    elif ten.endswith(".docx"):
        d = Document(); d.add_paragraph(du_lieu or "Mục lục. Chương 1. Chương 2."); d.save(p)
    else:
        open(p, "wb").write(du_lieu or b"rac")
    return p


class Bia(unittest.TestCase):
    def setUp(self):
        self.c = TestClient(main.app)
        self.vao = [mock.patch.object(flow_sort_name, "chandra_ocr", side_effect=chandra_gia),
                    mock.patch.object(flow_sort_name, "qwen_find_title", side_effect=qwen_gia),
                    mock.patch.object(flow_sort_name, "spellfix", side_effect=lambda t, job_id="": t)]
        for m in self.vao:
            m.start()

    def tearDown(self):
        for m in self.vao:
            m.stop()

    def gui(self, mode="title", **files):
        mo = {k: (os.path.basename(p), open(p, "rb")) for k, p in files.items()}
        try:
            return self.c.post(f"/api/sort/{mode}" if mode else "/api/sort", files=mo)
        finally:
            for _, f in mo.values():
                f.close()

    def test_chi_co_bia_anh(self):
        r = self.gui(bia=tao_file("bia1.jpg"))
        self.assertEqual(r.status_code, 200, r.text)
        self.assertEqual((r.json()["result"]["ten_sach"], r.json()["result"]["text_source"]), ("SỔ TAY BÌA", "bia"))

    def test_word_khong_ra_ten_thi_doc_bia(self):
        r = self.gui(word=tao_file("a.docx"), bia=tao_file("bia2.png"))
        self.assertEqual(r.status_code, 200, r.text)
        self.assertEqual((r.json()["result"]["ten_sach"], r.json()["result"]["text_source"]), ("SỔ TAY BÌA", "bia"))

    def test_word_ra_ten_thi_khong_doc_bia(self):
        with mock.patch.object(flow_sort_name, "qwen_find_title", side_effect=lambda t, job_id="": "TÊN TỪ WORD"):
            r = self.gui(word=tao_file("b.docx"), bia=tao_file("bia3.jpg"))
        self.assertEqual((r.json()["result"]["ten_sach"], r.json()["result"]["text_source"]), ("TÊN TỪ WORD", "word"))

    def test_bia_pdf(self):
        r = self.gui(bia=tao_file("bia4.pdf"))
        self.assertEqual(r.json()["result"]["ten_sach"], "SỔ TAY BÌA", r.text)

    def test_khong_bia_khong_ra_ten(self):
        r = self.gui(word=tao_file("c.docx"))
        self.assertEqual((r.status_code, r.json()["result"]["ten_sach"]), (200, ""))

    def test_bia_hong_khong_lam_hong_job(self):
        r = self.gui(word=tao_file("d.docx"), bia=tao_file("hong.jpg", b"khong phai anh"))
        self.assertEqual((r.status_code, r.json()["result"]["ten_sach"], r.json()["status"]), (200, "", "completed"))

    def test_chi_bia_chi_nhan_o_title(self):
        self.assertEqual(self.gui("metadata", bia=tao_file("bia5.jpg")).status_code, 400)
        self.assertEqual(self.gui("title", bia=tao_file("bia6.docx")).status_code, 400)  # sai đuôi


if __name__ == "__main__":
    try:
        unittest.main()
    finally:
        shutil.rmtree(TMP, ignore_errors=True)
