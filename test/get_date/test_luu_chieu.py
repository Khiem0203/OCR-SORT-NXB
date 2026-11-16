# Test tách tháng / năm lưu chiểu: bắt buộc từ dòng có 'in xong / lưu chiểu' (không cần máy chủ)
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from app.flow_sort_date import colophon_block, extract_luu_chieu
from app.flow_sort_metadata import extract_thong_tin


def ngay(text):
    r = extract_luu_chieu(text)
    return r["thang"], r["nam"]


def ngay_tab7(text):
    r = extract_thong_tin(text)
    return r["luu_chieu_thang"], r["luu_chieu_nam"]


class LuuChieu(unittest.TestCase):
    def test_dong_in_xong_day_du(self):
        self.assertEqual(ngay("ISBN 978-604\nIn xong và nộp lưu chiểu tháng 02 năm 2021."), (2, 2021))
        self.assertEqual(ngay("In xong và nộp lưu chiểu Quý II/2019"), (None, 2019))

    def test_khong_co_in_xong(self):
        # không có dòng In xong: lấy ngày Quyết định xuất bản; Giấy đăng ký / Giấy phép / quyết định khác thì không
        t = "Quyết định xuất bản số 58/QĐ ngày 28 tháng 01 năm 2021\nGiấy đăng ký xuất bản số 313-2021"
        self.assertEqual(ngay(t), (1, 2021))
        self.assertEqual(ngay_tab7(t), (1, 2021))
        self.assertTrue(colophon_block(t).startswith("Quyết định xuất bản số 58/QĐ"))
        t = "Giấy đăng ký xuất bản số 313-2021/CXBIPH ngày 02/02/2021\nGiấy phép số 12/GP ngày 05/3/2020"
        self.assertEqual((ngay(t), ngay_tab7(t), colophon_block(t)), ((None, None), (None, None), ""))

    def test_chi_xet_tu_chu_in_xong_tro_di(self):
        t = "Quyết định số 5830 tháng 01/2020. In xong và nộp lưu chiểu tháng 3 năm 2021"
        self.assertEqual(ngay(t), (3, 2021))
        self.assertEqual(ngay_tab7(t), (3, 2021))
        # cùng dòng chỉ có năm của quyết định đứng trước chữ in xong: không lấy
        self.assertEqual(ngay("Quyết định 2019. In xong và nộp lưu chiểu"), (None, None))

    def test_xuong_dong_giua_cau(self):
        t = "Mã số ISBN: 978-604-57\nIn xong và nộp lưu chiểu\ntháng 5 năm 2022."
        self.assertEqual(ngay(t), (5, 2022))
        self.assertEqual(ngay_tab7(t), (5, 2022))
        self.assertTrue(colophon_block(t).endswith("tháng 5 năm 2022."))

    def test_loi_font(self):
        self.assertEqual(ngay("in xong vaø noäp lou chieau thaùng 3 naêm 2021"), (None, 2021))
        self.assertEqual(ngay("In xong và nộp lưu chiếu tháng 7 năm 2020"), (7, 2020))

    def test_cac_dang_thang_nam(self):
        dang = {
            "tháng 3-2021.": (3, 2021), "tháng 3 - 2021": (3, 2021), "tháng 3 – 2021": (3, 2021), "tháng 3—2021": (3, 2021),
            "tháng 03/2021": (3, 2021), "tháng 3.2021": (3, 2021), "tháng 3,2021": (3, 2021), "tháng 3 /2021": (3, 2021),
            "tháng 3 năm 2021": (3, 2021), "T3/2021": (3, 2021), "T.3/2021": (3, 2021), ": 3/2021": (3, 2021),
            "tháng 3-21": (3, 2021), "tháng 12 - 98": (12, 1998),
            "tháng ba năm 2021": (3, 2021), "tháng mười hai năm 2020": (12, 2020), "tháng năm năm 2019": (5, 2019),
            "tháng giêng năm 2022": (1, 2022), "quý I/2021": (None, 2021), "năm 2021": (None, 2021),
            "tháng năm 2021": (None, 2021),  # thiếu tháng: chỉ có năm
            "tháng 13/2021": (None, 2021),
        }
        for duoi, kq in dang.items():
            for tien_to in ("In xong và nộp lưu chiểu ", "In xong, nộp lưu chiểu "):
                self.assertEqual(ngay(tien_to + duoi), kq, duoi)
                self.assertEqual(ngay_tab7(tien_to + duoi), kq, duoi)
        # số đầu tiên sau chữ In xong mới tính dạng '3/2021' không có chữ tháng; số đăng ký phía sau không tính
        self.assertEqual(ngay("In xong và nộp lưu chiểu năm 2021. Số ĐKXB 12-2021/CXBIPH"), (None, 2021))

    def test_quyet_dinh_xuat_ban(self):
        from app.flow_sort_date import extract_luu_chieu as e
        # In xong không ra tháng + năm -> lấy ngày Quyết định xuất bản
        for t, kq in {
            "Quyết định xuất bản: Số 79-QĐ/NXBCTQG ngày 05/3/2021\nIn xong và nộp lưu chiểu năm 2021.": (3, 2021),
            "Quyết định xuất bản:\nSố 79-QĐ/NXBCTQG ngày 05/3/2021\nIn xong và nộp lưu chiểu.": (3, 2021),
            "Số QĐXB: 12/QĐ-NXBTP ngày 05 tháng 11 năm 2020\nIn xong và nộp lưu chiểu quý I/2021": (11, 2020),
            "Quyết định xuất bản số: 5830-QĐ/NXBCTQG,\nngày 28/01/2021.": (1, 2021),  # không có dòng In xong
            "Số 79-QĐ/NXBCTQG ngày 05-3-21": (3, 2021),
        }.items():
            self.assertEqual(ngay(t), kq, t)
            self.assertEqual(ngay_tab7(t), kq, t)
        r = e("Số 79-QĐ/NXBCTQG ngày 05/3/2021\nIn xong và nộp lưu chiểu năm 2021.")
        self.assertEqual((r["theo"], r["cau"]), ("qdxb", "Số 79-QĐ/NXBCTQG ngày 05/3/2021"))
        # In xong đủ tháng + năm thì vẫn ưu tiên In xong
        r = e("Quyết định xuất bản số 79-QĐ/NXBCTQG ngày 05/3/2021\nIn xong và nộp lưu chiểu tháng 4 năm 2021.")
        self.assertEqual((r["thang"], r["nam"], r["theo"]), (4, 2021, "in_xong"))
        # Quyết định khác (không phải QĐ xuất bản), Giấy đăng ký xuất bản: không lấy
        self.assertEqual(ngay("Quyết định số 15/2020/QĐ-TTg ngày 24/4/2020\nSố ĐKXB: 313-2021/CXBIPH ngày 02/02/2021"), (None, None))
        # khối thông tin xuất bản giữ cả dòng QĐXB để tách lại ngày
        self.assertIn("ngày 05/3/2021", colophon_block("Số 79-QĐ/NXBCTQG ngày 05/3/2021\nISBN 978\nIn xong và nộp lưu chiểu năm 2021."))

    def test_lay_dong_cuoi(self):
        t = "In xong và nộp lưu chiểu tháng 1 năm 2019\n(tái bản)\nIn xong và nộp lưu chiểu tháng 9 năm 2021"
        self.assertEqual(ngay(t), (9, 2021))


class TrangCuoi(unittest.TestCase):
    # PDF 10 trang, dòng In xong ở trang 8: OCR lần lượt 10 -> 9 -> 8 rồi dừng, không render trang 7, 6
    def chay(self, trang_ngay, so_trang=10):
        from unittest import mock
        import numpy as np
        from app import flow_sort_date as fd
        da, ocr_goi = [], []
        self.ocr_goi = ocr_goi

        def render(pdf, out, no, job_id=""):
            da.append(no)
            return f"p{no}"

        def ocr(img, max_tokens=0, label="", **k):
            no = int(label.split()[-1])
            ocr_goi.append(label)
            return (f"Nội dung trang {no}\nISBN 978-604\nIn xong và nộp lưu chiểu tháng 4 năm 2019"
                    if no == trang_ngay else f"Mục lục trang {no}")
        ctx = {"job_id": "t", "job_folder": "x", "pdf_path": "a.pdf", "date_source": "ocr",
               "last_pages": fd.last_page_numbers(so_trang, fd.DATE_LAST_PAGES)}
        with mock.patch.object(fd, "render_page", side_effect=render), \
                mock.patch.object(fd.cv2, "imread", return_value=np.zeros((2, 2, 3), np.uint8)), \
                mock.patch.object(fd, "chandra_ocr", side_effect=ocr):
            fd.date_ocr(ctx)
            r = fd.date_result(dict(ctx, date_source="word"))  # không gọi spellfix
        self.metadata = r["metadata"]
        return da, r["luu_chieu_thang"], r["luu_chieu_nam"], r["metadata_page"]

    def test_toi_da_5_trang_tu_cuoi_len(self):
        # mỗi trang Chandra OCR cả trang đúng 1 lần
        self.assertEqual(self.chay(8), ([10, 9, 8], 4, 2019, 8))
        self.assertEqual(self.ocr_goi, ["trang 10", "trang 9", "trang 8"])
        self.assertEqual(self.metadata, "ISBN 978-604\nIn xong và nộp lưu chiểu tháng 4 năm 2019")  # chỉ giữ khối thông tin XB
        self.assertEqual(self.chay(10), ([10], 4, 2019, 10))
        self.assertEqual(self.chay(6), ([10, 9, 8, 7, 6], 4, 2019, 6))
        self.assertEqual(self.chay(5)[:3], ([10, 9, 8, 7, 6], None, None))  # trang 5 ngoài 5 trang cuối
        self.assertEqual(self.chay(1, so_trang=2)[:3], ([2, 1], 4, 2019))

    def test_trang_chi_co_nam_xet_tiep(self):
        # trang cuối In xong chỉ có năm -> xét tiếp trang trước; trang 8 có ngày QĐXB -> lấy tháng
        from unittest import mock
        import numpy as np
        from app import flow_sort_date as fd
        noi_dung = {10: "In xong và nộp lưu chiểu năm 2019", 9: "Mục lục", 8: "Quyết định xuất bản số 5-QĐ/NXB ngày 02/12/2018"}
        da = []
        ctx = {"job_id": "t", "job_folder": "x", "pdf_path": "a.pdf", "date_source": "ocr", "last_pages": [10, 9, 8, 7, 6]}
        with mock.patch.object(fd, "render_page", side_effect=lambda pdf, out, no, job_id="": da.append(no) or f"p{no}"), \
                mock.patch.object(fd.cv2, "imread", return_value=np.zeros((2, 2, 3), np.uint8)), \
                mock.patch.object(fd, "chandra_ocr", side_effect=lambda img, label="", **k: noi_dung.get(int(label.split()[-1]), "")):
            fd.date_ocr(ctx)
            r = fd.date_result(dict(ctx, date_source="word"))
        self.assertEqual((da, r["luu_chieu_thang"], r["luu_chieu_nam"], r["luu_chieu_theo"]), ([10, 9, 8], 12, 2018, "qdxb"))
        # không trang nào có tháng: dùng năm của trang In xong đã giữ tạm
        noi_dung.pop(8)
        da.clear()
        with mock.patch.object(fd, "render_page", side_effect=lambda pdf, out, no, job_id="": da.append(no) or f"p{no}"), \
                mock.patch.object(fd.cv2, "imread", return_value=np.zeros((2, 2, 3), np.uint8)), \
                mock.patch.object(fd, "chandra_ocr", side_effect=lambda img, label="", **k: noi_dung.get(int(label.split()[-1]), "")):
            fd.date_ocr(ctx)
            r = fd.date_result(dict(ctx, date_source="word"))
        self.assertEqual((da, r["luu_chieu_thang"], r["luu_chieu_nam"], r["metadata_page"]), ([10, 9, 8, 7, 6], None, 2019, 10))


if __name__ == "__main__":
    unittest.main()
