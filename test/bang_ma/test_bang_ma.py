# Test đổi phông cũ (TCVN3, VNI-Windows) sang Unicode ở cả 3 luồng (không cần máy chủ)
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from app import flow_sort_date, flow_sort_metadata, flow_sort_name

LUONG = (flow_sort_name, flow_sort_date, flow_sort_metadata)

VNI = ("HÖÔÙNG DAÃN CUÛA BAN TOÅ CHÖÙC TRUNG ÖÔNG\n"
       "Soá 07-HD/BTCTW ngaøy 11 thaùng 10 naêm 2011 Veà ñaùnh giaù chaát löôïng toå chöùc cô sôû ñaûng vaø ñaûng vieân\n"
       "Thöïc hieän Nghò quyeát Ñaïi hoäi XI, Chæ thò soá 03-CT/TW, tieáp tuïc ñaåy maïnh, taám göông, khaéc phuïc, yù kieán\n"
       "In xong vaø noäp löu chieåu thaùng 3 naêm 2012.")
UNICODE = ("HƯỚNG DẪN CỦA BAN TỔ CHỨC TRUNG ƯƠNG\n"
           "Số 07-HD/BTCTW ngày 11 tháng 10 năm 2011 Về đánh giá chất lượng tổ chức cơ sở đảng và đảng viên\n"
           "Thực hiện Nghị quyết Đại hội XI, Chỉ thị số 03-CT/TW, tiếp tục đẩy mạnh, tấm gương, khắc phục, ý kiến\n"
           "In xong và nộp lưu chiểu tháng 3 năm 2012.")
TCVN3 = "Nhµ xuÊt b¶n ChÝnh trÞ quèc gia, n¨m 2011, ®¶ng viªn, h­íng dÉn"


class BangMa(unittest.TestCase):
    def test_vni(self):
        for m in LUONG:
            self.assertEqual(m.fix_tcvn3_text(VNI), UNICODE, m.__name__)
        self.assertEqual(flow_sort_date.extract_luu_chieu(flow_sort_date.fix_tcvn3_text(VNI)), {
            "thang": 3, "nam": 2012, "cau": "In xong và nộp lưu chiểu tháng 3 năm 2012.", "theo": "in_xong"})

    def test_vni_chu_hoa_va_mu(self):
        f = flow_sort_name.convert_vni_text
        self.assertEqual(f("AÁn ÑOÄ, ÔÛ, Öu tieân, aêm, aéc, YÙ, Î, ó, ò, æ"), "Ấn ĐỘ, Ở, Ưu tiên, ăm, ắc, Ý, Ỵ, ĩ, ị, ỉ")

    def test_tcvn3_van_dung(self):
        for m in LUONG:
            self.assertEqual(m.fix_tcvn3_text(TCVN3), "Nhà xuất bản Chính trị quốc gia, năm 2011, đảng viên, hướng dẫn")

    def test_unicode_giu_nguyen(self):
        # text Unicode thật (kể cả dòng chỉ có chữ Latin-1 như 'Tôi là') không bị đổi nhầm sang VNI
        for t in (UNICODE, "Tôi là sinh viên. Bà ơi, cô ấy về rồi.\nTôi đi học", "Hello world ISBN 978"):
            for m in LUONG:
                self.assertEqual(m.fix_tcvn3_text(t), t)

    def test_ba_luong_giong_nhau(self):
        self.assertEqual(flow_sort_name.VNI_MAP, flow_sort_date.VNI_MAP)
        self.assertEqual(flow_sort_name.VNI_MAP, flow_sort_metadata.VNI_MAP)


if __name__ == "__main__":
    unittest.main()
