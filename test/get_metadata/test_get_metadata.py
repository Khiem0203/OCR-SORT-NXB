# Test quét thông tin trang cuối trên cả thư mục sách (mặc định sach_da_luu/), gọi API thật, xuất Excel giống tab 6 (--loai metadata: tab 5)
import argparse
import os
import re
import signal
import sys
import threading
import time
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
# Tìm thư mục exe chứa sxb_ocr.py: biến SXB_EXE_DIR, <ROOT>/exe, ~/exe
_EXE = [os.getenv("SXB_EXE_DIR", ""), os.path.join(ROOT, "exe"), os.path.join(os.path.dirname(ROOT), "exe"),
        os.path.expanduser("~/exe")]
_EXE = next((p for p in _EXE if p and os.path.isfile(os.path.join(p, "sxb_ocr.py"))), "")
if not _EXE:
    sys.exit("Không thấy sxb_ocr.py. Copy thư mục exe vào OCR_SORT_NXB/exe hoặc đặt SXB_EXE_DIR=/đường/dẫn/exe")
sys.path.insert(0, _EXE)

import sxb_ocr  # dùng đúng code của tool: liệt kê sách, gọi API, ghi Excel

_NAM_RE = re.compile(r"^(19|20)\d\d$")
_FILE_BIA_RE = re.compile(r"_B_[A-Z0-9]+_[MT]\d\d(_\d+)?$")  # tên bìa do tab 4 đặt: …_B_02_M00, …_B_02_M00_2


# File bìa tab 4 đã chép vào <sách>/Bia/ (không phải một cuốn)
def la_bia_da_luu(path):
    return (os.path.basename(os.path.dirname(path)).lower() == "bia"
            or bool(_FILE_BIA_RE.search(os.path.splitext(os.path.basename(path))[0])))
_THANG_RE = re.compile(r"^thang\s*(\d{1,2})$", re.I)


# Năm / tháng theo thư mục lưu (…/<Năm>/thang <MM>/<tên sách>/file)
def ngay_theo_thu_muc(path, goc):
    phan = os.path.relpath(os.path.dirname(path), goc).replace("\\", "/").split("/")
    nam = thang = None
    for i, p in enumerate(phan):
        if _NAM_RE.match(p):
            nam = int(p)
            m = _THANG_RE.match(phan[i + 1]) if i + 1 < len(phan) else None
            thang = int(m.group(1)) if m else None
    return nam, thang


# Sheet đối chiếu tháng / năm OCR với thư mục đã lưu
def them_doi_chieu(xlsx, dong, goc):
    from openpyxl import load_workbook
    from openpyxl.styles import Font
    wb = load_workbook(xlsx)
    ws = wb.create_sheet("Đối chiếu thư mục")
    ws.append(["STT", "Năm (thư mục)", "Tháng (thư mục)", "Năm (OCR)", "Tháng (OCR)", "Kết quả", "Trạng thái OCR", "File"])
    for c in ws[1]:
        c.font = Font(bold=True)
    dem = {}
    for d in dong:
        nam, thang = ngay_theo_thu_muc(d["pdf"] or d["word"], goc)
        if nam is None:
            kq = "KHÔNG RÕ THƯ MỤC"
        elif d.get("nam") is None:
            kq = "OCR KHÔNG RA"
        elif d["nam"] != nam:
            kq = "KHÁC NĂM"
        elif not thang or d.get("thang") in (None, 0):
            kq = "KHỚP NĂM"
        else:
            kq = "KHỚP" if d["thang"] == thang else "KHÁC THÁNG"
        dem[kq] = dem.get(kq, 0) + 1
        ws.append([d["stt"], nam, thang, d.get("nam"), d.get("thang"), kq, d["trang_thai"], d["pdf"] or d["word"]])
    for col, w in zip("ABCDEFGH", (6, 12, 12, 10, 10, 18, 16, 90)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(xlsx)
    return dem


def main():
    ap = argparse.ArgumentParser(description="Quét metadata cả thư mục sách qua API, xuất Excel như tool")
    ap.add_argument("input", nargs="?", default=os.path.join(ROOT, "sach_da_luu"),
                    help="Thư mục sách (mặc định sach_da_luu/)")
    ap.add_argument("--loai", choices=["metadata", "info", "ca_hai"], default="info",
                    help="info = tab 6 thông tin trang cuối (mặc định), metadata = tab 5 tháng / năm lưu chiểu")
    ap.add_argument("--api", default=sxb_ocr.API_MAC_DINH, help=f"Địa chỉ API (mặc định {sxb_ocr.API_MAC_DINH})")
    ap.add_argument("--song-song", type=int, default=2, help="Số cuốn gửi cùng lúc (mặc định 2)")
    ap.add_argument("--timeout", type=int, default=sxb_ocr.TIMEOUT_MAC_DINH, help="Chờ tối đa mỗi cuốn (giây)")
    ap.add_argument("--gioi-han", type=int, default=0, help="Chỉ chạy N cuốn đầu (0 = tất cả)")
    ap.add_argument("--ra", default=os.path.join(HERE, "ket_qua"), help="Thư mục lưu Excel")
    args = ap.parse_args()

    goc = os.path.abspath(args.input)
    if not os.path.isdir(goc):
        sys.exit(f"Không tìm thấy thư mục: {goc}")
    try:
        print(f"Máy chủ: {sxb_ocr.kiem_tra_ket_noi(args.api)}")
    except Exception as e:
        sys.exit(f"Không kết nối được {args.api}: {e}")

    bo_qua = []
    sach = sxb_ocr.liet_ke_sach(goc, bo_qua=bo_qua)
    bia = [s for s in sach if la_bia_da_luu(s[0] or s[1])]
    sach = [s for s in sach if not la_bia_da_luu(s[0] or s[1])]
    print(f"Thư mục {goc}: {len(sach)} cuốn, bỏ qua {len(bo_qua)} file (bìa / jpg / PDF '- số')"
          f" và {len(bia)} file trong thư mục Bia/")
    if args.gioi_han:
        sach = sach[:args.gioi_han]
    if not sach:
        sys.exit("Không có cuốn nào (PDF / Word)")
    os.makedirs(args.ra, exist_ok=True)

    dung = threading.Event()
    signal.signal(signal.SIGINT, lambda *a: (dung.set(), print("\nĐang dừng: chờ các cuốn đang gửi xong...")))
    tem = f"{datetime.now():%Y%m%d_%H%M%S}"
    loi = 0
    for loai in (["metadata", "info"] if args.loai == "ca_hai" else [args.loai]):
        xlsx = os.path.join(args.ra, f"{sxb_ocr.LOAI[loai][2]}_{tem}.xlsx")
        print(f"\n=== {sxb_ocr.LOAI[loai][1]}: {len(sach)} cuốn, song song {args.song_song} → {xlsx}")
        t0 = time.time()
        dong = sxb_ocr.quet_api(args.api, loai, sach, xlsx, args.song_song, args.timeout, print, dung)
        dem = {}
        for d in dong:
            dem[d["trang_thai"]] = dem.get(d["trang_thai"], 0) + 1
        loi += dem.get("LỖI", 0)
        print(f"Xong {len(dong)}/{len(sach)} cuốn sau {time.time() - t0:.0f}s: "
              + ", ".join(f"{k} {v}" for k, v in sorted(dem.items())))
        if loai == "metadata" and dong:
            dc = them_doi_chieu(xlsx, dong, goc)
            print("Đối chiếu với thư mục: " + ", ".join(f"{k} {v}" for k, v in sorted(dc.items())))
        print(f"Excel: {xlsx}")
        if dung.is_set():
            break
    sys.exit(1 if loi else 0)


if __name__ == "__main__":
    main()
