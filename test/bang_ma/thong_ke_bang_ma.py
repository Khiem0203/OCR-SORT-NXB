# Thống kê bảng mã các file Word trong file đối chiếu: đọc được (Unicode / TCVN3) hay lỗi bảng mã, phông nào
import argparse
import csv
import os
import re
import sys
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, ROOT)
os.chdir(ROOT)

from app import flow_sort_name as fn

# Tên phông trong file .doc nằm ở bảng phông dạng UTF-16; lấy các chuỗi trông như tên phông
_PHONG_RE = re.compile(r"[A-Za-z.][A-Za-z0-9 .\-]{2,30}")
_PHONG_CU_RE = re.compile(r"^\.vn|^vni|^vps|^vk |^vn3|^bk|^vw|^vn-|tcvn|abc", re.I)


# Tên phông dùng trong file Word
def phong_chu(path: str) -> list:
    if path.lower().endswith(".docx"):
        try:
            from docx import Document
            doc = Document(path)
            ten = Counter(r.font.name for p in doc.paragraphs[:300] for r in p.runs if r.font.name)
            return [t for t, _ in ten.most_common(5)]
        except Exception:
            return []
    try:
        with open(path, "rb") as f:
            raw = f.read()
    except OSError:
        return []
    ten = Counter(m.group(0).strip() for m in _PHONG_RE.finditer(raw.decode("utf-16-le", errors="ignore")))
    return [t for t, _ in ten.most_common() if _PHONG_CU_RE.search(t)][:5]


# Phân loại một file Word
def phan_loai(path: str) -> dict:
    d = {"file": path, "phong": ", ".join(phong_chu(path))}
    try:
        pages, err = fn._extract_word_pages(path, 3)
    except Exception as e:
        pages, err = [], f"{type(e).__name__}: {e}"
    text = "\n".join(pages)
    d["mau"] = re.sub(r"\s+", " ", text).strip()[:120]
    d["diem"] = round(fn.vn_score(text), 2)
    if err:
        d["loai"], d["ly_do"] = "ĐỌC LỖI", err
    else:
        loi = fn.word_quality_problem(pages)
        d["loai"] = "LỖI BẢNG MÃ" if "bảng mã" in loi else "KÉM" if loi else "OK"
        d["ly_do"] = loi
    return d


# Danh sách file Word trong sheet KHOP (hoặc mọi sheet) của file đối chiếu
def doc_file_word(xlsx: str, chi_khop: bool) -> list:
    import openpyxl
    wb = openpyxl.load_workbook(xlsx, read_only=True)
    ds = []
    for ws in wb.worksheets:
        if chi_khop and ws.title != "KHOP":
            continue
        it = ws.iter_rows(values_only=True)
        hdr = [str(h).strip() if h is not None else "" for h in next(it, ())]
        if "File Word chọn" not in hdr:
            continue
        i = hdr.index("File Word chọn")
        ds += [str(r[i]).strip() for r in it if i < len(r) and r[i]]
    wb.close()
    return list(dict.fromkeys(ds))


def main():
    ap = argparse.ArgumentParser(description="Thống kê bảng mã các file Word trong file đối chiếu")
    ap.add_argument("xlsx", help="File doi_chieu_*.xlsx")
    ap.add_argument("--tat-ca-sheet", action="store_true", help="Lấy cả sheet CAN_XU_LY (mặc định chỉ KHOP)")
    ap.add_argument("--song-song", type=int, default=8, help="Số file đọc cùng lúc (mặc định 8)")
    ap.add_argument("--limit", type=int, default=0, help="Chỉ chạy N file đầu")
    args = ap.parse_args()

    files = doc_file_word(args.xlsx, not args.tat_ca_sheet)
    thieu = [p for p in files if not os.path.isfile(p)]
    files = [p for p in files if os.path.isfile(p)]
    if args.limit:
        files = files[:args.limit]
    print(f"{len(files)} file Word ({len(thieu)} file không còn trên ổ, bỏ qua)")
    t0, rows = time.time(), []
    with ThreadPoolExecutor(max_workers=max(1, args.song_song)) as ex:
        for k, d in enumerate(ex.map(phan_loai, files), 1):
            rows.append(d)
            if k % 100 == 0 or k == len(files):
                print(f"  {k}/{len(files)} file, {time.time() - t0:.0f}s")

    out = os.path.join(HERE, f"thong_ke_bang_ma_{datetime.now():%Y%m%d_%H%M%S}.csv")
    cot = [("loai", "Loại"), ("diem", "Tỷ lệ âm tiết hợp lệ"), ("phong", "Phông (.docx: mọi phông, .doc: phông cũ)"), ("ly_do", "Lý do"),
           ("mau", "Mẫu chữ"), ("file", "File")]
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([h for _, h in cot])
        for d in sorted(rows, key=lambda d: (d["loai"], d["diem"])):
            w.writerow([d.get(k, "") for k, _ in cot])

    print("\nTheo loại:")
    for loai, n in Counter(d["loai"] for d in rows).most_common():
        print(f"   {loai:12s} {n}")
    for ten, loc in (("LỖI BẢNG MÃ", lambda d: d["loai"] == "LỖI BẢNG MÃ"), ("đọc được", lambda d: d["loai"] == "OK")):
        phong = Counter(p.strip() for d in rows if loc(d) for p in d["phong"].split(",") if p.strip())
        print(f"\nPhông hay gặp trong file {ten}:")
        for p, n in phong.most_common(15):
            print(f"   {p:30s} {n}")
    print(f"\nKết quả: {out}")


if __name__ == "__main__":
    main()
