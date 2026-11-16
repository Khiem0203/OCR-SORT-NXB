# Test luồng tháng, năm lưu chiểu trên cả thư mục
import argparse
import csv
import json
import os
import shutil
import sys
import time
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
DATA_DIR = os.path.join(ROOT, "test", "data")
START_CWD = os.getcwd()
sys.path.insert(0, ROOT)
os.chdir(ROOT)

import logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

from app import flow_sort_date as fd
from app.file_source import (
    to_local_path, skip_reason, _is_pdf, _is_docx, _word_rank, _find_docx_for, WORD_EXT)

CSV_COLS = [("stt", "STT"), ("sach", "Thư mục sách"), ("trang_thai", "Trạng thái"), ("nguon", "Nguồn"),
            ("thang", "Tháng"), ("nam", "Năm"), ("cau", "Dòng In xong"), ("trang", "Trang PDF"),
            ("score", "Điểm vùng cắt"), ("ly_do_bo_word", "Lý do bỏ Word"), ("loi", "Lỗi"),
            ("giay", "Thời gian (s)"), ("pdf", "File PDF"), ("word", "File Word")]


# Tìm thư mục input
def _resolve_folder(path: str) -> str:
    local = to_local_path(path)
    for p in ([local] if os.path.isabs(local) else [os.path.join(b, local) for b in (START_CWD, ROOT, DATA_DIR)]):
        if os.path.isdir(p):
            return os.path.abspath(p)
    sys.exit(f"Không tìm thấy thư mục: {path} (input phải là thư mục)")


# Liệt kê các cuốn trong thư mục
def find_books(folder: str) -> list:
    books = []
    for cur, dirs, files in os.walk(folder):
        dirs.sort()
        names = sorted(files)
        for n in names:
            if n.lower().endswith((".pdf", ".jpg", ".jpeg") + WORD_EXT) and skip_reason(n):
                print(f"Bỏ qua: {os.path.join(cur, n)} ({skip_reason(n)})")
        pdfs = [n for n in names if _is_pdf(n)]
        if pdfs:
            books += [(os.path.join(cur, n), _find_docx_for(os.path.join(cur, n))) for n in pdfs]
            continue
        seen = set()
        for n in sorted((n for n in names if _is_docx(n)), key=_word_rank):
            if os.path.splitext(n)[0].lower() not in seen:
                seen.add(os.path.splitext(n)[0].lower())
                books.append(("", os.path.join(cur, n)))
    return books


# Chạy luồng ngày cho một cuốn
def run(idx: int, pdf: str, word: str, out_dir: str, args) -> dict:
    src = pdf or word
    stem = os.path.splitext(os.path.basename(src))[0]
    book_dir = os.path.join(out_dir, f"{idx:03d}_{stem[:60]}")
    job_id = f"test_date_{os.path.basename(out_dir)}_{idx:03d}"
    word = "" if args.force_ocr else word
    print(f"Job id: {job_id}  (grep -r {job_id} logs/)\nPDF   : {pdf or '-'}\nWord  : {word or '-'}")
    row = {"stt": idx, "sach": os.path.relpath(os.path.dirname(src), args.input), "pdf": pdf, "word": word}
    t0 = time.time()
    try:
        os.makedirs(book_dir, exist_ok=True)
        context = {"job_id": job_id, "job_folder": book_dir, "pdf_path": pdf, "docx_path": word}
        fd.date_prepare(context)
        fd.date_ocr(context)
        r = fd.date_result(context)
        row.update(trang_thai="OK" if (r["luu_chieu_thang"] or r["luu_chieu_nam"]) else "KHÔNG RA NGÀY",
                   nguon=r["luu_chieu_source"], thang=r["luu_chieu_thang"], nam=r["luu_chieu_nam"],
                   cau=r["luu_chieu_cau"], trang=r["metadata_page"], score=r["metadata_score"],
                   ly_do_bo_word=r["date_word_reason"], loi=r["metadata_error"], metadata=r["metadata"],
                   metadata_raw=r["metadata_raw"], box=r["metadata_box"], page_count=context.get("page_count"))
    except Exception as e:
        logging.exception("Lỗi")
        row.update(trang_thai="LỖI", loi=str(e))
    row["giay"] = round(time.time() - t0, 2)
    with open(os.path.join(book_dir, "result.json"), "w", encoding="utf-8") as f:
        json.dump(row, f, ensure_ascii=False, indent=2)
    if args.no_images:
        for sub in ("pages",):
            shutil.rmtree(os.path.join(book_dir, sub), ignore_errors=True)
    return row


# Chạy test cả thư mục
def main():
    ap = argparse.ArgumentParser(description="Test luồng ngày lưu chiểu trên cả một thư mục sách")
    ap.add_argument("input", nargs="?", default=DATA_DIR, help="Thư mục sách (mặc định test/data/)")
    ap.add_argument("--force-ocr", action="store_true", help="Bỏ qua Word, luôn Chandra OCR trang cuối PDF")
    ap.add_argument("--pages", type=int, default=fd.DATE_LAST_PAGES,
                    help=f"Số trang cuối tìm ngày (mặc định {fd.DATE_LAST_PAGES})")
    ap.add_argument("--no-images", action="store_true", help="Không giữ ảnh trang cuối trong output")
    args = ap.parse_args()
    args.input = _resolve_folder(args.input)
    fd.DATE_LAST_PAGES = max(1, args.pages)

    books = find_books(args.input)
    if not books:
        sys.exit(f"Không có sách (PDF / Word) nào trong {args.input}")
    out_dir = os.path.join(HERE, f"ket_qua_{datetime.now():%Y%m%d_%H%M%S}")
    os.makedirs(out_dir, exist_ok=True)
    print(f"Input : {args.input}  ({len(books)} cuốn)\nOutput: {out_dir}")

    rows = []
    for i, (pdf, word) in enumerate(books, 1):
        print(f"\n######## [{i}/{len(books)}] {pdf or word}")
        row = run(i, pdf, word, out_dir, args)
        print(f"=> {row['trang_thai']}  nguồn={row.get('nguon') or '-'}  tháng/năm={row.get('thang')}/{row.get('nam')}"
              f"  ← {row.get('cau') or row.get('loi') or '-'}")
        rows.append(row)

    with open(os.path.join(out_dir, "summary.json"), "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)
    with open(os.path.join(out_dir, "summary.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([h for _, h in CSV_COLS])
        for r in rows:
            w.writerow(["" if r.get(k) is None else r.get(k) for k, _ in CSV_COLS])

    ok = sum(r["trang_thai"] == "OK" for r in rows)
    print(f"\nTỔNG: {ok}/{len(rows)} cuốn ra tháng/năm\nKết quả: {out_dir}")
    if any(r["trang_thai"] == "LỖI" for r in rows):
        sys.exit(1)


if __name__ == "__main__":
    main()
