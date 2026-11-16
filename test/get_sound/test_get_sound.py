# Test luồng metadata âm thanh (phần không dùng AI) trên cả thư mục
import argparse
import csv
import json
import os
import shutil
import sys
import time
from collections import defaultdict
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
DATA_DIR = os.path.join(ROOT, "test", "data")
START_CWD = os.getcwd()
sys.path.insert(0, ROOT)
os.chdir(ROOT)

import logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

from app import flow_sort_sound as fs

CSV_COLS = ([("stt", "STT"), ("thu_muc", "Thư mục"), ("trang_thai", "Trạng thái"), ("nguon", "Nguồn kỹ thuật"),
             ("so_truong", "Số trường")]
            + [(k, fs.TEN_TRUONG[k]) for k in fs.FIELDS]
            + [("thoi_luong_giay", "Thời lượng (giây)"), ("dung_luong_byte", "Dung lượng (byte)"),
               ("loi", "Lỗi"), ("giay", "Thời gian xử lý (s)"), ("file", "File")])


# Tìm thư mục input
def _resolve_folder(path: str) -> str:
    for p in ([path] if os.path.isabs(path) else [os.path.join(b, path) for b in (START_CWD, ROOT, DATA_DIR)]):
        if os.path.isdir(p):
            return os.path.abspath(p)
    sys.exit(f"Không tìm thấy thư mục: {path} (input phải là thư mục)")


# Liệt kê file âm thanh trong thư mục
def find_audio(folder: str) -> list:
    files = []
    for cur, dirs, names in os.walk(folder):
        dirs.sort()
        dirs[:] = [d for d in dirs if not d.startswith("ket_qua_")]
        files += [os.path.join(cur, n) for n in sorted(names) if n.lower().endswith(fs.AUDIO_EXT)]
    return files


# Tạo file mẫu có đủ thẻ (UTF-8, Latin-1 lỗi font, ảnh bìa) từ một mp3 có sẵn
def tao_mau(goc: str, out_dir: str) -> list:
    from mutagen.id3 import ID3, TIT2, TPE1, TPE2, TCOM, TALB, TRCK, TCON, TCOP, TLAN, TDRC, COMM, TMCL, APIC, Encoding
    os.makedirs(out_dir, exist_ok=True)
    png = bytes.fromhex("89504e470d0a1a0a0000000d4948445200000001000000010806000000"
                        "1f15c4890000000d49444154789c6360000002000154a24f5d0000000049454e44ae426082")
    mau = []
    for ten, ma in (("mau_utf8.mp3", Encoding.UTF8), ("mau_latin1_loi_font.mp3", Encoding.LATIN1)):
        p = os.path.join(out_dir, ten)
        shutil.copyfile(goc, p)
        chu = (lambda s: s) if ma == Encoding.UTF8 else (lambda s: s.encode("utf-8").decode("latin-1"))
        t = ID3()
        t.add(TIT2(encoding=ma, text=chu("Chương 3 - Tuổi thơ")))
        t.add(TPE1(encoding=ma, text=chu("Nguyễn Văn A")))
        t.add(TPE2(encoding=ma, text=chu("Nguyễn Văn A")))
        t.add(TCOM(encoding=ma, text=chu("Tô Hoài")))
        t.add(TMCL(encoding=Encoding.UTF8, people=[["Người đọc", "Trần Thị B"]]))
        t.add(TALB(encoding=ma, text=chu("Dế Mèn phiêu lưu ký")))
        t.add(TRCK(encoding=ma, text="3/12"))
        t.add(TCON(encoding=ma, text="Audiobook"))
        t.add(TCOP(encoding=ma, text=chu("NXB Kim Đồng")))
        t.add(TLAN(encoding=ma, text="vie"))
        t.add(TDRC(encoding=ma, text="2023-05-12T09:30:00"))
        t.add(COMM(encoding=ma, lang="vie", desc="", text=chu("Sách nói thử nghiệm")))
        t.add(APIC(encoding=Encoding.UTF8, mime="image/png", type=3, desc="Bìa", data=png))
        t.save(p)
        mau.append(p)
    print(f"Tạo {len(mau)} file mẫu có thẻ đầy đủ trong {out_dir}")
    return mau


# Chạy luồng cho một file
def run(idx: int, path: str, out_dir: str, args) -> dict:
    stem = os.path.splitext(os.path.basename(path))[0]
    file_dir = os.path.join(out_dir, f"{idx:03d}_{stem[:60]}")
    job_id = f"test_sound_{os.path.basename(out_dir)}_{idx:03d}"
    row = {"stt": idx, "thu_muc": os.path.relpath(os.path.dirname(path), args.input), "file": path}
    t0 = time.time()
    try:
        os.makedirs(file_dir, exist_ok=True)
        context = {"job_id": job_id, "job_folder": file_dir, "audio_path": path}
        fs.sound_prepare(context)
        r = fs.sound_result(context)
        kt = r["am_thanh_ky_thuat"]
        n = fs.so_truong(r["am_thanh"])
        row.update(r["am_thanh"], trang_thai="LỖI" if r["am_thanh_error"] else "OK", nguon=kt.get("nguon", ""),
                   so_truong=n, thoi_luong_giay=kt.get("thoi_luong_giay"), dung_luong_byte=kt.get("dung_luong_byte"),
                   loi=r["am_thanh_error"])
        with open(os.path.join(file_dir, "the.json"), "w", encoding="utf-8") as f:
            json.dump({"ky_thuat": kt, "the": r["am_thanh_the"]}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logging.exception("Lỗi")
        row.update(trang_thai="LỖI", loi=str(e))
    row["giay"] = round(time.time() - t0, 2)
    with open(os.path.join(file_dir, "result.json"), "w", encoding="utf-8") as f:
        json.dump(row, f, ensure_ascii=False, indent=2)
    return row


# Chạy test cả thư mục
def main():
    ap = argparse.ArgumentParser(description="Test luồng metadata âm thanh (ffprobe + thẻ, không AI) trên cả thư mục")
    ap.add_argument("input", nargs="?", default=os.path.join(DATA_DIR, "audio_1"),
                    help="Thư mục âm thanh (mặc định test/data/audio_1)")
    ap.add_argument("--no-ffprobe", action="store_true", help="Không dùng ffprobe, chỉ đọc bằng mutagen")
    ap.add_argument("--mau", action="store_true",
                    help="Tạo thêm 2 file mẫu có đủ thẻ (UTF-8 và Latin-1 lỗi font, có ảnh bìa) từ mp3 đầu tiên")
    ap.add_argument("--limit", type=int, default=0, help="Chỉ chạy N file đầu")
    args = ap.parse_args()
    args.input = _resolve_folder(args.input)
    if args.no_ffprobe:
        fs.FFPROBE_BIN = "__khong_dung_ffprobe__"
    elif not shutil.which(fs.FFPROBE_BIN):
        print(f"CHÚ Ý: không thấy {fs.FFPROBE_BIN}, thông số kỹ thuật đọc bằng mutagen")

    files = find_audio(args.input)
    if not files:
        sys.exit(f"Không có file âm thanh ({', '.join(fs.AUDIO_EXT)}) nào trong {args.input}")
    if args.limit:
        files = files[:args.limit]
    out_dir = os.path.join(HERE, f"ket_qua_{datetime.now():%Y%m%d_%H%M%S}")
    os.makedirs(out_dir, exist_ok=True)
    if args.mau:
        goc = next((p for p in files if p.lower().endswith(".mp3")), None)
        files += tao_mau(goc, os.path.join(out_dir, "mau")) if goc else []
    print(f"Input : {args.input}  ({len(files)} file)\nOutput: {out_dir}")

    rows = []
    for i, path in enumerate(files, 1):
        print(f"\n######## [{i}/{len(files)}] {path}")
        row = run(i, path, out_dir, args)
        print(f"=> {row['trang_thai']}  nguồn={row.get('nguon') or '-'}  {row.get('so_truong', 0)}/{len(fs.FIELDS)} trường"
              + (f"  lỗi: {row['loi']}" if row.get("loi") else ""))
        for k in fs.FIELDS:
            if row.get(k):
                print(f"   {fs.TEN_TRUONG[k]:34s} {row[k]}")
        rows.append(row)

    with open(os.path.join(out_dir, "summary.json"), "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)
    with open(os.path.join(out_dir, "summary.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([h for _, h in CSV_COLS])
        for r in rows:
            w.writerow(["" if r.get(k) is None else r.get(k) for k, _ in CSV_COLS])

    # Gom theo album: số file, tổng thời lượng
    album = defaultdict(lambda: [0, 0.0])
    for r in rows:
        if r.get("album"):
            album[r["album"]][0] += 1
            album[r["album"]][1] += r.get("thoi_luong_giay") or 0
    for ten, (n, giay) in album.items():
        print(f"\nAlbum {ten!r}: {n} file, tổng {fs._thoi_luong(giay)}")

    print("\nSố file có từng trường:")
    for k in fs.FIELDS:
        print(f"   {fs.TEN_TRUONG[k]:34s} {sum(bool(r.get(k)) for r in rows)}/{len(rows)}")
    ok = sum(r["trang_thai"] == "OK" for r in rows)
    print(f"\nTỔNG: {ok}/{len(rows)} file đọc được\nKết quả: {out_dir}")
    if any(r["trang_thai"] == "LỖI" for r in rows):
        sys.exit(1)


if __name__ == "__main__":
    main()
