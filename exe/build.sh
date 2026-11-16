#!/usr/bin/env bash
# Build SXBTool_OCR cho Ubuntu: dist-linux/SXBTool_OCR = tool SXB (tab 1-3) + OCR & copy, metadata, thông tin trang cuối (gọi API OCR SORT BOOK).
# Tự cài gói hệ thống còn thiếu (cần sudo), tạo môi trường Python riêng (.venv-linux), cài thư viện, tự kiểm tra rồi build.
# Chạy: bash build.sh      (muốn dùng Python cụ thể: PYTHON=/usr/bin/python3.11 bash build.sh)
set -euo pipefail
cd "$(dirname "$0")"
echo "=== Build SXBTool_OCR (Linux, 6 tab) ==="

# 1. Gói hệ thống: python3 + venv + tkinter (giao diện), antiword/catdoc (đọc .doc - Linux không có Word),
#    binutils + libpython3-dev (PyInstaller cần objdump và libpython3.x.so)
can=()
for g in python3 python3-venv python3-tk antiword catdoc binutils libpython3-dev; do
    dpkg -s "$g" >/dev/null 2>&1 || can+=("$g")
done
if ((${#can[@]})); then
    echo "Cài gói còn thiếu: ${can[*]}"
    SUDO=""; [ "$(id -u)" -ne 0 ] && SUDO="sudo"
    $SUDO apt-get update -qq && $SUDO apt-get install -y -qq "${can[@]}"
fi

PY="${PYTHON:-python3}"
"$PY" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)' || { echo "[LỖI] Cần Python 3.9 trở lên"; exit 1; }
"$PY" -c 'import tkinter' || { echo "[LỖI] Python này thiếu tkinter: sudo apt install python3-tk"; exit 1; }

# 2. Môi trường riêng + thư viện (pymupdf ghim dưới 2.0 như bản Windows; không cần pywin32)
[ -x .venv-linux/bin/python ] || "$PY" -m venv .venv-linux
V=.venv-linux/bin/python
"$V" -m pip install -q --disable-pip-version-check --upgrade pip
"$V" -m pip install -q --disable-pip-version-check openpyxl "pymupdf>=1.24,<2" python-docx requests pyinstaller
"$V" -c 'import sys, fitz; print("Python", sys.version.split()[0], "- PyMuPDF", fitz.VersionBind)'

# 3. Tự kiểm tra trước khi build (gồm test của sxb_tool.py, API giả lập)
echo "Tự kiểm tra (sxb_ocr.py --test)..."
"$V" -X utf8 sxb_ocr.py --test

# 4. Build
echo "Build dist-linux/SXBTool_OCR..."
"$V" -m PyInstaller --noconfirm --onefile --windowed --log-level WARN --name SXBTool_OCR \
    --distpath dist-linux --workpath build-linux/work --specpath build-linux sxb_ocr.py

# 5. Kiểm tra lại chính file vừa build (--test không mở giao diện)
dist-linux/SXBTool_OCR --test
echo
echo "=== XONG: $(pwd)/dist-linux/SXBTool_OCR ==="
echo "  Tab 1-4: Quét kho | Đối chiếu Excel | Kiểm tra nội dung | OCR & copy"
echo "  Tab 5-6: Quét metadata | Quét thông tin trang cuối (tab 4-6 gọi API localhost:9091, đổi bằng SXB_OCR_API)"
echo "  Chạy: ./dist-linux/SXBTool_OCR"
