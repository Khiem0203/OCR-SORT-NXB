# OCR-SORT-NXB: lọc và sắp xếp kho sách số cho NXB Chính trị quốc gia Sự thật

Bộ công cụ phục vụ nhiệm vụ **lọc sách** cho **Nhà xuất bản Chính trị quốc gia Sự thật**: từ một kho file số hoá lộn xộn (hàng chục nghìn file Word, PDF, ảnh bìa nằm rải rác trên nhiều ổ đĩa, nhiều bản trùng, đặt tên tuỳ ý) và **danh mục sách xuất bản theo năm** (Excel), tìm đúng file của từng cuốn sách, xác minh bằng OCR / AI, rồi chép vào kho chuẩn:

```
<thư mục đích>/<Năm>/thang <MM>/<ten_sach_mmyy_R_BBT_T01>/
    ten_sach_mmyy_R_BBT_T01.pdf      ← ruột sách
    ten_sach_mmyy_R_BBT_T01.docx
    Bia/ten_sach_mmyy_B_BBT_T01.jpg  ← bìa
```

Sách nào chưa chắc chắn (tên không khớp, không đọc được ngày lưu chiểu, năm nằm ngoài khoảng cần chuẩn hoá…) **không bị chép nhầm** mà được tách riêng vào `CAN_KIEM_TRA/` kèm Excel và ghi chú để người xem lại.

## Thành phần

| Phần | Chạy ở đâu | Việc làm |
|---|---|---|
| **Tool SXB** (`exe/`) | Máy có giao diện (Windows hoặc Linux), đóng gói thành 1 file chạy | Quét kho, đối chiếu với danh mục Excel, gọi API OCR, chép và đổi tên file, xuất Excel |
| **Service OCR SORT BOOK** (`app/`) | Máy chủ GPU (Linux) | API FastAPI cổng 9091: đọc Word / OCR PDF, tìm **tên sách**, **ngày lưu chiểu**, **thông tin xuất bản** ở trang cuối |
| **Model** | Máy chủ GPU | Chandra OCR 2 (vLLM, OCR trang), Qwen3-4B-Instruct (vLLM, tìm tên sách trong chữ OCR), ProtonX (sửa chính tả tiếng Việt) |

```
Danh mục Excel ──┐
                 ▼
Kho file ──▶ Tool SXB: 1. Quét kho → 2. Đối chiếu tên sách ↔ file → 3. Kiểm tra nội dung (năm)
                 │
                 ▼  4. OCR & copy (gọi API)
          Service OCR SORT BOOK ──▶ Chandra OCR / Qwen / ProtonX
                 │  tên sách + tháng / năm lưu chiểu
                 ▼
          Kho chuẩn Năm/thang MM/<tên mới>/   +   CAN_KIEM_TRA/<loại>/   +   Excel kết quả
```

## Quy trình lọc sách (Tool SXB)

| Tab | Việc |
|---|---|
| **1. Quét kho** | Liệt kê mọi file trong kho (tên, định dạng, ngày sửa, đường dẫn) ra Excel; dừng giữa chừng thì lần sau quét tiếp |
| **2. Đối chiếu Excel** | Khớp **tên sách** trong danh mục với **tên file / tên thư mục** (bỏ dấu, viết thường, trùng hẳn hoặc thư mục tên rút gọn), lọc theo **năm** → `KHỚP` / `LỆCH NĂM` / `KHÔNG TÌM THẤY`; chọn file Word / PDF ruột, bỏ file bìa |
| **3. Kiểm tra nội dung** | Mở file đọc năm xuất bản ở trang cuối / 8 trang đầu, tìm "cuốn sách <tên>" trong lời nói đầu cho sách chưa tìm thấy; PDF scan được đánh dấu cần OCR |
| **4. OCR & copy** | Gọi API OCR lấy **tên sách** và **ngày lưu chiểu**, chỉ chép sách có tên OCR **khớp** tên trong danh mục, xếp theo thời gian OCR, đổi tên theo quy tắc, xuất Excel; có **Hoàn tác** |
| **5. Cần kiểm tra** | Chép đủ file liên quan của các sách cần xem vào `CAN_KIEM_TRA/<loại>/` |
| **6. Quét metadata**, **7. Thông tin trang cuối** | Quét riêng ngày lưu chiểu / các trường thông tin xuất bản (NXB, biên tập, ISBN, giấy phép…) ra Excel |

### Quy tắc chính

- **Tên sách**: viết thường, bỏ dấu, bỏ ký tự đặc biệt và chú thích ("Sách chuyên khảo", "Tái bản"…) rồi so **bằng nhau** → `KHỚP`, còn lại `KHÁC`. Tên Word / PDF không khớp thì OCR thêm **file bìa**; vẫn không khớp thì `CẦN XEM` (không chép).
- **Ngày lưu chiểu** (chỉ theo OCR, **không dùng ngày trong danh mục**):
  1. dòng **"In xong và nộp lưu chiểu tháng M năm YYYY"** (nhận nhiều cách viết: `tháng 3-2021`, `T3/2021`, `tháng ba năm 2021`…), tìm từ trang cuối lên, tối đa 5 trang;
  2. dòng In xong không có tháng → ngày **Quyết định xuất bản** (`Số 79-QĐ/NXBCTQG ngày 05/3/2021`);
  3. chỉ có năm → `thang 00`; không có gì → `CAN_KIEM_TRA/KHONG_NGAY`.
- **Năm trước 2006** (ngoài khoảng đang chuẩn hoá, đổi bằng `SXB_NAM_BAT_DAU`) → `CAN_KIEM_TRA/TRUOC_2006`.
- **Bìa**: file tên có chữ "Bìa", file `.jpg`, PDF tên kết thúc "- số" (vd `... - 01.pdf`); chép vào thư mục con `Bia/`.
- **Tên file mới**: `ten_sach_mmyy_R_<mã BBT>_<lần XB>` (mã BBT lấy từ `BBT_distinct.xlsx`, lần XB: `M00` in mới, `T01`… tái bản).
- **Phông cũ** TCVN3 / VNI trong file Word tự đổi sang Unicode; Word lỗi bảng mã thì chuyển sang OCR PDF.

Chi tiết từng tab: [exe/README.md](exe/README.md).

## Cài đặt

### Máy chủ (service + model)

Yêu cầu: Linux có GPU NVIDIA, Python 3.12 (conda env `ocr`), vLLM, poppler (`pdftoppm`), `antiword` (đọc `.doc`), `ffmpeg`.

```bash
git clone https://github.com/Khiem0203/OCR-SORT-NXB.git
cd OCR-SORT-NXB
conda create -n ocr python=3.12 && conda activate ocr
pip install -r requirements.txt        # torch cài riêng theo bản CUDA của máy
sudo apt install -y poppler-utils antiword ffmpeg
cp .env.example .env                   # sửa đường dẫn model, cổng... trong .env
```

Model (tải về `~/.cache/huggingface`, điền đường dẫn snapshot vào `.env`):

| Model | Dùng cho | Cổng |
|---|---|---|
| `datalab-to/chandra-ocr-2` | OCR trang | 8101 (`service/vllm-chandra.service`) |
| `Qwen/Qwen3-4B-Instruct-2507` | Tìm tên sách | 8113 (`service/llm-qwen3.service`) |
| `protonx-models/protonx-legal-tc` | Sửa chính tả | 8006 (bật trong `start-service.sh`) |

Chạy bằng systemd (sửa `your_user`, đường dẫn trong các file `service/*.service` trước):

```bash
sudo cp service/*.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now vllm-chandra llm-qwen3 start-ocr
curl http://localhost:9091/health
```

Hoặc chạy tay: `bash bash_sh/sort/start-service.sh` / `bash bash_sh/sort/stop-service.sh`.

### Tool SXB

```bash
cd exe
bash build.sh          # Linux → exe/dist-linux/SXBTool_OCR
build.bat              # Windows → exe\dist\SXBTool_OCR.exe
python sxb_ocr.py      # hoặc chạy thẳng bằng Python
```

Tool gọi API ở `http://localhost:9091`; máy khác thì đặt biến `SXB_OCR_API=http://your_ip:9091` trước khi mở.

## API (tóm tắt)

| Endpoint | Chức năng |
|---|---|
| `POST /api/sort/title` | Tên sách (gửi PDF / Word / file bìa, hoặc đường dẫn) |
| `POST /api/sort/metadata` | Tháng / năm lưu chiểu ở trang cuối |
| `POST /api/sort/info` | Các trường thông tin xuất bản ở trang cuối |
| `POST /api/sort` | Tất cả các bước trên |
| `GET /api/request-status/{id}`, `GET /api/jobs/status`, `GET /health` | Trạng thái job, sức khoẻ service |

```bash
curl -F "pdf=@sach.pdf" -F "word=@sach.docx" http://your_ip:9091/api/sort/title
```

Chi tiết luồng xử lý, tham số, log: [docs/KY_THUAT.md](docs/KY_THUAT.md).

## Kiểm thử

```bash
python test/get_date/test_luu_chieu.py   # quy tắc tách ngày lưu chiểu
python test/bang_ma/test_bang_ma.py      # đổi phông TCVN3 / VNI
python test/api/test_api.py              # API, hàng đợi job (giả lập model)
python test/api/test_bia.py              # luồng OCR bìa
cd exe && python sxb_ocr.py --test       # tự kiểm tra của tool (cần pymupdf, openpyxl)
```

Các script `test/get_*/test_get_*.py` chạy thật trên thư mục sách (cần model đang chạy); dữ liệu mẫu đặt trong `test/data/` (không đưa lên git, xem [test/data/README.md](test/data/README.md)).

## Cấu trúc thư mục

```
app/            service FastAPI: main.py (API), jobs.py (hàng đợi), flow_sort_*.py (mỗi luồng tự chứa), spellfix_api.py
exe/            Tool SXB: sxb_tool.py (tab 1-3), sxb_ocr.py (tab 4-7), build.sh / build.bat
bash_sh/sort/   script bật / tắt service
service/        file systemd (vLLM Chandra, vLLM Qwen, service OCR)
test/           kiểm thử
docs/           tài liệu kỹ thuật
```

## Bảo mật

- Không đưa `.env`, dữ liệu sách, kết quả quét, log lên git (đã có trong `.gitignore`).
- Mọi địa chỉ máy chủ, tài khoản trong repo là giá trị mẫu (`your_ip`, `your_user`…); điền giá trị thật vào `.env` / file service trên máy chạy.
