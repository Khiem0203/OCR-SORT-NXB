# OCR SORT BOOK - chi tiết kỹ thuật service

Service OCR sách: tool bên ngoài gọi API, gửi file PDF (và file Word nếu có), hoặc chỉ gửi đường dẫn tới file/thư mục gốc (ví dụ `D:\sach\...`). Service tìm **tên sách** từ các trang đầu hoặc từ file Word, đọc vùng **METADATA** ở trang cuối và trả kết quả JSON ngay trong response. Không dùng S3, không gửi webhook.

Mọi file đầu vào (upload, copy từ thư mục ngoài, hoặc tải từ URL nếu bật `SORT_ALLOW_URL`) đều được đưa vào `temp_processing/<job_id>/` để các luồng/model trên WSL đọc trực tiếp. Xử lý xong thì thư mục này bị xóa; đặt `SORT_KEEP_TEMP=1` nếu muốn giữ lại. Thư mục job bị treo (API tắt giữa chừng) quá `SORT_TEMP_MAX_AGE_HOURS` giờ sẽ tự dọn ở request sau.

## Luồng xử lý

Mỗi luồng là một file độc lập, tự chứa đủ phần gọi Chandra / Spellfix / đọc Word của nó (không có file dùng chung). Job chạy ngay trong tiến trình API (`app/jobs.py`, thread pool `SORT_WORKERS` job cùng lúc), **không cần RabbitMQ, PostgreSQL, Celery**. `mode` quyết định job gồm những luồng nào:

| File | Hàm | Nội dung |
|---|---|---|
| `app/flow_sort_name.py` | `sort_name` | **Tên sách**: đọc Word hoặc render + Chandra OCR 8 trang đầu → Qwen tìm tên; không ra tên thì OCR 8 trang đầu PDF (nếu đang dùng Word), rồi OCR **file bìa** (nếu gửi kèm) → sửa chính tả (xem mục dưới) |
| `app/flow_sort_date.py` | `sort_date` | **Ngày lưu chiểu**: đọc trang cuối Word; không thấy thì Chandra OCR **cả trang**, tối đa 5 trang cuối PDF (từng trang từ cuối lên) → tách tháng / năm từ dòng "In xong và nộp lưu chiểu" |
| `app/flow_sort_metadata.py` | `sort_metadata` | **Thông tin trang cuối**: Word hoặc Chandra OCR cả trang cuối → regex |
| `app/flow_sort_sound.py` | `sort_sound` | **Âm thanh**: ffprobe + thẻ (không AI) |
| `app/tasks_new.py` | `run_job` | Chuẩn bị file, chạy các luồng theo `FLOW_TASKS[mode]`, ghi kết quả, xoá temp |
| `app/jobs.py` | `submit`, `status` | Hàng đợi job trong API: trạng thái `PENDING` / `STARTED` / `SUCCESS` / `FAILURE` giữ trong bộ nhớ `SORT_JOB_KEEP_HOURS` giờ |

| `mode` (endpoint) | Các bước |
|---|---|
| `all` (`/api/sort`) | prepare → name → date → metadata → ghi kết quả |
| `title` (`/api/sort/title`) | prepare → name → ghi kết quả |
| `metadata` (`/api/sort/metadata`) | prepare → date → ghi kết quả |
| `info` (`/api/sort/info`) | prepare → metadata → ghi kết quả |
| `sound` (`/api/sort/sound`) | prepare → sound → ghi kết quả |

```
POST /api/sort/... (multipart: file upload hoặc đường dẫn)
   │  file upload → lưu thẳng vào temp_processing/<job_id>/
   ▼
jobs.submit → thread pool SORT_WORKERS (trong tiến trình API)
   ▼
sort_prepare     đường dẫn ngoài (D:\… → /mnt/d/…) → copy; URL http(s) → tải về temp_processing/<job_id>/
sort_name        Word / render + Chandra 8 trang đầu → Qwen → (không ra tên: PDF → bìa) → ProtonX
sort_date        Word / Chandra OCR cả trang cuối → tháng, năm
sort_metadata    Word / Chandra cả trang cuối → regex
sort_sound       ffprobe + mutagen
   │  (chỉ chạy các luồng của mode, lần lượt, mỗi luồng ghi kết quả vào context["result"])
   ▼
sort_finish      ghi 1 dòng vào results/sort_YYYY-MM-DD.jsonl, xóa temp, trả kết quả về API
```

Các luồng chỉ gọi HTTP tới các dịch vụ model (Chandra, Qwen, Spellfix), nên thread trong API là đủ; muốn xử lý nhiều sách cùng lúc thì tăng `SORT_WORKERS`. Trạng thái job nằm trong bộ nhớ: khởi động lại API thì job đang chạy bị mất và `request-status` của job cũ trả 404 (kết quả đã xong vẫn còn trong file `results/`).

Lỗi ở bước nào thì các bước sau bỏ qua; bước cuối vẫn ghi kết quả với `status=failed` và API trả HTTP 500 kèm `error`.

## API

| Endpoint | Chức năng |
|---|---|
| `POST /api/sort` | Tên sách + METADATA (bắt buộc có PDF, Word tùy chọn) |
| `POST /api/sort/title` | Chỉ tìm tên sách (cần PDF hoặc Word, hoặc chỉ file bìa `bia` / `bia_path`) |
| `POST /api/sort/metadata` | Chỉ đọc METADATA trang cuối (cần PDF) |
| `POST /api/sort/info` | Chỉ lấy thông tin xuất bản trang cuối (`flow_sort_metadata`, PDF hoặc Word) |
| `POST /api/sort/sound` | Metadata file âm thanh (`flow_sort_sound`): field `audio` (upload) hoặc `audio_path` / `folder_path` |
| `GET /api/request-status/{request_id}` | Lấy kết quả khi gọi với `wait=false` hoặc khi hết thời gian chờ; không có id thì 404 |
| `GET /api/jobs/status` | Số job đang chờ / đang chạy / xong / lỗi |
| `GET /health` | Health check |

Form fields (gửi file upload hoặc đường dẫn; nếu gửi cả hai thì file upload được ưu tiên):

| Field | Ý nghĩa |
|---|---|
| `pdf`, `word` | File upload (.pdf, .docx) |
| `pdf_path` | PDF ở thư mục ngoài: `D:\sach\a.pdf`, `/mnt/d/sach/a.pdf`; URL http(s) chỉ nhận khi `SORT_ALLOW_URL=1` |
| (lọc file) | | Khi tự tìm PDF/Word trong thư mục (`folder_path`, Word đi kèm PDF) bỏ qua file tên có "Bìa"/"bia", ảnh `.jpg`/`.jpeg`, PDF tên kết thúc `- <số>` (vd `Tư tưởng Hồ Chí Minh - 01.pdf`). Đường dẫn chỉ định đích danh thì vẫn xử lý |
| `word_path` | Word ở thư mục ngoài hoặc URL. Bỏ trống mà thư mục chứa PDF có file Word thì **ưu tiên lấy Word** (thay cho OCR 8 trang đầu): file cùng tên với PDF trước, không có thì lấy file Word đầu tiên |
| `folder_path` | Thư mục của một cuốn sách: lấy PDF đầu tiên (theo tên), sau đó tìm Word như trên |
| `bia`, `bia_path` | (`/api/sort`, `/api/sort/title`) File bìa: ảnh `.jpg/.png/.tif/.bmp/.webp` hoặc PDF bìa (OCR `SORT_BIA_MAX_PAGES`=2 trang đầu). Chỉ dùng khi Word / PDF không ra tên sách; kết quả `text_source=bia` |
| `wait` | Mặc định `true` |
| `timeout` | Số giây chờ, mặc định `SORT_API_TIMEOUT=600` |

Đường dẫn Windows được đổi sang đường dẫn WSL (`D:\` → `/mnt/d/`). Sai đường dẫn hoặc sai định dạng file thì API trả HTTP 400 ngay.
- `wait=true`: API chờ xử lý xong rồi trả kết quả (HTTP 200).
- Gọi với `wait=false`, hoặc khi quá `timeout`: API trả HTTP 202 kèm `request_id`. Tra kết quả bằng `/api/request-status/{request_id}`.

```bash
curl -F "pdf=@sach.pdf" -F "word=@sach.docx" http://<host>:9091/api/sort
```

```bash
curl -F "pdf_path=D:\sach\lich-su\lich-su.pdf" http://<host>:9091/api/sort
```

```bash
curl -F "folder_path=D:\sach\lich-su" http://<host>:9091/api/sort
```

Response:
```json
{
  "job_id": "…", "request_id": "…", "status": "completed", "error": "",
  "result": {
    "ten_sach": "...", "ten_sach_raw": "...",
    "metadata": "...", "metadata_raw": "...", "metadata_score": null, "metadata_box": null,
    "metadata_error": "", "luu_chieu_thang": 3, "luu_chieu_nam": 2021,
    "luu_chieu_cau": "In xong và nộp lưu chiểu tháng 3 năm 2021.",
    "text_source": "word|ocr", "page_count": 120,
    "pdf_file": "sach.pdf", "docx_file": "sach.docx"
  }
}
```
Các trường `*_raw` là bản trước khi sửa chính tả. `luu_chieu_thang` / `luu_chieu_nam` (số) tách bằng regex từ dòng "In xong và nộp lưu chiểu tháng 3 năm 2021" trong METADATA (nhận cả `tháng 03/2021`, `tháng 3-2021`, `tháng 3 – 2021`, `tháng 3,2021`, `T3/2021`, `3/2021` ngay sau chữ In xong, `tháng 3-21`, `tháng ba năm 2021`, chữ không dấu); không thấy thì `null`.

## Dịch vụ phụ thuộc

| Dịch vụ | Port | Chạy bằng |
|---|---|---|
| FastAPI | 9091 | `bash_sh/sort/start-service.sh` |
| Chandra OCR 2 (vLLM) | 8101 | `service/vllm-chandra.service` |
| Qwen3-4B-Instruct (vLLM) | 8113 | `service/llm-qwen3.service` |
| Spellfix ProtonX (GPU 0, chung với Qwen) | 8006 | `bash_sh/sort/start-service.sh` (env `ocr`, cần torch ≥ 2.6 + transformers 5; đổi Python qua `SPELLFIX_PYTHON`) |
| ffprobe (FFmpeg) | - | `sudo apt install ffmpeg`, dùng cho `/api/sort/sound`; không có thì đọc thông số bằng mutagen |

OCR SORT BOOK không cần RabbitMQ, PostgreSQL. Chạy uvicorn **1 tiến trình** (không dùng `--workers`) vì trạng thái job nằm trong bộ nhớ.

## Biến môi trường mới (tùy chọn, đều có mặc định)

| Biến | Mặc định | Ý nghĩa |
|---|---|---|
| `SORT_DATE_LAST_PAGES` | `5` | Số trang cuối tối đa (Word và PDF) để tìm dòng "In xong và nộp lưu chiểu"; PDF xét lần lượt từ trang cuối lên, dừng ở trang ra ngày |
| `SORT_INFO_LAST_PAGES` | `2` | Số trang cuối (Word và PDF) để tìm thông tin xuất bản |
| `SORT_CHANDRA_INFO_MAX_TOKENS` | `3000` | max_tokens khi Chandra OCR cả trang cuối |
| `SORT_INFO_MIN_FIELDS` | `3` | Word không có dòng Chịu trách nhiệm / In xong thì cần ít nhất chừng này trường mới dùng |
| `SORT_FIRST_PAGES` | `8` | Số trang đầu đem OCR |
| `SORT_PDF_DPI` | `200` | DPI khi render PDF |
| `SORT_CHANDRA_WORKERS` | `4` | Số trang OCR song song |
| `SORT_CHANDRA_MAX_TOKENS` | `2400` | `max_tokens` khi OCR toàn trang |
| `SORT_CHANDRA_DATE_MAX_TOKENS` | `3000` | `max_tokens` khi OCR cả trang cuối tìm ngày lưu chiểu |
| `SORT_CHANDRA_MAX_SIDE` | `1800` | Cạnh dài tối đa (px) của ảnh gửi Chandra |
| `SORT_CHANDRA_TIMEOUT` | `600` | Timeout gọi Chandra (giây) |
| `SORT_QWEN_MAX_CHARS` | `8000` | Độ dài tối đa của text gửi Qwen (0 = gửi toàn bộ). Nếu Qwen báo vượt `max-model-len` thì tự cắt còn 70% rồi gửi lại, tối đa 7 lần |
| `SORT_QWEN_MAX_TOKENS` | `300` | `max_tokens` câu trả lời của Qwen |
| `SORT_QWEN_TIMEOUT` | `300` | Timeout gọi Qwen (giây) |
| `SORT_QWEN_PAGE_CHARS` | `900` (8 trang vừa 8000 ký tự) | Mỗi trang OCR chỉ lấy chừng này ký tự đầu khi ghép gửi Qwen (0 = không giới hạn), để trang nào cũng góp phần |
| `SORT_API_TIMEOUT` | `600` | Thời gian chờ mặc định của API (giây) |
| `SORT_KEEP_TEMP` | `0` | `1` để giữ lại `temp_processing/<job_id>` sau khi xử lý |
| `SORT_WINDOWS_MOUNT` | `/mnt` | Gốc mount ổ đĩa Windows trong WSL |
| `SORT_ALLOWED_ROOTS` | (trống = không giới hạn) | Danh sách thư mục được phép đọc, cách nhau dấu phẩy, vd `D:\sach,E:\kho`. Nên đặt; để trống API ghi cảnh báo lúc khởi động |
| `SORT_ALLOW_URL` | `0` | `1` để nhận `pdf_path`/`word_path` dạng URL http(s) |
| `SORT_TEMP_MAX_AGE_HOURS` | `24` | Dọn thư mục job cũ trong `temp_processing` |
| `SORT_WORKERS` | `4` | Số job chạy cùng lúc trong API |
| `SORT_JOB_KEEP_HOURS` | `24` | Giữ trạng thái job đã xong cho `/api/request-status` |
| `SORT_RESULT_DIR` | `<project>/results` | Thư mục file kết quả `sort_YYYY-MM-DD.jsonl` |
| `SPELLFIX_API_URL` | `http://localhost:8006/spellfix/correct` | Địa chỉ dịch vụ sửa chính tả |
| `SPELLFIX_ENABLED` | `1` | `0` để tắt sửa chính tả |
| `SPELLFIX_GPU` | `0` | GPU chạy spellfix (`bash_sh/sort/start-service.sh` đặt `CUDA_VISIBLE_DEVICES`); GPU 1 dành cho Chandra |
| `SPELLFIX_DTYPE` | `float16` (qua `bash_sh/sort/start-service.sh`) | `float32` / `float16` / `bfloat16`. Hết VRAM lúc nạp model thì tự chuyển sang CPU |
| `SPELLFIX_*` (trong `spellfix_api.py`) | | model, beams, batch, dtype, ngưỡng độ dài |
| `LOG_DIR` | `<project>/logs` | Thư mục gốc chứa log |
| `LOG_LEVEL` | `INFO` | Mức log |
| `LOG_KEEP_DAYS` | `30` | Số ngày giữ file log |

## Log (`app/logger.py`)

Mỗi nhóm ghi vào `logs/<nhóm>/<nhóm>_YYYY-MM-DD.log`, mỗi ngày một file (nhiều tiến trình ghi chung an toàn), file cũ hơn `LOG_KEEP_DAYS` ngày tự xóa. Dòng log có dạng `thời gian [LEVEL] [pid] [job_id] nội dung`, nên dùng `grep <job_id> -r logs/` để lần theo một job qua mọi nhóm.

| Thư mục | Nội dung |
|---|---|
| `logs/api/` | Mỗi request: IP, method, path, status, thời gian; từng job nhận vào, kết quả, timeout, lỗi 400 |
| `logs/input/` | Đổi đường dẫn, chọn PDF/Word, copy/tải vào `temp_processing`, đọc Word |
| `logs/first_8_pages/` | Render 8 trang đầu, Chandra OCR từng trang (số ký tự, thời gian, trang lỗi) |
| `logs/last_page/` | Render trang cuối, Chandra OCR cả trang, trang nào ra / không ra dòng In xong |
| `logs/qwen/` | Text gửi Qwen, phản hồi gốc, lỗi parse JSON, không tìm thấy tên sách |
| `logs/spellfix/` | Các đoạn bị sửa (trước → sau), lỗi dịch vụ |
| `logs/database/` | Ghi kết quả vào `results/sort_YYYY-MM-DD.jsonl`, lỗi ghi file |
| `logs/connection/` | Kiểm tra cổng lúc API khởi động (Chandra, Qwen, Spellfix), lỗi không kết nối được trong lúc chạy |
| `logs/task/` | Vòng đời job: bắt đầu, từng luồng xong / lỗi, xóa temp |
| `logs/spellfix_service/` | Service ProtonX: nạp model, từng request, đoạn bị từ chối do lệch độ dài |

Ngoài ra `bash_sh/sort/start-service.sh` vẫn ghi stdout thô của từng tiến trình vào `logs/api/api_server.log` và `logs/services/`.

## Tool Windows gọi API (`exe/`)
`exe/sxb_ocr.py` (đóng gói `exe\build.bat` → `SXBTool_OCR.exe`): 3 tab của tool SXB (quét kho, đối chiếu Excel, kiểm tra nội dung) + tab **OCR & copy** (OCR tên sách + thời gian lưu chiểu rồi chép vào `Năm/thang MM/<tên mới>/`), **Cần kiểm tra** (copy file các sách cần xem vào `CAN_KIEM_TRA/<loại>/`), **Quét metadata**, **Quét thông tin trang cuối** gọi tới `http://localhost:9091` (OCR/AI chạy trên chính máy này, đổi bằng biến `SXB_OCR_API`), kết quả ra Excel. Chi tiết: [exe/README.md](../exe/README.md).

## Luồng ngày lưu chiểu (flow_sort_date)
Đầu vào giống luồng tên sách (PDF + Word `.docx`/`.doc` cùng thư mục, cùng quy tắc chọn / bỏ qua file). Lấy tháng, năm ở dòng "In xong và nộp lưu chiểu tháng 3 năm 2021".
1. **Word trước**: đọc `SORT_DATE_LAST_PAGES` trang **cuối** Word (bộ đọc Word giống luồng tên: ngắt trang thật, text box, bảng, phông cũ TCVN3 / VNI). Thấy dòng In xong / lưu chiểu có tháng/năm thì lấy luôn khối thông tin xuất bản quanh dòng đó, không OCR.
2. **Không được thì PDF**: xét **lần lượt từng trang từ trang cuối lên, tối đa `SORT_DATE_LAST_PAGES` (5) trang**, chỉ render trang đang xét. Mỗi trang: **Chandra OCR cả trang** (`SORT_CHANDRA_DATE_MAX_TOKENS`) → regex; chưa ra ngày thì sang trang trước. Dừng ở trang đầu tiên ra **tháng + năm** (dòng In xong, hoặc ngày Quyết định xuất bản); trang chỉ ra năm thì giữ tạm, xét tiếp trang trước, hết 5 trang mà không có tháng thì dùng năm đó; `metadata` = khối thông tin xuất bản quanh dòng đó. `metadata_score`, `metadata_box` luôn `null` (giữ trường cho tương thích).
3. **Regex**: lấy ở dòng có chữ "in xong" / "lưu chiểu"; dòng đó **không ra tháng + năm** (không có, hoặc chỉ có năm như `quý I/2021`) thì lấy **ngày Quyết định xuất bản** (`Quyết định xuất bản số 79-QĐ/NXBCTQG ngày 05/3/2021`, `Số QĐXB: 12/QĐ-NXB ngày 05 tháng 3 năm 2021`, `Số 79-QĐ/NXBCTQG ngày …` với nhãn ở dòng trên; ngày xuống dòng thì nối dòng sau), `luu_chieu_theo` = `qdxb`; vẫn không có thì năm của dòng In xong. Không lấy ngày Giấy đăng ký / Giấy phép / quyết định khác. Dòng In xong (nhận cả lỗi font / OCR như `lou chieau`, `lưu chiếu`); chỉ xét phần từ chữ đó trở đi, dòng đó chưa có năm thì nối dòng ngay sau (xuống dòng giữa câu). Không có cả dòng In xong lẫn Quyết định xuất bản thì **không trả ngày** (không lấy ngày Giấy phép / Số đăng ký). Nhận `tháng 3 năm 2021`, `tháng 03/2021`, `tháng 3-2021`, `tháng 3 – 2021` (gạch dài), `tháng 3,2021`, `T3/2021`, `3/2021` (số đầu tiên ngay sau chữ In xong), năm 2 chữ số `tháng 3-21`, tháng viết chữ `tháng ba / mười hai / giêng năm 2021`, chữ không dấu; chỉ có năm (`năm 2021`, `quý I năm 2021`) thì trả năm, tháng `null`; nhiều dòng In xong thì lấy dòng cuối. Luồng thông tin trang cuối (tab 7) tách *Lưu chiểu* theo cùng quy tắc.
4. Kết quả thêm: `luu_chieu_source` (`word`/`ocr`), `metadata_page` (trang PDF lấy được), `date_word_reason` (lý do bỏ Word). Text OCR được sửa chính tả để hiển thị, regex chạy trên text gốc trước.
5. `/api/sort/metadata` nhận cả Word (`word`, `word_path`); chỉ có Word mà không thấy ngày thì trả rỗng, không báo lỗi. Test: xem mục **Script test** bên dưới.

## Luồng thông tin trang cuối (flow_sort_metadata)
1. **Word trước**: đọc `SORT_INFO_LAST_PAGES` trang cuối Word, lấy từ dòng *Chịu trách nhiệm xuất bản* (kèm khối NXB ngay phía trên) trở xuống.
2. **Không được thì PDF**: render N trang cuối, Chandra OCR **cả trang** từ trang cuối lên, dừng ở trang có *Chịu trách nhiệm xuất bản* / *In xong*. Trường còn trống được bổ sung từ các trang đã OCR khác (vd bìa sau có địa chỉ, điện thoại NXB).
3. **Regex** tách các trường nguồn *Trang cuối*: `nha_xuat_ban`, `ban_bien_tap`, `isbn`, `luu_chieu` (+ `luu_chieu_thang`, `luu_chieu_nam`), `giam_doc_tong_bien_tap`, `che_ban`, `trinh_bay_bia`, `dia_chi_nxb`, `dien_thoai_nxb`, `sdt_quan_ly_xuat_ban`, `sdt_tong_bien_tap`, `fax_nxb`, `dia_chi_co_so_in`, `ngay_xac_nhan_dkxb` (dd/mm/yyyy). Các trường nguồn Excel bỏ qua.
4. Kết quả: `thong_tin` (dict các trường trên), `thong_tin_text`, `thong_tin_source` (`word`/`ocr`), `thong_tin_page`, `thong_tin_error`, `thong_tin_word_reason`. Mode `all` (`/api/sort`) chạy cả luồng này.

## Luồng âm thanh (flow_sort_sound)
Phần không dùng AI theo `app/metadata/Metadata_03_AmThanh_Video.xlsx`, chạy ở `sort_sound_queue`.
1. **ffprobe** đọc thông số kỹ thuật (thời lượng, codec, tần số mẫu, kênh, bitrate) và thẻ chung (vd `originator` của máy ghi âm, `com.apple.quicktime.*` của iPhone). Không có ffprobe thì dùng mutagen.
2. **mutagen** đọc thẻ theo từng chuẩn: ID3 (mp3), MP4 (m4a), Vorbis (flac/ogg), ASF (wma); thẻ ID3 Latin-1 bị lỗi font tiếng Việt được sửa lại; ảnh bìa nhúng lưu ra `sound/anh_bia.*` trong thư mục job.
3. Kết quả: `am_thanh` (19 trường: `ten_tep_goc`, `dinh_dang`, `dung_luong`, `thoi_gian_ghi`, `thiet_bi`, `gps`, `tac_gia_trong_tep`, `thoi_luong`, `codec`, `tan_so_kenh_bitrate`, `tieu_de`, `album`, `track`, `loai_noi_dung`, `ngon_ngu`, `anh_bia`, `ban_quyen_nguon`, `tac_gia`, `nguoi_doc`), `am_thanh_ky_thuat` (số liệu gốc), `am_thanh_the` (toàn bộ thẻ), `am_thanh_error`, `audio_file`.
4. Các trường *gợi ý* (loại nội dung, ngôn ngữ, bản quyền, tác giả, người đọc) chỉ lấy từ thẻ có sẵn, cần người dùng xác nhận. `artist` không được coi là người đọc. Nhóm mô tả nội dung (ASR, tóm tắt) chưa làm.

## Luồng tên sách: đọc Word
- Nhận Word `.docx` và `.doc` (cùng thư mục với PDF: ưu tiên file cùng tên, `.docx` trước `.doc`).
- Chỉ đọc **N trang đầu** (`SORT_FIRST_PAGES`) theo ngắt trang thật trong Word (ngắt trang thủ công, ngắt section…); đọc cả text box (trang bìa hay dùng), bảng, content control. Không có ngắt trang thì chia mỗi 2.500 ký tự là 1 trang.
- Phông cũ **TCVN3** (`.VnTime`, `.VnArial`…) và **VNI-Windows** (`VNI-Times`…, chữ dạng `HÖÔÙNG DAÃN`, `ñaûng`) tự đổi sang Unicode: thử cả hai cách đổi, chọn bản ra tiếng Việt thật nhất. Chỉ nhận bản đổi khi ≥ 50% từ là âm tiết tiếng Việt hợp lệ và hơn bản gốc (`vn_score`), dòng đã có chữ Unicode tiếng Việt giữ nguyên, nên text Unicode thật hay bảng mã khác (VPS, VietWare…) không bị đổi ra chữ vô nghĩa (VPS / VietWare vẫn báo "nghi lỗi bảng mã").
- `.doc` cần cài trên server một trong: `antiword`, `catdoc` hoặc LibreOffice (`soffice`), vd `sudo apt install antiword catdoc`.
- Word **không dùng được** thì OCR 8 trang đầu PDF: đọc lỗi, không có chữ, dưới 50% là chữ cái (rác / lỗi mã hoá), **nghi lỗi bảng mã** (còn ký tự phông cũ và dưới 50% âm tiết tiếng Việt hợp lệ). Không lọc theo số ký tự tối thiểu. Không có PDF thì dùng tạm Word kém, trừ lỗi bảng mã (báo lỗi luôn, không gửi chữ rác cho Qwen).
- Đọc Word mà Qwen **không ra tên** (tên chỉ có ở ảnh bìa) và có PDF → OCR PDF rồi tìm lại; lý do ghi ở `word_fallback_reason` trong kết quả.
- Tên lấy từ Word là chữ gốc nên không qua sửa chính tả; tên từ OCR vẫn sửa bằng ProtonX.

## Script test
Các script độc lập, input là **cả một thư mục** sách (mặc định `test/data/`, quét luôn thư mục con, chọn file giống server), output ghi **ngay trong thư mục chứa script**: `ket_qua_<thời gian>/` gồm `summary.csv` (mở bằng Excel), `summary.json` và mỗi cuốn một thư mục con (`result.json`, text từng trang, ảnh trang).

| Script | Luồng | Tuỳ chọn hay dùng |
|---|---|---|
| `test/get_name/test_get_name.py` | Tên sách (`flow_sort_name`) | `--force-ocr`, `--pages`, `--dpi`, `--max-chars`, `--page-chars`, `--no-qwen`, `--no-spellfix`, `--show N`, `--no-images` |
| `test/get_date/test_get_date.py` | Ngày lưu chiểu (`flow_sort_date`) | `--force-ocr`, `--pages`, `--no-images` |
| `test/get_date/test_luu_chieu.py` | Quy tắc tách tháng / năm (bắt buộc dòng In xong / lưu chiểu), không cần máy chủ | |
| `test/get_metadata/test_get_metadata.py` | Thông tin trang cuối (`flow_sort_metadata`) | `--force-ocr`, `--pages`, `--no-images` |
| `test/bang_ma/test_bang_ma.py` | Đổi TCVN3 / VNI sang Unicode ở cả 3 luồng, text Unicode giữ nguyên (không cần máy chủ) | |
| `test/bang_ma/thong_ke_bang_ma.py` | Bảng mã các file Word trong `doi_chieu_*.xlsx` (sheet KHOP): OK (Unicode / TCVN3 / VNI đổi được), LỖI BẢNG MÃ, KÉM, ĐỌC LỖI; tên phông, tỷ lệ âm tiết hợp lệ, mẫu chữ → `test/bang_ma/thong_ke_bang_ma_*.csv` | `--tat-ca-sheet`, `--song-song N`, `--limit N` |
| `test/api/test_bia.py` | Luồng tên sách có file bìa (giả lập Chandra, Qwen): chỉ bìa, Word không ra tên → bìa, PDF bìa, bìa hỏng | (không có) |
| `test/api/test_api.py` | API + hàng đợi job (`app/jobs.py`): chờ kết quả, `wait=false`, quá timeout, 2 file song song, lỗi, 404; giả lập các luồng nên không cần dịch vụ model | (không có, chạy `python test/api/test_api.py`) |
| `test/get_sound/test_get_sound.py` | Metadata âm thanh (`flow_sort_sound`), mặc định `test/data/audio_1` | `--no-ffprobe`, `--mau` (tạo 2 file mẫu có đủ thẻ, Latin-1 lỗi font, ảnh bìa), `--limit N` |

```bash
python test/get_name/test_get_name.py
```
```bash
python test/get_date/test_get_date.py test/data/sach_1
```

## Vận hành

Chạy / dừng (dùng `.env` ở thư mục gốc):

```bash
bash_sh/sort/start-service.sh   # spellfix + API 9091 (SORT_WORKERS job cùng lúc, không cần RabbitMQ/PostgreSQL)
bash_sh/sort/stop-service.sh    # dừng API 9091 (job đang chạy bị huỷ), spellfix
```
