# Tool SXB + OCR (Windows)

Một giao diện tkinter gồm 7 tab, đóng gói được thành 1 file `SXBTool_OCR.exe`:

| Tab | Chức năng |
|---|---|
| 1. Quét kho | Kho → `kho_*.xlsx` (quá 1.048.575 dòng tự tách sheet); tool tắt / bấm Dừng giữa chừng thì lần sau hỏi quét tiếp |
| 2. Đối chiếu Excel | Khớp tên sách (bỏ dấu) với tên file / tên thư mục: **trùng hẳn** (bỏ chú thích 'Sách chuyên khảo...' vẫn được) hoặc thư mục tên rút gọn; tên sách nằm trong tên file dài hơn không tính, lọc theo năm → `KHỚP` / `LỆCH NĂM` / `KHÔNG TÌM THẤY`; thêm sheet `THU_MUC_KHO`, `CAN_OCR` |
| 3. Kiểm tra nội dung | Mở PDF/DOCX/DOC (.doc qua Microsoft Word) đọc năm xuất bản, dò lời nói đầu; PDF/DOCX đọc song song nhiều tiến trình (mặc định nửa số luồng CPU); có cache, file làm tool sập lần trước thì tự bỏ qua |
| 4. OCR & copy | Gộp "Đổi tên & lưu" và "Quét tên sách": sách KHỚP trong `doi_chieu_*.xlsx` → OCR tên sách + tháng/năm lưu chiểu → copy/move sang `Năm/thang MM/<tên mới>/` theo thời gian OCR, tên `ten_sach_mmyy_R_BBT_T01` → Excel. Xem mục dưới |
| 5. Cần kiểm tra | Từ Excel kết quả tab 4: copy toàn bộ file liên quan (Word, PDF, ảnh, bìa...) của các sách cần xem vào `CAN_KIEM_TRA/<loại>/<Năm>/<Năm Excel>_<Dòng Excel>_<tên sách>/`, kèm Excel danh sách. Xem mục dưới |
| 6. Quét metadata | Gọi `POST /api/sort/metadata` của OCR SORT BOOK: tháng/năm "In xong và nộp lưu chiểu" |
| 7. Quét thông tin trang cuối | Gọi `POST /api/sort/info`: NXB, biên tập, ISBN, lưu chiểu, giám đốc, chế bản, bìa, địa chỉ, điện thoại, fax, cơ sở in, ngày xác nhận ĐKXB |

Tab 1–3 là `sxb_tool.py` của sxb-tool (bản 4 tab, đã đổi `import fitz` thành `import pymupdf`; chỉ thêm tham số `them_tab` cho `gui()` để gắn tab OCR). Tab 4 "Đổi tên & lưu" của sxb-tool được `sxb_ocr.py` gỡ ra, thay bằng tab "OCR & copy" dùng lại các hàm đổi tên, chép file của nó; tab 4–7 trong `sxb_ocr.py`. Chi tiết tab 1–4 (thư mục `ketqua/` cạnh exe, cache, chạy tiếp khi tool tắt, đọc đa nhân): xem README của sxb-tool.
Khi sxb-tool có bản mới: copy `sxb_tool.py`, `chuan_hoa_excel.py` sang đây, thêm lại hook `them_tab` (tham số của `gui()` và 2 dòng trước `root.mainloop()`), chạy `--test`. `sxb_ocr.py` phải gọi `multiprocessing.freeze_support()` đầu `__main__`, không thì exe mở thêm cửa sổ mỗi khi đọc song song.
`chuan_hoa_excel.py` gom các sheet năm của file Excel tổng hợp về một bộ cột cố định.

## Tab 4: OCR & copy
- Đầu vào: **File doi_chieu_*.xlsx** của tab 2/3 (mặc định file mới nhất trong `ketqua/2_doi_chieu`); chỉ xử lý dòng sheet `KHOP`, file chọn theo cột *File PDF chọn* / *File Word chọn*, file bìa theo *Tất cả file tìm được* (giống "Đổi tên & lưu" cũ). **File BBT_distinct.xlsx** để lấy mã BBT (bỏ trống thì tự tìm cạnh file đối chiếu / cạnh exe).
- **Sách LỆCH NĂM** (sheet `CAN_XU_LY`, *Tình trạng* bắt đầu bằng `LỆCH NĂM`; tick *Xử lý sách LỆCH NĂM*, mặc định bật): chưa có file chọn sẵn nên lấy các file Word / PDF trong *Tất cả file tìm được*, mỗi thư mục chọn 1 cặp ruột (như tab 2, bỏ bìa), thư mục có năm sửa gần năm Excel nhất xếp trước; OCR lần lượt **tối đa 3 cặp**, dừng ở cặp có tên OCR **KHỚP** tên Excel (không cặp nào khớp thì lấy cặp đầu, sau đó OCR bìa / CẦN XEM như sách thường). Ghi chú: 'LỆCH NĂM…: thử 2/3 cặp file, chọn …'. Dòng `KHÔNG TÌM ĐƯỢC` không xử lý.
- **PDF scan** (sheet `CAN_OCR`, tick *Xử lý PDF scan*, mặc định bật): OCR tên + thời gian từng file; tên OCR **KHỚP** đúng một sách trong Excel (trùng tên nhiều sách thì ưu tiên sách ở cột *Sách liên quan*, rồi sách cùng năm OCR) → nếu sách đó chưa có file tên KHỚP thì dùng scan làm file ruột (chép theo sách đó), đã có thì scan là `TRÙNG` (không chép); không khớp sách nào → copy vào `<thư mục đích>/CAN_KIEM_TRA/KHONG_KHOP_SACH/<năm OCR>/<tên file>_<mã>/` kèm `_thong_tin.txt` (có trong nhật ký, Hoàn tác được). Kết quả từng file ở sheet **PDF scan** của Excel kết quả (`ĐÃ GHÉP` / `TRÙNG` / `KHÔNG KHỚP SÁCH` / `OCR LỖI` / `KHÔNG CÒN FILE`). Dùng được cả file đối chiếu đã tách chỉ còn 2 sheet `CAN_XU_LY` (LỆCH NĂM) + `CAN_OCR`.
- Mỗi cặp PDF + Word gọi máy chủ 2 lần: `/api/sort/title` (tên sách) và `/api/sort/metadata` (tháng, năm in xong & nộp lưu chiểu). Nhiều dòng Excel cùng trỏ một file thì chỉ OCR một lần.
- Word / PDF không ra tên sách (hoặc sách chỉ có bìa) mà có file bìa OCR được (ảnh `.jpg/.png/.tif/.bmp/.webp` hoặc PDF bìa, không tính `.ai/.psd/.eps`) thì gửi riêng file bìa lên `/api/sort/title`, ưu tiên bìa 1 / bìa trước. Cột *Tên dùng* ghi `OCR (bìa)`. Mỗi cuốn chỉ thử bìa một lần, kết quả lưu cache (cả các cuốn đã OCR từ lần trước).
- Thời gian OCR quyết định thư mục `<thư mục đích>/<Năm>/thang <MM>/` và phần `mmyy` của tên mới; OCR chỉ ra năm thì `thang 00`, tên không có `mmyy`; Ngày chỉ lấy từ OCR (dòng "In xong và nộp lưu chiểu", tối đa 5 trang cuối), **không dùng ngày Excel**. Năm OCR **trước 2006** (mốc `SXB_NAM_BAT_DAU`, mặc định 2006; năm 2006 vẫn hợp lệ): **không chép vào đích**, trạng thái `TRƯỚC 2006`, luôn **copy** (kể cả khi tick Di chuyển) mọi file liên quan (file cùng năm trong *Tất cả file tìm được* + Word / PDF đã chọn) vào `<thư mục đích>/CAN_KIEM_TRA/TRUOC_2006/<năm OCR>/<Năm Excel>_<Dòng Excel>_<tên>/` kèm `_thong_tin.txt` và `CAN_KIEM_TRA/can_kiem_tra_tab4_<ngày giờ>.xlsx`; có trong nhật ký nên Hoàn tác được. **Cả Word lẫn PDF không ra dòng In xong** (hoặc năm ngoài 1945 – năm sau): không chép vào thư mục năm, trạng thái `KHÔNG RA NGÀY`, cột *Thời gian theo* = `KHÔNG RA NGÀY`, luôn **copy** mọi file liên quan vào `<thư mục đích>/CAN_KIEM_TRA/KHONG_NGAY/<Năm Excel>/<Năm Excel>_<Dòng Excel>_<tên>/` (giống `TRUOC_2006`). Sách OCR lỗi (máy chủ lỗi) hoặc không có Word / PDF ruột thì không copy (chạy lại OCR / xem ở tab 5). **Tên sách theo**: Excel (mặc định, tên chuẩn trong danh mục) hoặc OCR (không ra thì dùng Excel). Phần `R_BBT_lần XB` lấy từ Excel như cũ.
- Cột **Tên khớp** so tên Excel với tên OCR: viết thường, bỏ dấu, ký tự không phải chữ / số thành dấu cách, bỏ chú thích cuối tên ('Sách chuyên khảo', 'Tái bản'...), rồi so **bằng nhau** → `KHỚP`, khác đi dù một chữ → `KHÁC` (tên này nằm trong tên kia, hay chỉ giống phần lớn, đều `KHÁC`: 'Kinh tế số Việt Nam' ≠ 'Kinh tế số Việt Nam: Thực trạng và giải pháp', '…xã Nam Vân' ≠ '…xã An Lĩnh'); trống = OCR không ra tên. Không còn nhãn `GẦN GIỐNG`. Tick **Chỉ chép sách có tên OCR KHỚP tên Excel** (mặc định bật): tên Word / PDF không `KHỚP` (khác hoặc không ra tên) thì **OCR thêm file bìa** (bìa chọn như mục Bìa), lấy tên khớp hơn (bìa `KHỚP` thì chép, *Tên dùng* = `OCR (bìa)`, ghi chú 'tên lấy theo bìa'); vẫn không `KHỚP` (hoặc không có bìa) thì **không chép**, trạng thái `CẦN XEM` - thường do tab 2/3 chọn nhầm file; sửa xong chạy lại. Bỏ tick: chép mọi sách (tên chỉ để tham khảo), chỉ OCR bìa khi không ra tên. Với `ocr_copy_20261105_145917.xlsx`: 1.827 sách KHỚP; 1.045 không khớp, trong đó chỉ 182 có bìa để OCR.
- **Chạy thử**: OCR + lập kế hoạch, in danh sách nguồn → đích, xuất Excel, chưa chép. **Thực hiện**: OCR (dùng lại kết quả của lần chạy thử) rồi chép / di chuyển. Kết quả OCR lưu dần ở `ocr_cache.jsonl` trong thư mục kết quả: bấm **Dừng** hoặc tool tắt thì lần sau chỉ OCR các cuốn chưa xong và các cuốn lỗi; file đã chép thì bỏ qua.
- **File bìa** (chép vào `<tên mới>/Bia/`, đổi `_R_` thành `_B_`; cũng là file đem OCR khi chưa ra tên) trong thư mục file ruột:
  - tên có chữ **Bìa** đứng đầu một từ, mọi định dạng (`Bìa.ai`, `bia_sau.psd`, `BiaSach-LTTT.png`; không tính `Colombia`), cả trong thư mục con; trừ khi chính tên sách có chữ Bìa;
  - mọi file **.jpg / .jpeg**;
  - **PDF tên kết thúc `- <số>`** (vd `Tư tưởng Hồ Chí Minh - 01.pdf`).
  File khác (`.png`, `.ai`, `.psd`, `.eps`, PDF 1–2 trang…) không có chữ Bìa trong tên thì không coi là bìa. Chọn bìa để OCR: ưu tiên file tên có Bìa (bìa 1 / bìa trước trước, bìa 4 / bìa sau cuối), rồi PDF `- số`, rồi `.jpg`.
- **Hoàn tác...**: chọn `nhat_ky_chep_<ngày giờ>.csv` (mỗi lần Thực hiện ghi nguồn → đích của các file sắp chép, trước khi chép) để xoá bản chép (copy) hoặc trả file về chỗ cũ (di chuyển), rồi xoá thư mục rỗng; hiện số file và lý do bỏ qua, hỏi lại trước khi làm. Bản chép khác dung lượng nguồn, nguồn không còn, hoặc chỗ cũ đã có file thì giữ nguyên. Lần chạy cũ chưa có nhật ký thì chọn `ocr_copy_*.xlsx` của lần đó: chỉ đúng khi lần đó **copy** (xoá các file ở cột *Đường dẫn đích*; dòng nào file ruột nguồn đã mất thì giữ nguyên).
- Excel `ketqua/4_ocr_copy/ocr_copy_<ngày giờ>.xlsx`, sắp theo thời gian mới: STT, Năm (Excel), Dòng Excel, Tên sách (Excel), Tên sách (OCR), Tên khớp, Tên dùng, Tháng / Năm lưu chiểu, Thời gian theo (`OCR` / `OCR (QĐXB)`: dòng In xong không ra tháng nên lấy ngày Quyết định xuất bản / `OCR (chỉ năm)` / `KHÔNG RA NGÀY`), BBT, Lần XB, Tên file mới, File Word, File PDF, Đường dẫn đích, Trạng thái (`ĐÃ LƯU`, `CHẠY THỬ`, `TRƯỚC 2006`, `CẦN XEM`, `LỖI CHÉP`, `KHÔNG CÓ FILE`, `KHÔNG RA NGÀY`, `CHƯA CHÉP`), Ghi chú.

## Tab 5: Cần kiểm tra
- Đầu vào: **Kết quả ocr_copy_*.xlsx** của tab 4 (hoặc file thống kê có sheet `Bảng gốc`), **File doi_chieu_*.xlsx** đã dùng ở tab 4 (để biết các file của từng sách, nối theo *Năm (Excel)* + *Dòng Excel*), **Nơi tạo CAN_KIEM_TRA** (chọn đúng thư mục tên `CAN_KIEM_TRA` thì dùng luôn).
- Phân loại, mỗi sách chỉ vào **loại đầu tiên khớp** theo thứ tự bảng dưới (loại khớp khác ghi ở cột *Lý do khác*):

| Thư mục | Điều kiện | Mặc định |
|---|---|---|
| `TRUOC_2006` | Năm lưu chiểu OCR trước 2006 (Excel cũ có *Thời gian theo* = `Excel (QĐXB)` thì tính năm Excel), ngoài khoảng đang chuẩn hoá: xem lại ngày tháng. Thư mục con theo năm lưu này (`TRUOC_2006/1951/…`), các loại khác theo năm Excel; thư mục sách `<Năm Excel>_<Dòng Excel>_<tên>` (cố định, tab 4 đã copy thì tab 5 bỏ qua). **Xét trước mọi loại**. Đổi mốc bằng biến `SXB_NAM_BAT_DAU` | chọn |
| `KHONG_NGAY` | `KHÔNG RA NGÀY`: Word và PDF không ra dòng In xong (Excel cũ: *Thời gian theo* = `Excel (QĐXB)`); trừ OCR lỗi và sách không có file ruột. Xét ngay sau `TRUOC_2006`, cùng thư mục tab 4 đã copy | chọn |
| `1_TEN_OCR_KHAC` | Trạng thái `CẦN XEM` (tên OCR, kể cả bìa, không KHỚP Excel: khác / không ra tên; chưa chép) | chọn |
| `2_KHONG_CO_FILE_RUOT` | `KHÔNG CÓ FILE` (chỉ có bìa / file đã mất) | chọn |
| `3_LOI_CHEP` | `LỖI CHÉP` | chọn |
| `4_OCR_LOI` | Ghi chú có `OCR lỗi` | chọn |
| `5_OCR_KHONG_RA_TEN` | OCR không ra tên | chọn |
| `6_NAM_OCR_KHAC_EXCEL` | Năm lưu chiểu OCR khác năm Excel | không |
| `7_CHI_CO_NAM` | Thời gian theo `OCR (chỉ năm)` | không |

- File chép: các file trong *Tất cả file tìm được* **cùng năm** với sách + Word / PDF đã chọn (không có file cùng năm thì lấy hết); tick "Chép cả file khác năm" để lấy hết. Giữ cấu trúc thư mục con tính từ thư mục chung của các file.
- Mỗi thư mục sách có `_thong_tin.txt` (tên Excel / OCR, lưu chiểu, trạng thái, ghi chú tab 4, thư mục gốc). `CAN_KIEM_TRA/can_kiem_tra_<ngày giờ>.xlsx`: sheet *Tổng hợp* (số sách, số file, giải thích từng thư mục) + mỗi loại một sheet.
- Chỉ **copy**, không đụng file gốc; chạy lại thì file đã có bỏ qua. **Chạy thử** chỉ in kế hoạch (Excel ghi cạnh file kết quả tab 4).
- Với `ocr_copy_20261105_145917.xlsx` (chạy trước khi bỏ ngày Excel): 8 sách năm trước 2006 (1951, 1968, 1979, 2005) vào `TRUOC_2006`; 1.057 sách ngày theo QĐXB Excel vào `KHONG_NGAY`.

## Tab 6, 7: gọi API OCR
- Gọi API OCR/AI trên chính máy này: `http://localhost:9091` (service `app/main.py` của project này, chạy bằng `bash_sh/sort/start-service.sh`); muốn gọi máy khác thì đặt biến môi trường `SXB_OCR_API` (vd `http://your_ip:9091`) trước khi mở exe; nút **Kiểm tra kết nối** gọi `/health`.
- Chọn 1 file hoặc thư mục (tick quét cả thư mục con). Mỗi PDF là 1 cuốn, có Word `.docx`/`.doc` cùng thư mục (ưu tiên cùng tên, `.docx` trước `.doc`) thì gửi kèm để máy chủ đọc Word thay OCR 8 trang đầu; thư mục chỉ có Word thì gửi Word. Tab 6: máy chủ đọc trang cuối Word trước, không thấy ngày thì cắt khối thông tin xuất bản ở trang cuối PDF rồi OCR. Tab 7: máy chủ đọc trang cuối Word trước, không có thì OCR cả trang cuối PDF.
- **Excel đường dẫn** (tuỳ chọn, ưu tiên hơn ô thư mục): quét theo danh sách file trong Excel kết quả của các tab trước. Nút **Mới nhất** lấy file mới nhất trong `ketqua/` (`3_doi_ten` → `2_doi_chieu` → `1_quet_kho`), **Bỏ** để quay lại quét thư mục.
  - `doi_chieu_*.xlsx` sau tab 4: lấy file ở cột *Đường dẫn đích* (bỏ file trong thư mục `Bia/`).
  - `doi_chieu_*.xlsx` của tab 2/3: mỗi dòng 1 cuốn theo cột *File PDF chọn* / *File Word chọn*.
  - `kho_*.xlsx` của tab 1: cột *Đường dẫn*, gom theo thư mục và lọc giống khi quét thư mục.
  - Excel kết quả tab 4–6: cột *File PDF* / *File Word* (vd chạy lại các cuốn lỗi ở tab khác).
  - File trong Excel không còn trên ổ thì bỏ qua, ghi lý do "không còn file" ra log.
- Khi quét thư mục, **bỏ qua**: file có chữ "Bìa"/"bia" trong tên (không phân biệt hoa thường, dấu), ảnh `.jpg`/`.jpeg`, PDF có tên kết thúc bằng `- <số>` (vd `Tư tưởng Hồ Chí Minh - 01.pdf`). Danh sách file bị bỏ qua in ra ô log. Chọn đích danh 1 file thì vẫn gửi.
- Mặc định **upload file** lên máy chủ. Tick **Chỉ gửi đường dẫn** khi file nằm trên ổ máy chủ đọc được (vd `D:\...` của chính máy chủ) để khỏi upload.
- Chọn số cuốn gửi song song và thời gian chờ tối đa mỗi cuốn; máy chủ trả 202 (xử lý lâu) thì tool tự hỏi lại `/api/request-status`.
- Kết quả ghi vào **Thư mục kết quả**, mặc định trong `ketqua/` cạnh exe giống tab 1–4: `ketqua/6_metadata/metadata_<ngày giờ>.xlsx`, `ketqua/7_thong_tin/thong_tin_<ngày giờ>.xlsx`. Mỗi lần quét một file mới, không ghi đè.
- Excel ghi lại sau mỗi cuốn (bấm **Dừng** vẫn giữ phần đã xong):
  - metadata: Tháng, Năm (in xong & nộp lưu chiểu, chỉ số), Nguồn (Word/OCR), Trang, Dòng In xong, METADATA, METADATA trước sửa chính tả, Điểm vùng cắt / Vùng cắt (luôn trống), Số trang, Trạng thái, Lỗi, Thời gian, File PDF
  - thông tin trang cuối: Nhà xuất bản, Ban biên tập / Biên tập viên, Mã số sách / ISBN, Lưu chiểu, Tháng/Năm lưu chiểu, Giám đốc - Tổng Biên tập, Chế bản, Trình bày bìa, Địa chỉ NXB, Điện thoại NXB, SĐT Quản lý xuất bản, SĐT Tổng Biên tập, Fax NXB, Địa chỉ cơ sở in, Ngày xác nhận ĐKXB, Số trường, Nguồn, Trang, Text trang cuối, Lý do bỏ Word, Số trang, Trạng thái, Lỗi, Thời gian, File PDF, File Word
- Cấu hình (địa chỉ API, thư mục, tuỳ chọn) lưu ở `%USERPROFILE%\.sxb_tool_ocr.json`.

## Chạy / kiểm tra / đóng gói
```bash
pip install openpyxl "pymupdf<2" python-docx pywin32 requests
python -X utf8 exe/sxb_ocr.py          # giao diện 6 tab
python -X utf8 exe/sxb_ocr.py --test   # tự kiểm tra (gồm test của sxb_tool.py, API giả lập)
python -X utf8 exe/sxb_tool.py         # chỉ tool SXB (2 tab; thêm --4tab để ra 4 tab)
python -X utf8 exe/chuan_hoa_excel.py "Tong hop SXB 2006-2025.xlsx"
```
Đóng gói: bấm đúp `exe\build.bat` (tự tìm Python, tự cài thư viện còn thiếu kể cả `pyinstaller`, chạy `--test` rồi mới build) → `exe\dist\SXBTool_OCR.exe` (~40 MB), copy sang máy khác chạy được ngay. Lần đầu mở mất vài giây vì exe tự giải nén; Windows Defender cảnh báo thì chọn "Run anyway".

Ubuntu: `bash exe/build.sh` trên chính máy Ubuntu (PyInstaller không build chéo từ Windows) → `exe/dist-linux/SXBTool_OCR`. Script tự cài gói hệ thống còn thiếu (`python3-tk`, `python3-venv`, `antiword`, `catdoc`, `binutils`, `libpython3-dev`, cần sudo), tạo `.venv-linux`, chạy `--test` rồi build và chạy `--test` lại trên file vừa build. Trên Linux file `.doc` đọc bằng `antiword`/`catdoc` thay cho Microsoft Word.
