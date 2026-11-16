# Tool SXB và gọi API OCR SORT BOOK
import hashlib, json, os, re, sys, threading, time, datetime as dt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import sxb_tool

API_MAC_DINH = os.getenv("SXB_OCR_API", "http://localhost:9091")  # OCR/AI chạy trên máy này
CAU_HINH = os.path.join(os.path.expanduser("~"), ".sxb_tool_ocr.json")
TIMEOUT_MAC_DINH = 900


# Chọn file
_BIA_RE = re.compile(r"(?<![a-z0-9])bia(?![a-z0-9])")
_PHAN_SO_RE = re.compile(r"-\s*\d+\s*$")


# Lý do bỏ qua file
def ly_do_bo_qua(ten):
    goc, duoi = os.path.splitext(os.path.basename(ten))
    duoi = duoi.lower()
    if duoi in (".jpg", ".jpeg"):
        return "ảnh jpg"
    if _BIA_RE.search(sxb_tool.khong_dau(goc).lower()):
        return "tên có chữ Bìa"
    if duoi == ".pdf" and _PHAN_SO_RE.search(goc):
        return "PDF tên kết thúc '- số'"
    return ""


# File Word dùng được
def _la_word(ten):
    return ten.lower().endswith((".docx", ".doc")) and not ten.startswith("~$") and not ly_do_bo_qua(ten)


# Thứ tự ưu tiên Word: .docx trước
def _thu_tu_word(ten):
    return (not ten.lower().endswith(".docx"), ten)


# Danh sách Word dùng được
def _cac_word(files):
    ds, da_co = [], set()
    for f in sorted((f for f in files if _la_word(f)), key=_thu_tu_word):
        goc = os.path.splitext(f)[0].lower()
        if goc not in da_co:
            da_co.add(goc)
            ds.append(f)
    return ds


# File PDF dùng được
def _la_pdf(ten):
    return ten.lower().endswith(".pdf") and not ly_do_bo_qua(ten)


# Tìm Word đi kèm PDF
def tim_word(pdf):
    thu_muc = os.path.dirname(pdf)
    goc = os.path.splitext(os.path.basename(pdf))[0].lower()
    ds = _cac_word(os.listdir(thu_muc))
    cung_ten = [n for n in ds if os.path.splitext(n)[0].lower() == goc]
    chon = cung_ten or ds
    return os.path.join(thu_muc, chon[0]) if chon else ""


# Liệt kê các cuốn cần quét
def liet_ke_sach(dau_vao, de_quy=True, can_pdf=False, bo_qua=None):
    if os.path.isfile(dau_vao):
        if dau_vao.lower().endswith(".pdf"):
            return [(dau_vao, "" if can_pdf else tim_word(dau_vao))]
        return [] if can_pdf or not _la_word(os.path.basename(dau_vao)) else [("", dau_vao)]
    sach = []
    for goc, dirs, files in os.walk(dau_vao):
        dirs.sort()
        if bo_qua is not None:
            bo_qua += [(os.path.join(goc, f), ly_do_bo_qua(f)) for f in sorted(files)
                       if f.lower().endswith((".pdf", ".docx", ".doc", ".jpg", ".jpeg")) and ly_do_bo_qua(f)]
        pdfs = sorted(f for f in files if _la_pdf(f))
        if pdfs:
            sach += [(os.path.join(goc, f), "" if can_pdf else tim_word(os.path.join(goc, f))) for f in pdfs]
        elif not can_pdf:
            sach += [("", os.path.join(goc, f)) for f in _cac_word(files)]
        if not de_quy:
            break
    return sach


# Excel đường dẫn: cột nhận được (kết quả tab 1-4 và tab 5-7)
COT_DICH = ("Đường dẫn đích",)
COT_PDF = ("File PDF chọn", "File PDF")
COT_WORD = ("File Word chọn", "File Word")
COT_DUONG_DAN = ("Đường dẫn",)


# Thư mục kết quả của tab 1-4 (ketqua/ cạnh exe)
def thu_muc_ketqua():
    noi = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(sxb_tool.__file__))
    return os.path.join(noi, "ketqua")


# File Excel kết quả mới nhất của tab 1-4
def excel_moi_nhat():
    goc = thu_muc_ketqua()
    return sxb_tool.moi_nhat(os.path.join(goc, "3_doi_ten", "doi_chieu_*.xlsx"),
                             os.path.join(goc, "2_doi_chieu", "doi_chieu_*.xlsx"),
                             os.path.join(goc, "1_quet_kho", "kho_*.xlsx"))


# Đọc các cột đường dẫn của mọi sheet -> {cột: [giá trị từng dòng]}
def _doc_cot(xlsx, cot):
    import openpyxl
    wb = openpyxl.load_workbook(sxb_tool.duong_dan_dai(xlsx), read_only=True)
    out = []
    for ws in wb.worksheets:
        it = ws.iter_rows(values_only=True)
        hdr = [str(h).strip() if h is not None else "" for h in next(it, ())]
        vi_tri = {k: next((hdr.index(h) for h in ten if h in hdr), None) for k, ten in cot.items()}
        if all(v is None for v in vi_tri.values()):
            continue
        for row in it:
            d = {k: str(row[i]).strip() if i is not None and i < len(row) and row[i] is not None else ""
                 for k, i in vi_tri.items()}
            if any(d.values()):
                out.append(d)
    wb.close()
    return out


# Ghép danh sách file thành các cuốn giống khi quét thư mục
def _gom_sach(paths, bo_qua):
    theo_tm = {}
    for p in paths:
        theo_tm.setdefault(os.path.dirname(p), []).append(os.path.basename(p))
    sach = []
    for tm in sorted(theo_tm):
        files = sorted(set(theo_tm[tm]))
        bo_qua += [(os.path.join(tm, f), ly_do_bo_qua(f)) for f in files if ly_do_bo_qua(f)]
        words = _cac_word(files)
        pdfs = [f for f in files if _la_pdf(f)]
        for f in pdfs:
            goc = os.path.splitext(f)[0].lower()
            cung_ten = [w for w in words if os.path.splitext(w)[0].lower() == goc]
            w = (cung_ten or words or [""])[0]
            sach.append((os.path.join(tm, f), os.path.join(tm, w) if w else ""))
        if not pdfs:
            sach += [("", os.path.join(tm, w)) for w in words]
    return sach


# Đọc Excel đường dẫn -> các cuốn [(pdf, word)]
def liet_ke_tu_excel(xlsx, bo_qua=None):
    bo_qua = [] if bo_qua is None else bo_qua
    rows = _doc_cot(xlsx, {"dich": COT_DICH, "pdf": COT_PDF, "word": COT_WORD, "duong_dan": COT_DUONG_DAN})
    sach = []
    if any(r["dich"] for r in rows):
        for r in rows:
            ds = [p.strip() for p in r["dich"].splitlines() if p.strip()]
            ds = [p for p in ds if os.path.basename(os.path.dirname(p)).lower() != "bia"]
            sach += _gom_sach(ds, bo_qua)
    elif any(r["pdf"] or r["word"] for r in rows):
        for r in rows:
            pdf, word = r["pdf"].splitlines()[0] if r["pdf"] else "", r["word"].splitlines()[0] if r["word"] else ""
            if pdf or word:
                sach.append((pdf, word))
    else:
        sach = _gom_sach([r["duong_dan"] for r in rows if r["duong_dan"]], bo_qua)
    con, thieu = [], []
    for pdf, word in dict.fromkeys(sach):
        if word and not os.path.isfile(sxb_tool.duong_dan_dai(word)):
            thieu.append(word); word = ""
        if pdf and not os.path.isfile(sxb_tool.duong_dan_dai(pdf)):
            thieu.append(pdf); pdf = ""
        if pdf or word:
            con.append((pdf, word))
    bo_qua += [(p, "không còn file") for p in thieu]
    return con


# Gọi API

# Kiểm tra kết nối máy chủ
def kiem_tra_ket_noi(api):
    import requests
    r = requests.get(api.rstrip("/") + "/health", timeout=10)
    r.raise_for_status()
    return r.json()


# Upload một cuốn lên máy chủ và chờ kết quả
def goi_api(api, loai, pdf="", word="", timeout=TIMEOUT_MAC_DINH, dung=None, bia=""):
    import requests
    url = f"{api.rstrip('/')}/api/sort/{loai}"
    data = {"wait": "true", "timeout": str(timeout)}
    files, mo = {}, []
    try:
        for ten, p in (("pdf", pdf), ("word", word), ("bia", bia)):
            if p:
                f = open(sxb_tool.duong_dan_dai(p), "rb"); mo.append(f)
                files[ten] = (os.path.basename(p), f)
        r = requests.post(url, data=data, files=files or None, timeout=timeout + 60)
    finally:
        for f in mo:
            f.close()
    if r.status_code == 202:
        return _cho_ket_qua(api, r.json()["request_id"], timeout, dung)
    try:
        kq = r.json()
    except ValueError:
        r.raise_for_status()
        raise
    if r.status_code >= 400 and "result" not in kq:
        raise RuntimeError(f"HTTP {r.status_code}: {kq.get('detail', kq)}")
    return kq


# Chờ kết quả job
def _cho_ket_qua(api, request_id, timeout, dung=None):
    import requests
    het = time.time() + timeout
    while time.time() < het:
        if dung is not None and dung.is_set():
            raise RuntimeError("Đã dừng")
        r = requests.get(f"{api.rstrip('/')}/api/request-status/{request_id}", timeout=30)
        r.raise_for_status()
        st = r.json()
        if st["status"] == "SUCCESS" and isinstance(st.get("result"), dict):
            return dict(st["result"], request_id=request_id)
        if st["status"] in ("FAILURE", "REVOKED"):
            raise RuntimeError(f"Máy chủ báo {st['status']}: {st.get('result')}")
        time.sleep(5)
    raise TimeoutError(f"Quá {timeout}s chưa có kết quả (request_id={request_id})")


# Excel kết quả
COT_TEN_SACH = [("stt", "STT", 6), ("ten_sach", "Tên sách", 45), ("ten_sach_raw", "Tên sách (trước sửa chính tả)", 45),
                ("nguon", "Nguồn text", 10), ("so_trang", "Số trang", 9), ("trang_thai", "Trạng thái", 12),
                ("loi", "Lỗi", 40), ("giay", "Thời gian (s)", 12), ("pdf", "File PDF", 70), ("word", "File Word", 70)]
COT_METADATA = [("stt", "STT", 6), ("thang", "Tháng", 7), ("nam", "Năm", 7), ("nguon", "Nguồn", 8),
                ("trang", "Trang", 7), ("cau", "Dòng In xong", 50),
                ("metadata", "METADATA", 70), ("metadata_raw", "METADATA (trước sửa chính tả)", 70),
                ("diem", "Điểm vùng cắt", 12), ("vung", "Vùng cắt (x1,y1,x2,y2)", 22), ("so_trang", "Số trang", 9),
                ("trang_thai", "Trạng thái", 12), ("loi", "Lỗi", 40), ("giay", "Thời gian (s)", 12), ("pdf", "File PDF", 70),
                ("word", "File Word", 70)]
COT_THONG_TIN = [("stt", "STT", 6), ("nha_xuat_ban", "Nhà xuất bản", 35),
                 ("ban_bien_tap", "Ban biên tập / Biên tập viên", 40), ("isbn", "Mã số sách / ISBN", 20),
                 ("luu_chieu", "Lưu chiểu", 40), ("luu_chieu_thang", "Tháng lưu chiểu", 9),
                 ("luu_chieu_nam", "Năm lưu chiểu", 9), ("giam_doc_tong_bien_tap", "Giám đốc - Tổng Biên tập", 30),
                 ("che_ban", "Chế bản", 25), ("trinh_bay_bia", "Trình bày bìa", 25), ("dia_chi_nxb", "Địa chỉ NXB", 40),
                 ("dien_thoai_nxb", "Điện thoại NXB", 18), ("sdt_quan_ly_xuat_ban", "SĐT Quản lý xuất bản", 18),
                 ("sdt_tong_bien_tap", "SĐT Tổng Biên tập", 18), ("fax_nxb", "Fax NXB", 18),
                 ("dia_chi_co_so_in", "Địa chỉ cơ sở in", 40), ("ngay_xac_nhan_dkxb", "Ngày xác nhận ĐKXB", 14),
                 ("so_truong", "Số trường", 9), ("nguon", "Nguồn", 8), ("trang", "Trang", 7),
                 ("text", "Text trang cuối", 70), ("ly_do_bo_word", "Lý do bỏ Word", 30), ("so_trang", "Số trang", 9),
                 ("trang_thai", "Trạng thái", 14), ("loi", "Lỗi", 40), ("giay", "Thời gian (s)", 12),
                 ("pdf", "File PDF", 70), ("word", "File Word", 70)]
TRUONG_THONG_TIN = [k for k, _, _ in COT_THONG_TIN[1:17]]

# Cột, tên sheet, tên file, thư mục trong ketqua/ theo loại quét
LOAI = {"title": (COT_TEN_SACH, "Tên sách", "ten_sach", "5_ten_sach"),
        "metadata": (COT_METADATA, "Metadata", "metadata", "6_metadata"),
        "info": (COT_THONG_TIN, "Thông tin trang cuối", "thong_tin", "7_thong_tin")}


# Tạo một dòng Excel
def dong_ket_qua(loai, stt, pdf, word, kq, giay, loi=""):
    r = (kq or {}).get("result") or {}
    loi = loi or (kq or {}).get("error", "") or r.get("metadata_error", "") or r.get("thong_tin_error", "")
    d = dict(stt=stt, pdf=pdf, word=word, so_trang=r.get("page_count"), giay=round(giay, 1), loi=loi,
             trang_thai="LỖI" if (loi and not (kq or {}).get("result")) or (kq or {}).get("status") == "failed" else "OK")
    if loai == "info":
        tt = r.get("thong_tin") or {}
        d.update({k: tt.get(k) for k in TRUONG_THONG_TIN}, text=r.get("thong_tin_text", ""),
                 trang=r.get("thong_tin_page"), ly_do_bo_word=r.get("thong_tin_word_reason", ""),
                 nguon={"word": "Word", "ocr": "OCR"}.get(r.get("thong_tin_source"), r.get("thong_tin_source", "")))
        d["so_truong"] = sum(bool(d[k]) for k in TRUONG_THONG_TIN if k not in ("luu_chieu_thang", "luu_chieu_nam"))
        if d["trang_thai"] == "OK" and not d["so_truong"]:
            d["trang_thai"] = "KHÔNG RA THÔNG TIN"
    elif loai == "title":
        d.update(ten_sach=r.get("ten_sach", ""), ten_sach_raw=r.get("ten_sach_raw", ""),
                 nguon={"word": "Word", "ocr": "OCR"}.get(r.get("text_source"), r.get("text_source", "")))
        if d["trang_thai"] == "OK" and not d["ten_sach"]:
            d["trang_thai"] = "KHÔNG RA TÊN"
    else:
        box = r.get("metadata_box")
        d.update(metadata=r.get("metadata", ""), metadata_raw=r.get("metadata_raw", ""), diem=r.get("metadata_score"),
                 vung=",".join(map(str, box)) if box else "",
                 thang=r.get("luu_chieu_thang"), nam=r.get("luu_chieu_nam"), cau=r.get("luu_chieu_cau", ""),
                 trang=r.get("metadata_page"),
                 nguon={"word": "Word", "ocr": "OCR"}.get(r.get("luu_chieu_source"), r.get("luu_chieu_source", ""))
                 + (" (QĐXB)" if r.get("luu_chieu_theo") == "qdxb" else ""))  # ngày lấy theo Quyết định xuất bản
        if d["trang_thai"] == "KHÔNG THẤY" or (d["trang_thai"] == "OK" and d["thang"] is None and d["nam"] is None):
            d["trang_thai"] = "KHÔNG RA NGÀY"
        if d["trang_thai"] == "OK" and not d["metadata"]:
            d["trang_thai"] = "KHÔNG THẤY"
    return d


# Ghi file Excel kết quả
def ghi_excel(loai, dong, xlsx):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter
    cot, sheet = LOAI[loai][:2]
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append([h for _, h, _ in cot])
    for c in ws[1]:
        c.font, c.alignment = Font(bold=True), Alignment(horizontal="center", wrap_text=True)
    wrap = Alignment(wrap_text=True, vertical="top")
    for d in dong:
        ws.append([d.get(k, "") for k, _, _ in cot])
        for c in ws[ws.max_row]:
            c.alignment = wrap
    for i, (_, _, rong) in enumerate(cot):
        ws.column_dimensions[get_column_letter(i + 1)].width = rong
    ws.freeze_panes = "A2"
    wb.save(xlsx)
    return xlsx


# Quét nhiều cuốn song song
def quet_api(api, loai, sach, xlsx, so_luong=1, timeout=TIMEOUT_MAC_DINH, log=print, dung=None):
    from concurrent.futures import ThreadPoolExecutor
    dong, khoa = [None] * len(sach), threading.Lock()

    def mot(i):
        pdf, word = sach[i]
        if dung is not None and dung.is_set():
            return
        t0 = time.time()
        try:
            kq = goi_api(api, loai, pdf, word, timeout, dung)
            d = dong_ket_qua(loai, i + 1, pdf, word, kq, time.time() - t0)
        except Exception as e:
            d = dong_ket_qua(loai, i + 1, pdf, word, None, time.time() - t0, loi=f"{type(e).__name__}: {e}")
        with khoa:
            dong[i] = d
            xong = sum(x is not None for x in dong)
            ket = d.get("ten_sach") if loai == "title" else f"{d.get('so_truong')} trường, ISBN {d.get('isbn') or '-'}" \
                if loai == "info" else (
                f"tháng {d['thang']}/{d['nam']}  " if d.get("thang") else "") + (d.get("metadata") or "").replace("\n", " | ")[:80]
            log(f"[{xong}/{len(sach)}] {d['trang_thai']} {d['giay']}s  {os.path.basename(pdf or word)}  →  "
                f"{ket or d['loi']}")
            ghi_excel(loai, [x for x in dong if x], xlsx)

    with ThreadPoolExecutor(max_workers=max(1, so_luong)) as ex:
        list(ex.map(mot, range(len(sach))))
    return [x for x in dong if x]


# OCR & copy: OCR tên sách + thời gian lưu chiểu -> chép vào <đích>/<Năm>/thang <MM>/<tên mới>/ -> Excel

COT_OCR_COPY = [("stt", "STT", 6), ("nam_excel", "Năm (Excel)", 8), ("dong", "Dòng Excel", 8),
                ("ten_excel", "Tên sách (Excel)", 40), ("ten_ocr", "Tên sách (OCR)", 40), ("ten_khop", "Tên khớp", 11),
                ("nguon_ten", "Tên dùng", 9),
                ("thang_lc", "Tháng lưu chiểu", 9), ("nam_lc", "Năm lưu chiểu", 9), ("nguon_ngay", "Thời gian theo", 14),
                ("bbt", "BBT", 14), ("lan", "Lần XB", 8), ("ten_moi", "Tên file mới", 50),
                ("file_word", "File Word", 60), ("file_pdf", "File PDF", 60), ("dich", "Đường dẫn đích", 80),
                ("trang_thai", "Trạng thái", 14), ("ghi_chu", "Ghi chú", 50)]
LOAI["copy"] = (COT_OCR_COPY, "OCR & copy", "ocr_copy", "4_ocr_copy")
NAM_HOP_LE = (1945, dt.date.today().year + 1)  # năm OCR ngoài khoảng này coi như đọc sai (không ra ngày)
KHONG_RA_NGAY = "KHÔNG RA NGÀY"  # OCR không ra tháng / năm lưu chiểu: không lấy ngày Excel, không chép


def _khoa_ocr(pdf, word):
    return f"{pdf}|{word}"


# Kết quả OCR các lần trước (chạy lại không OCR lại cuốn đã xong)
def doc_cache_ocr(path):
    kq = {}
    try:
        with open(path, encoding="utf-8") as f:
            for dong in f:
                try:
                    d = json.loads(dong)
                    kq[d["khoa"]] = d
                except (ValueError, KeyError):
                    pass
    except OSError:
        pass
    return kq


# OCR một cuốn: tên sách (/api/sort/title) + tháng, năm lưu chiểu (/api/sort/metadata)
def ocr_mot(api, pdf, word, timeout=TIMEOUT_MAC_DINH, dung=None):
    d, loi = {"khoa": _khoa_ocr(pdf, word)}, []

    # Gọi một API; máy chủ báo failed (HTTP 500 có kết quả rỗng) cũng tính là lỗi để lần sau OCR lại
    def goi(loai, nhan):
        try:
            kq = goi_api(api, loai, pdf, word, timeout, dung)
        except Exception as e:
            loi.append(f"{nhan}: {type(e).__name__}: {e}")
            return {}
        if kq.get("status") == "failed" or kq.get("error"):
            loi.append(f"{nhan}: {kq.get('error') or 'máy chủ báo failed'}")
        return kq.get("result") or {}
    r = goi("title", "tên sách")
    d.update(ten_sach=r.get("ten_sach", ""), nguon_ten=r.get("text_source", ""))
    r = goi("metadata", "thời gian")
    d.update(thang=r.get("luu_chieu_thang"), nam=r.get("luu_chieu_nam"), cau=r.get("luu_chieu_cau", ""),
             theo=r.get("luu_chieu_theo", ""))  # in_xong / qdxb (ngày Quyết định xuất bản)
    d["loi"] = "; ".join(loi)
    return d


# So tên Excel với tên OCR -> KHỚP / KHÁC ('' nếu OCR không ra tên). Viết thường, bỏ dấu, ký tự không phải chữ / số
# thành dấu cách, bỏ chú thích cuối tên ('Sách chuyên khảo', 'Tái bản'...) rồi so bằng nhau; tên này nằm trong tên kia
# hay chỉ giống phần lớn đều là KHÁC ('Kinh tế số Việt Nam' vs 'Kinh tế số Việt Nam: Thực trạng và giải pháp')
def so_ten(ten_excel, ten_ocr):
    if not ten_ocr:
        return ""
    return "KHỚP" if sxb_tool.loi_ten(ten_excel) == sxb_tool.loi_ten(ten_ocr) else "KHÁC"


BIA_OCR_EXT = (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp", ".pdf")  # bìa OCR được (bỏ .ai, .psd, .eps)
_BIA_TRUOC_RE = re.compile(r"bia ?1|truoc|front")
_BIA_SAU_RE = re.compile(r"bia ?4|sau|back")


# Tên có chữ "bìa" đứng đầu một từ (đã bỏ dấu, chữ thường): 'Bìa', 'bia_truoc', 'BiaSach', '(hh) biatrong'; không bắt 'Colombia'
_TEN_BIA_RE = re.compile(r"(?<![a-z])bia")


# File bìa: tên có chữ Bìa (mọi định dạng), file .jpg / .jpeg, PDF tên kết thúc '- <số>' (vd 'Tư tưởng Hồ Chí Minh - 01')
def la_file_bia(path, key_sach=""):
    goc, duoi = os.path.splitext(os.path.basename(path))
    duoi = duoi.lower()
    if _TEN_BIA_RE.search(sxb_tool.khong_dau(goc).lower()) and not _TEN_BIA_RE.search(key_sach):
        return True
    return duoi in (".jpg", ".jpeg") or (duoi == ".pdf" and bool(_PHAN_SO_RE.search(goc)))


# Các file bìa của một sách: file bìa (la_file_bia) trong thư mục ruột; file tên có chữ Bìa thì cả thư mục con
# (sách chỉ có bìa: thư mục của file mới nhất). Thay sxb_tool.ds_bia (lấy mọi ảnh / file thiết kế, PDF 1-2 trang)
def ds_bia_chat(r):
    ds = sxb_tool._dong_file(r["cac_file"])
    tm = {os.path.dirname(x) for x in (r["file_word"], r["file_pdf"]) if x} or ({os.path.dirname(max(ds)[1])} if ds else set())
    chon, key = {r["file_word"], r["file_pdf"]}, sxb_tool.loi_ten(r["ten"])
    kq = []
    for _, p in ds:
        if p in chon or not la_file_bia(p, key):
            continue
        d = os.path.dirname(p)
        ten_bia = bool(_TEN_BIA_RE.search(sxb_tool.khong_dau(os.path.splitext(os.path.basename(p))[0]).lower()))
        if d in tm or (ten_bia and any(d.startswith(t + os.sep) for t in tm)):
            kq.append(p)
    return sorted(set(kq))


# Chọn 1 file bìa để OCR: ảnh / PDF bìa còn trên ổ, ưu tiên bìa trước (bìa 1), để bìa sau (bìa 4) sau cùng
def chon_bia(ds):
    ten = lambda p: sxb_tool.chuan_hoa(os.path.splitext(os.path.basename(p))[0])
    ok = [p for p in ds if p.lower().endswith(BIA_OCR_EXT) and os.path.isfile(sxb_tool.duong_dan_dai(p))]
    # tên có chữ Bìa trước, rồi PDF '- số', rồi ảnh .jpg khác; trong mỗi nhóm: bìa 1 / bìa trước đầu tiên, bìa 4 / bìa sau cuối
    ok.sort(key=lambda p: (not _TEN_BIA_RE.search(ten(p)), not _BIA_TRUOC_RE.search(ten(p)), bool(_BIA_SAU_RE.search(ten(p))),
                           not p.lower().endswith(".pdf"), p))
    return ok[0] if ok else ""


# OCR riêng file bìa: /api/sort/title chỉ gửi bìa -> tên sách
def ocr_bia(api, bia, timeout=TIMEOUT_MAC_DINH, dung=None):
    kq = goi_api(api, "title", timeout=timeout, dung=dung, bia=bia)
    if kq.get("status") == "failed" or kq.get("error"):
        raise RuntimeError(kq.get("error") or "máy chủ báo failed")
    return (kq.get("result") or {}).get("ten_sach", "")


TOI_DA_UNG_VIEN = 3  # sách LỆCH NĂM: số cặp file ruột tối đa đem OCR để tìm cặp có tên KHỚP


# Các cặp (PDF, Word) ruột ứng viên của sách LỆCH NĂM: mỗi thư mục một cặp (chọn như tab 2, bỏ bìa), thư mục có file
# năm sửa (Date Modified) gần năm Excel nhất xếp trước; tối đa toi_da cặp
def ung_vien_lech(r, toi_da=TOI_DA_UNG_VIEN):
    theo_tm = {}
    for d, p in sxb_tool._dong_file(r.get("cac_file")):
        if p.lower().endswith((".pdf", ".doc", ".docx")) and os.path.isfile(sxb_tool.duong_dan_dai(p)):
            theo_tm.setdefault(os.path.dirname(p), []).append((d, p))
    try:
        nam = int(r.get("nam") or 0)
    except (TypeError, ValueError):
        nam = 0
    lech = lambda tm: min(abs(int(d[:4]) - nam) if d[:4].isdigit() else 9999 for d, _ in theo_tm[tm])
    key, kq = sxb_tool.loi_ten(r.get("ten", "")), []
    for tm in sorted(theo_tm, key=lambda t: (lech(t), t)):
        w, p, _ = sxb_tool.chon_file(theo_tm[tm], key)
        if w or p:
            kq.append((p, w))
        if len(kq) >= toi_da:
            break
    return kq


_LIEN_QUAN_RE = re.compile(r"^\s*(\d{4})\s*-\s*dòng\s*(\d+)\s*-", re.I)


# Ghép PDF scan (CAN_OCR) vào sách theo tên OCR: tên KHỚP đúng một sách (trùng tên nhiều sách thì ưu tiên sách trong
# cột 'Sách liên quan', rồi sách cùng năm với năm OCR) -> dòng sách đó, không thì None + lý do
def ghep_scan(s, ten_map):
    o = s.get("_o") or {}
    ung = ten_map.get(sxb_tool.loi_ten(o.get("ten_sach") or ""), []) if o.get("ten_sach") else []
    if len(ung) > 1:
        lq = {m.groups() for m in map(_LIEN_QUAN_RE.match, str(s.get("sach") or "").splitlines()) if m}
        ung = [r for r in ung if (str(r["nam_excel"]), str(r["dong"])) in lq] or ung
        if len(ung) > 1 and o.get("nam"):
            ung = [r for r in ung if str(r["nam_excel"]) == str(o["nam"])] or ung
    if len(ung) == 1:
        return ung[0], ""
    return None, (f"tên OCR trùng {len(ung)} sách trong Excel" if ung else
                  "OCR không ra tên" if not o.get("ten_sach") else "tên OCR không KHỚP sách nào trong Excel")


# Ghép kết quả OCR vào dòng đối chiếu: thời gian OCR quyết định thư mục Năm/thang MM và phần mmyy của tên mới
def ap_dung_ocr(r, o, theo_ten="ocr"):
    o = o or {}
    # tên Word / PDF và tên trên bìa (nếu đã OCR bìa): lấy tên khớp tên Excel hơn (KHỚP > KHÁC > không ra)
    hang = {"KHỚP": 2, "KHÁC": 1, "": 0}
    ung = [(o.get("ten_sach") or "", o.get("nguon_ten") == "bia"), (o.get("ten_bia") or "", True)]
    ung = [(so_ten(r["ten"], t), t, la_bia) for t, la_bia in ung if t]
    khop, ten, tu_bia = max(ung, key=lambda x: hang[x[0]]) if ung else ("", "", False)
    r.update(ten_excel=r["ten"], nam_excel=r["nam"], ten_ocr=ten, ten_khop=khop)
    if tu_bia and o.get("ten_sach") and o.get("nguon_ten") != "bia":
        r["ghi_chu"] = "; ".join(x for x in (r.get("ghi_chu", ""), f"tên lấy theo bìa (Word / PDF ra '{o['ten_sach']}')") if x)
    thang, nam = o.get("thang"), o.get("nam")
    if nam and not NAM_HOP_LE[0] <= int(nam) <= NAM_HOP_LE[1]:
        thang = nam = None
    r["thang_lc"], r["nam_lc"] = thang or "", nam or ""
    if nam:
        r["nam"], r["qdxb"] = int(nam), f"{int(thang):02d}/{nam}" if thang else str(nam)
        r["nguon_ngay"] = ("OCR (QĐXB)" if o.get("theo") == "qdxb" else "OCR") if thang else "OCR (chỉ năm)"
    else:  # không dùng QĐXB của Excel: sách này không chép, để xem lại (tab 5)
        r["qdxb"], r["nguon_ngay"] = "", KHONG_RA_NGAY
    dung_ocr = theo_ten == "ocr" and bool(r["ten_ocr"])
    r["nguon_ten"] = ("OCR (bìa)" if tu_bia else "OCR") if dung_ocr else "Excel"
    r["ten_moi"] = sxb_tool.ten_file_moi(r["ten_ocr"] if dung_ocr else r["ten"], r["qdxb"], r["bbt"], r["lan"])


def _them_ghi_chu(r, s):
    r["ghi_chu"] = "; ".join(x for x in (r.get("ghi_chu", ""), s) if x)


# Chạy cả bước: OCR các sách KHỚP -> lập kế hoạch chép -> chép (thu=True: chỉ lập kế hoạch) -> ghi Excel
def chay_ocr_copy(dc, dich, ra, api=API_MAC_DINH, so_luong=2, timeout=TIMEOUT_MAC_DINH, theo_ten="excel",
                  di_chuyen=False, thu=False, cache="", bbt="", log=print, bao=lambda *a: None, bo_khac=True,
                  lay_lech=True, lay_scan=True):
    dung, khoa = sxb_tool.DUNG, threading.Lock()
    os.makedirs(ra, exist_ok=True)
    if cache:
        os.makedirs(os.path.dirname(cache) or ".", exist_ok=True)
    tat_ca = sxb_tool.doc_doi_chieu(dc)
    kq = [r for r in tat_ca if r.get("tinh_trang") == "KHỚP"]
    # Sách LỆCH NĂM (sheet CAN_XU_LY): chưa chọn file -> thử tối đa TOI_DA_UNG_VIEN cặp file ruột đến khi tên OCR KHỚP
    lech = [r for r in tat_ca if lay_lech and str(r.get("tinh_trang") or "").startswith("LỆCH NĂM")]
    for r in lech:
        r.update(_lech=r["tinh_trang"], tinh_trang="KHỚP", file_pdf="", file_word="", _ung_vien=ung_vien_lech(r))
    # PDF scan (sheet CAN_OCR): OCR từng file, ghép vào sách có tên KHỚP; không khớp sách nào -> CAN_KIEM_TRA/KHONG_KHOP_SACH
    quet = {}
    for s in (sxb_tool.doc_can_ocr(dc) if lay_scan else []):
        p = str(s.get("path") or "").strip()
        if p and p not in quet:
            quet[p] = dict(s, path=p, pdf=p if os.path.isfile(sxb_tool.duong_dan_dai(p)) else "", dich="", ghi_chu="")
    sxb_tool.MA_BBT.clear()
    sxb_tool.MA_BBT.update(sxb_tool.doc_ma_bbt(bbt) if bbt and os.path.isfile(bbt) else {})
    log(f"{len(kq)} sách KHỚP, {len(lech)} sách LỆCH NĂM, {len(quet)} PDF scan (CAN_OCR) trong {os.path.basename(dc)}; "
        "mã BBT: " + (f"{len(sxb_tool.MA_BBT)} mã từ {bbt}" if sxb_tool.MA_BBT else "chưa có, dùng BBT bỏ dấu"))
    kq += lech

    # Mỗi cặp (PDF, Word) chỉ OCR một lần dù nhiều dòng Excel cùng trỏ tới
    cap = {}
    for r in kq:
        r["dich"] = ""
        con = lambda p: p if p and os.path.isfile(sxb_tool.duong_dan_dai(p)) else ""
        pdf, word = con(r["file_pdf"]), con(r["file_word"])
        for p, moi in ((r["file_pdf"], pdf), (r["file_word"], word)):
            if p and not moi:
                _them_ghi_chu(r, f"không còn file {os.path.basename(p)}")
        r["file_pdf"], r["file_word"] = pdf, word
        r["_khoa"] = _khoa_ocr(pdf, word) if pdf or word else ""
        if r["_khoa"]:
            cap.setdefault(r["_khoa"], (pdf, word))
    da = doc_cache_ocr(cache) if cache else {}
    can = [k for k in cap if k not in da]
    log(f"{len(cap)} cuốn cần OCR: {len(cap) - len(can)} đã có kết quả từ lần trước, {len(can)} gửi máy chủ")

    from concurrent.futures import ThreadPoolExecutor
    xong, t0 = [0], time.time()

    def mot(k):
        if dung.is_set():
            return
        pdf, word = cap[k]
        o = ocr_mot(api, pdf, word, timeout, dung)
        with khoa:
            da[k] = o
            xong[0] += 1
            if cache and not o["loi"]:
                with open(cache, "a", encoding="utf-8") as f:
                    f.write(json.dumps(o, ensure_ascii=False) + "\n")
            log(f"[{xong[0]}/{len(can)}] {os.path.basename(pdf or word)} → {o.get('ten_sach') or '-'} | "
                f"{o.get('thang') or '--'}/{o.get('nam') or '----'}" + (f" | LỖI {o['loi']}" if o["loi"] else ""))
            bao("OCR", xong[0], len(can))

    with ThreadPoolExecutor(max_workers=max(1, so_luong)) as ex:
        list(ex.map(mot, can))

    # OCR một cặp file, có trong cache (không lỗi) thì dùng lại
    def lay_ocr(pdf, word):
        k = _khoa_ocr(pdf, word)
        with khoa:
            o = da.get(k)
        if o is not None and not o.get("loi"):
            return k, o
        o = ocr_mot(api, pdf, word, timeout, dung)
        with khoa:
            da[k] = o
            if cache and not o["loi"]:
                with open(cache, "a", encoding="utf-8") as f:
                    f.write(json.dumps(o, ensure_ascii=False) + "\n")
        return k, o

    # Sách LỆCH NĂM: OCR lần lượt các cặp ứng viên, dừng ở cặp có tên KHỚP tên Excel; không cặp nào khớp thì lấy cặp đầu
    # (không lỗi) để bước sau OCR bìa / ghi CẦN XEM. File liên quan thu về thư mục của cặp đã chọn
    xong_lech = [0]

    def thu_lech(r):
        if dung.is_set():
            return
        chon, i, khop = None, 0, False
        for i, (pdf, word) in enumerate(r["_ung_vien"], 1):
            k, o = lay_ocr(pdf, word)
            if chon is None or (chon[3].get("loi") and not o.get("loi")):
                chon = (pdf, word, k, o)
            if not o.get("loi") and so_ten(r["ten"], o.get("ten_sach")) == "KHỚP":
                chon, khop = (pdf, word, k, o), True
                break
            if dung.is_set():
                break
        with khoa:
            xong_lech[0] += 1
            if chon:
                pdf, word, k, o = chon
                tm = os.path.dirname(pdf or word)
                r.update(file_pdf=pdf, file_word=word, _khoa=k, cac_file=sxb_tool._ds_file(
                    x for x in sxb_tool._dong_file(r["cac_file"]) if os.path.dirname(x[1]) == tm
                    or os.path.dirname(x[1]).startswith(tm + os.sep)))
                _them_ghi_chu(r, f"{r['_lech']}: thử {i}/{len(r['_ung_vien'])} cặp file, chọn {os.path.basename(pdf or word)}"
                                 + (" (tên KHỚP)" if khop else " (chưa cặp nào tên KHỚP)"))
            else:
                _them_ghi_chu(r, f"{r['_lech']}: không có file Word / PDF ruột")
            log(f"[lệch năm {xong_lech[0]}/{len(lech)}] {r['ten'][:50]} → "
                + (f"{os.path.basename(chon[0] or chon[1])} | {chon[3].get('ten_sach') or '-'}"
                   + (" | KHỚP" if khop else "") if chon else "không có file"))
            bao("OCR lệch năm", xong_lech[0], len(lech))

    if lech:
        log(f"{len(lech)} sách LỆCH NĂM: thử tối đa {TOI_DA_UNG_VIEN} cặp file / sách đến khi tên OCR KHỚP")
        with ThreadPoolExecutor(max_workers=max(1, so_luong)) as ex:
            list(ex.map(thu_lech, lech))

    # PDF scan (CAN_OCR): OCR tên + thời gian từng file
    xong_scan, ds_scan = [0], [s for s in quet.values() if s["pdf"]]

    def ocr_scan(s):
        if dung.is_set():
            return
        s["_khoa"], s["_o"] = lay_ocr(s["pdf"], "")
        with khoa:
            xong_scan[0] += 1
            o = s["_o"]
            log(f"[scan {xong_scan[0]}/{len(ds_scan)}] {os.path.basename(s['pdf'])} → {o.get('ten_sach') or '-'} | "
                f"{o.get('thang') or '--'}/{o.get('nam') or '----'}" + (f" | LỖI {o['loi']}" if o.get("loi") else ""))
            bao("OCR PDF scan", xong_scan[0], len(ds_scan))

    if ds_scan:
        with ThreadPoolExecutor(max_workers=max(1, so_luong)) as ex:
            list(ex.map(ocr_scan, ds_scan))

    # Tên từ Word / PDF không KHỚP tên Excel (khác hoặc không ra tên) hoặc sách chỉ có bìa -> OCR file bìa nếu có
    # (bỏ tick 'chỉ chép sách KHỚP' thì chỉ OCR bìa khi không ra tên); mỗi khoá thử một lần, lưu cache
    bia = {}
    for r in kq:
        k = r["_khoa"]
        o = da.get(k) if k else None
        ten = (o or {}).get("ten_sach")
        if k and (o is None or o.get("thu_bia") or (ten and (not bo_khac or so_ten(r["ten"], ten) == "KHỚP"))):
            continue
        p = chon_bia(ds_bia_chat(r))
        if not p:
            continue
        if not k:
            k = r["_khoa_bia"] = f"bia|{p}"
            if da.get(k, {}).get("thu_bia"):
                continue
        bia.setdefault(k, p)
    if bia:
        log(f"{len(bia)} cuốn tên Word / PDF không khớp Excel (hoặc không ra tên) → OCR file bìa")
    xong_bia = [0]

    def mot_bia(k):
        if dung.is_set():
            return
        p = bia[k]
        try:
            ten, loi = ocr_bia(api, p, timeout, dung), ""
        except Exception as e:
            ten, loi = "", f"{type(e).__name__}: {e}"
        with khoa:
            o = dict(da.get(k) or {"khoa": k, "loi": ""})
            o["bia"] = p
            if ten:
                o["ten_bia"] = ten  # tên trên bìa: so riêng với tên Excel, lấy tên khớp hơn (ap_dung_ocr)
            if not loi:
                o["thu_bia"] = True
            da[k] = o
            xong_bia[0] += 1
            if cache and not loi and not o.get("loi"):
                with open(cache, "a", encoding="utf-8") as f:
                    f.write(json.dumps(o, ensure_ascii=False) + "\n")
            log(f"[bìa {xong_bia[0]}/{len(bia)}] {os.path.basename(p)} → {ten or '-'}" + (f" | LỖI {loi}" if loi else ""))
            bao("OCR bìa", xong_bia[0], len(bia))

    with ThreadPoolExecutor(max_workers=max(1, so_luong)) as ex:
        list(ex.map(mot_bia, bia))

    for r in kq:
        o = da.get(r["_khoa"]) or da.get(r.get("_khoa_bia", ""))
        ap_dung_ocr(r, o, theo_ten)
        if o and o.get("loi"):
            _them_ghi_chu(r, f"OCR lỗi ({o['loi']})")
        elif r["_khoa"] and not o:
            _them_ghi_chu(r, "chưa OCR (đã dừng)")

    # Ghép PDF scan vào sách: sách đã có file tên KHỚP thì scan là bản trùng (không chép); sách chưa KHỚP thì dùng scan làm
    # file ruột (OCR của scan quyết định tên, thời gian); không ghép được -> copy vào CAN_KIEM_TRA/KHONG_KHOP_SACH
    if quet:
        ten_map = {}
        for r in kq:
            ten_map.setdefault(sxb_tool.loi_ten(r["ten"]), []).append(r)
        for s in quet.values():
            o = s.get("_o") or {}
            s.update(ten_ocr=o.get("ten_sach") or "", thang=o.get("thang") or "", nam=o.get("nam") or "", ghep="")
            if not s["pdf"]:
                s["trang_thai"], s["ghi_chu"] = "KHÔNG CÒN FILE", "không thấy file trên ổ"
                continue
            if "_o" not in s:
                s["trang_thai"] = "CHƯA OCR"
                continue
            if o.get("loi"):
                s["trang_thai"], s["ghi_chu"] = "OCR LỖI", o["loi"]
                continue
            r, ly_do = ghep_scan(s, ten_map)
            if r is None:
                s["trang_thai"], s["ghi_chu"] = "KHÔNG KHỚP SÁCH", ly_do
                continue
            s["ghep"] = f"{r['nam_excel']} - dòng {r['dong']} - {r['ten']}"
            if r["ten_khop"] == "KHỚP" and r["_khoa"] and r["_khoa"] != s["_khoa"]:
                s["trang_thai"], s["ghi_chu"] = "TRÙNG", "sách đã có file khác tên KHỚP: không chép bản scan"
                continue
            moc = dt.datetime.fromtimestamp(os.path.getmtime(sxb_tool.duong_dan_dai(s["pdf"]))).strftime("%Y-%m-%d %H:%M:%S")
            r.update(file_pdf=s["pdf"], file_word="", _khoa=s["_khoa"], tinh_trang="KHỚP", nam=r["nam_excel"],
                     cac_file=sxb_tool._ds_file([(moc, s["pdf"])]),
                     ghi_chu="; ".join(x for x in str(r.get("ghi_chu") or "").split("; ") if x and not x.startswith("OCR lỗi")))
            ap_dung_ocr(r, o, theo_ten)
            _them_ghi_chu(r, f"file ruột là PDF scan (CAN_OCR): {os.path.basename(s['pdf'])}")
            s["trang_thai"] = "ĐÃ GHÉP"
        dem_scan = {}
        for s in quet.values():
            dem_scan[s["trang_thai"]] = dem_scan.get(s["trang_thai"], 0) + 1
        log("PDF scan: " + ", ".join(f"{k} {v}" for k, v in sorted(dem_scan.items())))
    # Sắp theo thời gian mới (Năm, tháng) giống thứ tự thư mục đích
    kq.sort(key=lambda r: (int(r["nam"] or 0), sxb_tool.thang_qdxb(r["qdxb"]), r["ten_moi"]))
    for i, r in enumerate(kq, 1):
        r["stt"] = i

    tem = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    xlsx = os.path.join(ra, f"ocr_copy_{tem}.xlsx")
    if dung.is_set():
        for r in kq:
            r["trang_thai"] = "CHƯA CHÉP"
        ghi_excel("copy", kq, xlsx)
        ghi_sheet_scan(xlsx, quet)
        log(f"ĐÃ DỪNG sau {xong[0]}/{len(can)} cuốn OCR, chưa chép file nào. Chạy lại sẽ OCR tiếp phần còn lại.")
        return xlsx, {"ocr": xong[0], "chep": 0, "co_san": 0, "loi": []}

    dem = {k: sum(r["ten_khop"] == k for r in kq) for k in ("KHỚP", "KHÁC")}
    log(f"Tên Excel / OCR: {dem['KHỚP']} khớp, {dem['KHÁC']} khác, "
        f"{len(kq) - sum(dem.values())} OCR không ra tên")
    # Sách không chép vào thư mục năm mà COPY toàn bộ file liên quan vào <đích>/CAN_KIEM_TRA/<loại>/ để xem lại ngày tháng:
    # - TRUOC_<năm>: năm lưu chiểu OCR trước NAM_BAT_DAU (ngoài khoảng đang chuẩn hoá)
    # - KHONG_NGAY: có Word / PDF ruột mà cả hai không ra dòng 'In xong và nộp lưu chiểu' (OCR lỗi thì không tính: chạy lại)
    truoc = [r for r in kq if 0 < _nam_luu(r) < NAM_BAT_DAU]
    ko_ngay = [r for r in kq if r["nguon_ngay"] == KHONG_RA_NGAY and r["_khoa"] and "OCR lỗi" not in r.get("ghi_chu", "")]
    la_truoc, la_ko_ngay = {id(r) for r in truoc}, {id(r) for r in ko_ngay}
    # Chỉ chép sách có tên OCR (Word / PDF hoặc bìa) KHỚP tên Excel; còn lại (khác hoặc không ra tên) nhiều khả năng
    # tab 2/3 chọn nhầm file -> không chép, CẦN XEM. Sách không có file ruột / OCR lỗi để trạng thái riêng
    xem = [r for r in kq if bo_khac and r["_khoa"] and r["ten_khop"] != "KHỚP" and "OCR lỗi" not in r.get("ghi_chu", "")
           and id(r) not in la_truoc | la_ko_ngay]
    la_xem = {id(r) for r in xem}
    if xem:
        from collections import Counter
        dem_xem = Counter(r["ten_khop"] or "không ra tên" for r in xem)
        log(f"{len(xem)} sách tên OCR không KHỚP tên Excel (kể cả bìa) không chép, trạng thái CẦN XEM: "
            + ", ".join(f"{k} {v}" for k, v in dem_xem.most_common()))
    con_ko_ngay = sum(r["nguon_ngay"] == KHONG_RA_NGAY and id(r) not in la_ko_ngay | la_xem for r in kq)
    if con_ko_ngay:
        log(f"{con_ko_ngay} sách không ra ngày (OCR lỗi / không có Word, PDF ruột): không chép (không dùng ngày Excel)")
    # ke_hoach_luu của sxb_tool chọn bìa bằng sxb_tool.ds_bia: thay tạm bằng bản chặt hơn
    goc_ds_bia, sxb_tool.ds_bia = sxb_tool.ds_bia, ds_bia_chat
    try:
        kh = sxb_tool.ke_hoach_luu([r for r in kq if id(r) not in la_truoc | la_xem
                                    and r["nguon_ngay"] != KHONG_RA_NGAY], dich)
    finally:
        sxb_tool.ds_bia = goc_ds_bia
    log(f"{len(kh)} file sẽ {'di chuyển' if di_chuyen else 'chép'} sang {dich}")
    goc_kt = os.path.join(dich, THU_MUC_KIEM_TRA)
    can_kt = truoc + ko_ngay
    dong_kt, kh_kt = ke_hoach_kiem_tra(
        [dict(r, trang_thai=TRANG_THAI_TRUOC if id(r) in la_truoc else KHONG_RA_NGAY) for r in can_kt],
        [dict(r, nam=r["nam_excel"]) for r in can_kt], goc_kt, [MA_TRUOC, MA_KHONG_NGAY])
    kh_kt = [(x, s, p) for x, s, p in kh_kt if os.path.isfile(sxb_tool.duong_dan_dai(s))]
    for ma, ds in ((MA_TRUOC, truoc), (MA_KHONG_NGAY, ko_ngay)):
        if ds:
            log(f"{len(ds)} sách → {os.path.join(goc_kt, ma)} (copy {sum(x['ma'] == ma for x, _, _ in kh_kt)} file liên quan), "
                + ("năm lưu chiểu OCR trước " + str(NAM_BAT_DAU) if ma == MA_TRUOC else "Word và PDF không ra dòng In xong"))
    theo_dong = {(str(r["nam_excel"]), str(r["dong"])): r for r in can_kt}
    # PDF scan không ghép được sách nào: copy vào CAN_KIEM_TRA/KHONG_KHOP_SACH/<năm OCR>/<tên file>_<mã>/ kèm _thong_tin.txt
    kh_scan = []
    for s in quet.values():
        if s.get("trang_thai") == "KHÔNG KHỚP SÁCH":
            goc_ten = os.path.splitext(os.path.basename(s["pdf"]))[0]
            ma = hashlib.md5(s["pdf"].encode("utf-8")).hexdigest()[:6]  # 2 file cùng tên ở 2 thư mục không đè nhau
            s["thu_muc"] = os.path.join(goc_kt, MA_KHONG_KHOP_SACH, str(s["nam"] or "khong_nam"), f"{_ten_thu_muc(goc_ten)}_{ma}")
            kh_scan.append((s, s["pdf"], os.path.join(s["thu_muc"], os.path.basename(s["pdf"]))))
    if kh_scan:
        log(f"{len(kh_scan)} PDF scan không ghép được sách → copy vào {os.path.join(goc_kt, MA_KHONG_KHOP_SACH)}")
    chep, co_san, loi = 0, 0, []
    if thu:
        for r, src, p in kh:
            r["dich"] = "\n".join(x for x in (r["dich"], p) if x)
        for x, src, p in kh_kt:
            r = theo_dong[(str(x["nam_excel"]), str(x["dong"]))]
            r["dich"] = "\n".join(v for v in (r["dich"], p) if v)
        for s, src, p in kh_scan:
            s["dich"] = p
        log("\n".join(f"{s}\n   → {p}" for _, s, p in (kh + kh_kt + kh_scan)[:300])
            + ("\n..." if len(kh) + len(kh_kt) + len(kh_scan) > 300 else "") + "\n(chạy thử - chưa chép)")
    else:
        # Nhật ký nguồn -> đích các file sắp chép / di chuyển (bỏ file đã có sẵn), ghi trước để Hoàn tác được cả khi tool tắt giữa chừng
        nk = ghi_nhat_ky(os.path.join(ra, f"nhat_ky_chep_{tem}.csv"), di_chuyen,
                         [(s, p) for _, s, p in kh if not sxb_tool.da_co(s, p)],
                         [(s, p) for _, s, p in kh_kt + kh_scan if not sxb_tool.da_co(s, p)])
        log(f"Nhật ký chép (dùng cho Hoàn tác): {nk}")
        chep, co_san, loi = sxb_tool.luu_file(kh, di_chuyen, lambda i, xb, tb: bao("Chép file", i, len(kh)))
        if kh_scan and not sxb_tool.DUNG.is_set():
            for s, _, _ in kh_scan:
                _ghi_thong_tin_scan(s)
            c3, s3, l3 = sxb_tool.luu_file(kh_scan, False, lambda i, xb, tb: bao("Chép PDF scan", i, len(kh_scan)))
            chep, co_san, loi = chep + c3, co_san + s3, loi + l3
        if kh_kt and not sxb_tool.DUNG.is_set():
            for x in dong_kt:
                _ghi_thong_tin(x)
            c2, s2, l2 = sxb_tool.luu_file(kh_kt, False, lambda i, xb, tb: bao("Chép cần kiểm tra", i, len(kh_kt)))
            chep, co_san, loi = chep + c2, co_san + s2, loi + l2
            for x in dong_kt:
                r = theo_dong[(str(x["nam_excel"]), str(x["dong"]))]
                r["dich"] = x["dich"]
                if "LỖI lưu" in x["ghi_chu"]:
                    _them_ghi_chu(r, x["ghi_chu"])
                x["trang_thai"] = "ĐÃ CHÉP" if x["dich"] else "KHÔNG CÓ FILE"
            ghi_excel_kiem_tra(dong_kt, os.path.join(goc_kt, f"can_kiem_tra_tab4_{tem}.xlsx"), [MA_TRUOC, MA_KHONG_NGAY])
    for r in truoc:
        _them_ghi_chu(r, f"năm OCR {r['nam_lc']} trước {NAM_BAT_DAU}: không chép vào đích, copy vào {THU_MUC_KIEM_TRA}/{MA_TRUOC}")
    for r in ko_ngay:
        _them_ghi_chu(r, f"Word và PDF không ra dòng In xong: copy vào {THU_MUC_KIEM_TRA}/{MA_KHONG_NGAY}")
    for r in kq:
        r["trang_thai"] = (TRANG_THAI_TRUOC if id(r) in la_truoc else
                           KHONG_RA_NGAY if id(r) in la_ko_ngay else
                           "CẦN XEM" if id(r) in la_xem else
                           "KHÔNG CÓ FILE" if not r["_khoa"] and not r["dich"] else
                           KHONG_RA_NGAY if r["nguon_ngay"] == KHONG_RA_NGAY else
                           "LỖI CHÉP" if "LỖI lưu" in r.get("ghi_chu", "") else
                           "CHẠY THỬ" if thu else "ĐÃ LƯU" if r["dich"] else "CHƯA CHÉP")
    ghi_excel("copy", kq, xlsx)
    ghi_sheet_scan(xlsx, quet)
    log(f"Xong: OCR {xong[0]} cuốn, {'chạy thử' if thu else f'chép {chep} file, {co_san} đã có sẵn, {len(loi)} lỗi'}"
        f" - {time.time() - t0:.0f}s\nKết quả: {xlsx}")
    return xlsx, {"ocr": xong[0], "chep": chep, "co_san": co_san, "loi": loi}


# Giao diện

# Đọc cấu hình
def doc_cau_hinh():
    try:
        with open(CAU_HINH, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


# Lưu cấu hình
def ghi_cau_hinh(ch):
    try:
        with open(CAU_HINH, "w", encoding="utf-8") as f:
            json.dump(ch, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


# Hoàn tác: theo nhật ký chép (nhat_ky_chep_*.csv) hoặc Excel kết quả ocr_copy_*.xlsx (lần chạy chưa có nhật ký)

COT_NHAT_KY = ["Thao tác", "Nguồn", "Đích"]


# Ghi nhật ký nguồn -> đích
def ghi_nhat_ky(path, di_chuyen, cap, cap_copy=()):
    import csv
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(COT_NHAT_KY)
        w.writerows(("move" if di_chuyen else "copy", s, p) for s, p in cap)
        w.writerows(("copy", s, p) for s, p in cap_copy)  # luôn copy (vd sách vào CAN_KIEM_TRA)
    return path


# Kế hoạch hoàn tác -> [(việc, nguồn, đích, lý do)]; việc: "xoá" (bản chép), "trả về" (file đã di chuyển), "bỏ qua"
def ke_hoach_hoan_tac(path):
    import csv
    dd = sxb_tool.duong_dan_dai
    co = lambda p: bool(p) and os.path.exists(dd(p))
    kh = []
    if path.lower().endswith(".csv"):
        with open(path, encoding="utf-8-sig", newline="") as f:
            rows = [r for r in list(csv.reader(f))[1:] if len(r) >= 3]
        for tt, src, dst in (r[:3] for r in rows):
            if not co(dst):
                kh.append(("bỏ qua", src, dst, "đích không còn"))
            elif tt == "move":
                kh.append(("bỏ qua", src, dst, "chỗ cũ đã có file") if co(src) else ("trả về", src, dst, ""))
            elif not co(src):
                kh.append(("bỏ qua", src, dst, "nguồn không còn: giữ bản ở đích"))
            elif os.path.getsize(dd(src)) != os.path.getsize(dd(dst)):
                kh.append(("bỏ qua", src, dst, "bản ở đích khác nguồn (đã sửa?)"))
            else:
                kh.append(("xoá", src, dst, ""))
        return kh
    # Excel kết quả: chỉ đúng với lần chạy COPY; dòng nào file ruột nguồn đã mất (bị di chuyển) thì giữ nguyên cả dòng
    for r in sxb_tool._doc_sheet(path, LOAI["copy"][1], [(k, h) for k, h, _ in COT_OCR_COPY]):
        if r.get("trang_thai") == "CHẠY THỬ":
            continue
        nguon = {os.path.splitext(p)[1].lower(): p for p in (str(r.get("file_word") or ""), str(r.get("file_pdf") or "")) if p}
        dich = [p for p in str(r.get("dich") or "").splitlines() if p]
        mat = any(os.path.basename(os.path.dirname(d)) != "Bia" and not co(nguon.get(os.path.splitext(d)[1].lower(), ""))
                  for d in dich)
        for dst in dich:
            src = "" if os.path.basename(os.path.dirname(dst)) == "Bia" else nguon.get(os.path.splitext(dst)[1].lower(), "")
            if not co(dst):
                kh.append(("bỏ qua", src, dst, "đích không còn"))
            elif mat:
                kh.append(("bỏ qua", src, dst, "file ruột nguồn không còn (lần đó di chuyển?): giữ bản ở đích"))
            else:
                kh.append(("xoá", src, dst, ""))
    return kh


# Thực hiện hoàn tác, xoá thư mục rỗng còn lại (Bia -> thư mục sách -> thang MM -> Năm, không đi cao hơn)
def hoan_tac(kh, log=print):
    import shutil
    dd = sxb_tool.duong_dan_dai
    xoa, tra, loi, thu_muc = 0, 0, [], set()
    for viec, src, dst, _ in kh:
        thu_muc.add(os.path.dirname(dst))  # cả dòng bỏ qua: thư mục có thể đã tạo rồi để trống
        if viec == "bỏ qua":
            continue
        try:
            if viec == "xoá":
                os.remove(dd(dst)); xoa += 1
            else:
                os.makedirs(dd(os.path.dirname(src)), exist_ok=True)
                shutil.move(dd(dst), dd(src)); tra += 1
        except OSError as e:
            loi.append(f"{dst}: {e}")
    xoa_tm = 0
    for d in sorted(thu_muc, key=len, reverse=True):
        for _ in range(4 if os.path.basename(d) == "Bia" else 3):
            try:
                os.rmdir(dd(d))  # chỉ xoá được thư mục rỗng
                xoa_tm += 1
            except OSError:
                break
            d = os.path.dirname(d)
    log(f"Hoàn tác xong: xoá {xoa} bản chép, trả {tra} file về chỗ cũ, xoá {xoa_tm} thư mục rỗng, {len(loi)} lỗi"
        + "".join(f"\n  LỖI {x}" for x in loi[:50]))
    return xoa, tra, xoa_tm, loi


# Tab 5. Cần kiểm tra: từ Excel kết quả tab 4 chép toàn bộ file của các sách cần xem vào CAN_KIEM_TRA/<loại>/
THU_MUC_KIEM_TRA = "CAN_KIEM_TRA"
_ten_ocr = lambda r: str(r.get("ten_ocr") or "").strip()
NAM_BAT_DAU = int(os.getenv("SXB_NAM_BAT_DAU", "2006"))  # dữ liệu đang chuẩn hoá từ năm này (năm này vẫn hợp lệ)
MA_TRUOC = f"TRUOC_{NAM_BAT_DAU}"
TRANG_THAI_TRUOC = f"TRƯỚC {NAM_BAT_DAU}"  # tab 4: năm OCR trước mốc, không chép vào đích, copy vào CAN_KIEM_TRA
MA_KHONG_NGAY = "KHONG_NGAY"  # tab 4: Word và PDF không ra dòng In xong, copy vào CAN_KIEM_TRA/KHONG_NGAY
MA_KHONG_KHOP_SACH = "KHONG_KHOP_SACH"  # tab 4: PDF scan (CAN_OCR) không ghép được sách nào


# Không ra ngày lưu chiểu (Excel cũ: ngày theo QĐXB Excel); trừ sách không có file ruột và OCR lỗi (chạy lại)
def _khong_ngay(r):
    return ((r.get("trang_thai") == KHONG_RA_NGAY or r.get("nguon_ngay") in (KHONG_RA_NGAY, "Excel (QĐXB)"))
            and r.get("trang_thai") != "KHÔNG CÓ FILE" and "OCR lỗi" not in str(r.get("ghi_chu") or ""))


# Năm sách được lưu: năm lưu chiểu OCR; Excel cũ (trước khi bỏ ngày Excel) ghi 'Excel (QĐXB)' thì năm Excel; 0 nếu không rõ
def _nam_luu(r):
    nam = r.get("nam_lc") or (r.get("nam_excel") if r.get("nguon_ngay") == "Excel (QĐXB)" else "")
    try:
        return int(str(nam or "").strip())
    except ValueError:
        return 0


# (thư mục, tên loại, giải thích, điều kiện, mặc định chọn); một sách chỉ vào loại đầu tiên khớp (TRUOC_<năm> xét trước), loại khác ghi ở "Lý do khác"
LOAI_KIEM_TRA = [
    (MA_TRUOC, f"Năm trước {NAM_BAT_DAU}",
     f"Năm lưu chiểu OCR trước {NAM_BAT_DAU}: ngoài khoảng đang chuẩn hoá, cần xem lại ngày tháng (xét trước mọi loại)",
     lambda r: 0 < _nam_luu(r) < NAM_BAT_DAU, True),
    (MA_KHONG_NGAY, "Không ra ngày", "Word và PDF không ra dòng 'In xong và nộp lưu chiểu' (không dùng ngày Excel), "
     "không chép vào thư mục năm", _khong_ngay, True),
    ("1_TEN_OCR_KHAC", "Tên OCR không khớp Excel", "CẦN XEM: tên OCR (Word / PDF / bìa) không KHỚP tên Excel (khác hoặc không ra tên), chưa chép (nghi tab 2/3 chọn nhầm file)",
     lambda r: r.get("trang_thai") == "CẦN XEM", True),
    ("2_KHONG_CO_FILE_RUOT", "Không có file ruột", "KHÔNG CÓ FILE: không có Word / PDF ruột (chỉ có bìa hoặc file đã mất), chưa chép",
     lambda r: r.get("trang_thai") == "KHÔNG CÓ FILE", True),
    ("3_LOI_CHEP", "Lỗi chép", "LỖI CHÉP: chép sang thư mục đích bị lỗi",
     lambda r: r.get("trang_thai") == "LỖI CHÉP", True),
    ("4_OCR_LOI", "OCR lỗi", "Máy chủ OCR lỗi, nên chạy lại OCR",
     lambda r: "OCR lỗi" in str(r.get("ghi_chu") or ""), True),
    ("5_OCR_KHONG_RA_TEN", "OCR không ra tên", "Đã chép theo tên Excel nhưng OCR (Word, PDF, bìa) không đọc ra tên sách",
     lambda r: not _ten_ocr(r), True),
    ("6_NAM_OCR_KHAC_EXCEL", "Năm OCR khác Excel", "Năm lưu chiểu OCR khác năm trong Excel: sách được xếp theo năm OCR",
     lambda r: bool(r.get("nam_lc")) and bool(r.get("nam_excel")) and str(r["nam_lc"]) != str(r["nam_excel"]), False),
    ("7_CHI_CO_NAM", "OCR chỉ ra năm", "OCR chỉ đọc được năm, sách nằm ở thang 00",
     lambda r: r.get("nguon_ngay") == "OCR (chỉ năm)", False),
]
COT_KIEM_TRA = [("stt", "STT", 6), ("loai", "Loại", 22), ("ly_do_khac", "Lý do khác", 30), ("nam_excel", "Năm (Excel)", 8),
                ("dong", "Dòng Excel", 8), ("ten_excel", "Tên sách (Excel)", 40), ("ten_ocr", "Tên sách (OCR)", 40),
                ("ten_khop", "Tên khớp", 11), ("thang_lc", "Tháng lưu chiểu", 9), ("nam_lc", "Năm lưu chiểu", 9),
                ("nguon_ngay", "Thời gian theo", 14), ("trang_thai_goc", "Trạng thái tab 4", 14),
                ("ghi_chu_goc", "Ghi chú tab 4", 40), ("so_file", "Số file chép", 9), ("thu_muc", "Thư mục kiểm tra", 80),
                ("thu_muc_goc", "Thư mục gốc", 80), ("trang_thai", "Trạng thái", 14), ("ghi_chu", "Ghi chú", 40)]


# Đọc Excel kết quả tab 4 (ocr_copy_*.xlsx hoặc sheet 'Bảng gốc' của file thống kê)
def doc_ket_qua_copy(xlsx):
    cot = [(k, h) for k, h, _ in COT_OCR_COPY]
    for sheet in (LOAI["copy"][1], "Bảng gốc"):
        rows = [r for r in sxb_tool._doc_sheet(xlsx, sheet, cot) if r.get("ten_excel")]
        if rows:
            return rows
    raise ValueError(f"{os.path.basename(xlsx)} không có sheet '{LOAI['copy'][1]}' (kết quả tab 4)")


# Tên thư mục an toàn từ tên sách
def _ten_thu_muc(s, dai=60):
    s = re.sub(r"[^a-z0-9]+", "_", sxb_tool.khong_dau(str(s or "")).lower()).strip("_")
    return s[:dai].rstrip("_") or "khong_ten"


# Các file liên quan của sách: file cùng năm trong 'Tất cả file tìm được' + Word / PDF đã chọn; ca_khac_nam: lấy hết
def file_lien_quan(d, r, ca_khac_nam=False):
    ds = sxb_tool._dong_file((d or {}).get("cac_file"))
    nam = str((d or {}).get("nam") or r.get("nam_excel") or "")
    chon = [p for dm, p in ds if ca_khac_nam or dm[:4] == nam] or [p for _, p in ds]
    chon += [str(x) for x in (r.get("file_word"), r.get("file_pdf"), (d or {}).get("file_word"), (d or {}).get("file_pdf")) if x]
    return sorted(set(chon))


# Phân loại + lập kế hoạch chép -> (danh sách dòng, kế hoạch [(dòng, nguồn, đích)])
def ke_hoach_kiem_tra(ket_qua, doi_chieu, goc_ra, chon_loai=None, ca_khac_nam=False):
    chon_loai = set(chon_loai or [m for m, *_, md in LOAI_KIEM_TRA if md])
    dc = {(str(d.get("nam")), str(d.get("dong"))): d for d in doi_chieu}
    dong, kh, stt = [], [], 0
    for r in ket_qua:
        khop = [(m, ten) for m, ten, _, dk, _ in LOAI_KIEM_TRA if m in chon_loai and dk(r)]
        if not khop:
            continue
        stt += 1
        (ma, ten_loai), khac = khop[0], [t for _, t in khop[1:]]
        d = dc.get((str(r.get("nam_excel")), str(r.get("dong"))))
        files = file_lien_quan(d, r, ca_khac_nam)
        # Thư mục năm: theo năm Excel; riêng TRUOC_<năm> theo năm lưu (năm trước mốc, là lý do vào loại này)
        nam_tm = _nam_luu(r) if ma.startswith("TRUOC_") else r.get("nam_excel")
        # tên thư mục theo Năm Excel + Dòng Excel (cố định): tab 4 và tab 5 chép cùng một sách vào cùng một chỗ
        tm = os.path.join(goc_ra, ma, str(nam_tm or "khong_nam"),
                          f"{r.get('nam_excel') or 0}_{r.get('dong') or stt:0>4}_{_ten_thu_muc(r.get('ten_excel'))}")
        try:  # giữ cấu trúc thư mục con tính từ thư mục chung của các file
            goc = os.path.commonpath([os.path.dirname(p) for p in files]) if files else ""
        except ValueError:  # khác ổ đĩa
            goc = ""
        x = dict(stt=stt, loai=ten_loai, ma=ma, ly_do_khac="; ".join(khac), nam_excel=r.get("nam_excel"), dong=r.get("dong"),
                 ten_excel=r.get("ten_excel"), ten_ocr=_ten_ocr(r), ten_khop=r.get("ten_khop"), thang_lc=r.get("thang_lc"),
                 nam_lc=r.get("nam_lc"), nguon_ngay=r.get("nguon_ngay"), trang_thai_goc=r.get("trang_thai"),
                 ghi_chu_goc=r.get("ghi_chu"), so_file=len(files), thu_muc=tm, thu_muc_goc=goc, ghi_chu="",
                 dich="", trang_thai="KHÔNG CÓ FILE" if not files else "")
        if d is None:
            x["ghi_chu"] = "không thấy dòng này trong file đối chiếu"
        dong.append(x)
        for p in files:
            kh.append((x, p, os.path.join(tm, os.path.relpath(p, goc) if goc else os.path.basename(p))))
    return dong, kh


# Ghi _thong_tin.txt trong thư mục mỗi sách
def _ghi_thong_tin(x):
    os.makedirs(sxb_tool.duong_dan_dai(x["thu_muc"]), exist_ok=True)
    nd = [f"Loại: {x['loai']}" + (f" (kèm: {x['ly_do_khac']})" if x["ly_do_khac"] else ""),
          f"Năm / dòng Excel: {x['nam_excel']} / {x['dong']}", f"Tên sách (Excel): {x['ten_excel']}",
          f"Tên sách (OCR): {x['ten_ocr'] or '-'}  ({x['ten_khop'] or 'không ra tên'})",
          f"Lưu chiểu OCR: {x['thang_lc'] or '--'}/{x['nam_lc'] or '----'}  theo {x['nguon_ngay'] or '-'}",
          f"Trạng thái tab 4: {x['trang_thai_goc']}", f"Ghi chú tab 4: {x['ghi_chu_goc'] or '-'}",
          f"Thư mục gốc: {x['thu_muc_goc'] or '-'}"]
    with open(sxb_tool.duong_dan_dai(os.path.join(x["thu_muc"], "_thong_tin.txt")), "w", encoding="utf-8") as f:
        f.write("\n".join(nd) + "\n")


# _thong_tin.txt cho PDF scan không ghép được sách
def _ghi_thong_tin_scan(s):
    os.makedirs(sxb_tool.duong_dan_dai(s["thu_muc"]), exist_ok=True)
    nd = [f"PDF scan (sheet CAN_OCR): {s['path']}", f"Tên sách (OCR): {s.get('ten_ocr') or '-'}",
          f"Lưu chiểu OCR: {s.get('thang') or '--'}/{s.get('nam') or '----'}", f"Lý do không ghép: {s.get('ghi_chu') or '-'}",
          "Sách liên quan (tab 2):", *(f"  {x}" for x in str(s.get("sach") or "-").splitlines())]
    with open(sxb_tool.duong_dan_dai(os.path.join(s["thu_muc"], "_thong_tin.txt")), "w", encoding="utf-8") as f:
        f.write("\n".join(nd) + "\n")


COT_SCAN = [("stt", "STT", 6), ("path", "Đường dẫn", 70), ("trang", "Số trang", 8), ("ly_do", "Lý do OCR", 26),
            ("ten_ocr", "Tên sách (OCR)", 40), ("thang", "Tháng lưu chiểu", 9), ("nam", "Năm lưu chiểu", 9),
            ("sach", "Sách liên quan (tab 2)", 50), ("ghep", "Ghép với sách (năm - dòng - tên)", 50),
            ("trang_thai", "Trạng thái", 16), ("dich", "Đường dẫn đích", 70), ("ghi_chu", "Ghi chú", 40)]


# Thêm sheet 'PDF scan' (kết quả từng file của sheet CAN_OCR) vào Excel kết quả tab 4
def ghi_sheet_scan(xlsx, quet):
    if not quet:
        return
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter
    wb = load_workbook(xlsx)
    ws = wb.create_sheet("PDF scan")
    ws.append([h for _, h, _ in COT_SCAN])
    for c in ws[1]:
        c.font, c.alignment = Font(bold=True), Alignment(horizontal="center", wrap_text=True)
    wrap = Alignment(wrap_text=True, vertical="top")
    for i, s in enumerate(quet.values(), 1):
        ws.append([i if k == "stt" else s.get(k, "") for k, _, _ in COT_SCAN])
        for c in ws[ws.max_row]:
            c.alignment = wrap
    for i, (_, _, rong) in enumerate(COT_SCAN, 1):
        ws.column_dimensions[get_column_letter(i)].width = rong
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(xlsx)


# Excel danh sách: sheet Tổng hợp (số sách mỗi loại + giải thích thư mục) và mỗi loại một sheet
def ghi_excel_kiem_tra(dong, xlsx, chon_loai):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    ws = wb.active
    ws.title = "Tổng hợp"
    ws.append(["Thư mục", "Loại", "Số sách", "Số file", "Giải thích", "Đã chọn"])
    for ma, ten, giai_thich, _, _ in LOAI_KIEM_TRA:
        ds = [x for x in dong if x["ma"] == ma]
        ws.append([ma, ten, len(ds), sum(x["so_file"] for x in ds), giai_thich, "có" if ma in chon_loai else "không"])
    ws.append(["", "Tổng", len(dong), sum(x["so_file"] for x in dong), "Mỗi sách chỉ nằm ở loại đầu tiên khớp, "
               "loại khác ghi ở cột 'Lý do khác'. Mỗi thư mục sách có _thong_tin.txt", ""])
    for c in ws[1] + ws[ws.max_row]:
        c.font = Font(bold=True)
    for col, w in zip("ABCDEF", (24, 22, 9, 9, 90, 8)):
        ws.column_dimensions[col].width = w
    wrap = Alignment(wrap_text=True, vertical="top")
    for ma, ten, *_ in LOAI_KIEM_TRA:
        ds = [x for x in dong if x["ma"] == ma]
        if not ds:
            continue
        s = wb.create_sheet(ma[:31])
        s.append([h for _, h, _ in COT_KIEM_TRA])
        for c in s[1]:
            c.font, c.alignment = Font(bold=True), Alignment(horizontal="center", wrap_text=True)
        for x in ds:
            s.append([x.get(k, "") for k, _, _ in COT_KIEM_TRA])
            for c in s[s.max_row]:
                c.alignment = wrap
        for i, (_, _, rong) in enumerate(COT_KIEM_TRA):
            s.column_dimensions[get_column_letter(i + 1)].width = rong
        s.freeze_panes = "A2"
        s.auto_filter.ref = s.dimensions
    wb.save(xlsx)
    return xlsx


# Chạy cả bước: đọc Excel tab 4 + đối chiếu -> phân loại -> chép vào <thư mục>/CAN_KIEM_TRA/ -> Excel danh sách
def chay_can_kiem_tra(ket_qua_xlsx, doi_chieu_xlsx, thu_muc, chon_loai=None, ca_khac_nam=False, thu=False,
                      log=print, bao=lambda *a: None):
    chon_loai = list(chon_loai or [m for m, *_, md in LOAI_KIEM_TRA if md])
    goc_ra = thu_muc if os.path.basename(os.path.normpath(thu_muc)) == THU_MUC_KIEM_TRA else os.path.join(thu_muc, THU_MUC_KIEM_TRA)
    kq = doc_ket_qua_copy(ket_qua_xlsx)
    dc = sxb_tool.doc_doi_chieu(doi_chieu_xlsx) if doi_chieu_xlsx else []
    dong, kh = ke_hoach_kiem_tra(kq, dc, goc_ra, chon_loai, ca_khac_nam)
    log(f"{len(kq)} sách trong {os.path.basename(ket_qua_xlsx)} → {len(dong)} sách cần kiểm tra, {len(kh)} file")
    for ma, ten, *_ in LOAI_KIEM_TRA:
        if ma in chon_loai:
            ds = [x for x in dong if x["ma"] == ma]
            log(f"   {ma:24s} {len(ds):5d} sách  {sum(x['so_file'] for x in ds):6d} file   ({ten})")
    con = []
    for x, s, p in kh:
        if os.path.isfile(sxb_tool.duong_dan_dai(s)):
            con.append((x, s, p))
        else:
            x["ghi_chu"] = "; ".join(v for v in (x["ghi_chu"], f"không còn file {os.path.basename(s)}") if v)
    if not thu:
        os.makedirs(goc_ra, exist_ok=True)
    tem = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    xlsx = os.path.join(goc_ra if not thu else (os.path.dirname(ket_qua_xlsx) or "."), f"can_kiem_tra_{tem}.xlsx")
    chep = co_san = 0
    loi = []
    if thu:
        log("\n".join(f"{s}\n   → {p}" for _, s, p in con[:300]) + ("\n..." if len(con) > 300 else "") + "\n(chạy thử - chưa chép)")
        for x in dong:
            x["trang_thai"] = x["trang_thai"] or "CHẠY THỬ"
    else:
        for x in dong:
            _ghi_thong_tin(x)
        chep, co_san, loi = sxb_tool.luu_file(con, False, lambda i, xb, tb: bao("Chép file", i, len(con)))
        for x in dong:
            x["trang_thai"] = ("CHƯA CHÉP" if sxb_tool.DUNG.is_set() and not x["dich"] and x["so_file"] else
                               "LỖI CHÉP" if "LỖI lưu" in x["ghi_chu"] else x["trang_thai"] or ("ĐÃ CHÉP" if x["dich"] else "KHÔNG CÓ FILE"))
    ghi_excel_kiem_tra(dong, xlsx, chon_loai)
    log(f"Xong: {len(dong)} sách, " + ("chạy thử" if thu else f"chép {chep} file, {co_san} đã có sẵn, {len(loi)} lỗi")
        + f"\nThư mục: {goc_ra}\nDanh sách: {xlsx}")
    return xlsx, {"sach": len(dong), "file": len(con), "chep": chep, "co_san": co_san, "loi": loi}


# Tab 4. OCR & copy: file doi_chieu -> OCR tên sách + thời gian -> chép vào Năm/thang MM/<tên mới> -> Excel
def tab_ocr_copy(root, nb, ch, hang, vi_tri=None):
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    goc = thu_muc_ketqua()
    t = ttk.Frame(nb, padding=10)
    if vi_tri is not None and vi_tri < len(nb.tabs()):
        nb.insert(vi_tri, t, text="4. OCR & copy")
    else:
        nb.add(t, text="4. OCR & copy")
    ttk.Label(t, text="Sách KHỚP trong file đối chiếu: OCR tên sách và tháng/năm lưu chiểu, chép vào "
                      "<thư mục đích>/<Năm>/thang <MM>/<tên mới>/ theo thời gian OCR, đổi tên theo quy tắc, xuất Excel. "
                      "Ngày chỉ theo OCR (dòng In xong và nộp lưu chiểu, tối đa 5 trang cuối); năm trước 2006 / không ra ngày thì không chép vào thư mục năm mà copy vào <đích>/CAN_KIEM_TRA/TRUOC_2006 / KHONG_NGAY. OCR không ra tên thì dùng tên Excel.",
              foreground="#555", wraplength=760).grid(row=0, column=0, columnspan=4, sticky="w")
    noi = os.path.dirname(goc)
    tim_bbt = lambda dc: next((p for d in (os.path.dirname(dc), noi, os.path.dirname(noi))
                               for p in [os.path.join(d, sxb_tool.XL_MA_BBT)] if os.path.isfile(p)), "")
    dc_md = sxb_tool.moi_nhat(os.path.join(goc, "2_doi_chieu", "doi_chieu_*.xlsx"))
    v_dc, v_bbt = tk.StringVar(value=dc_md), tk.StringVar(value=tim_bbt(dc_md) if dc_md else "")
    v_dich = tk.StringVar(value=ch.get("dich_copy") or os.path.join(goc, "sach_da_luu"))
    v_ra = tk.StringVar(value=os.path.join(goc, LOAI["copy"][3]))
    chon_xl = lambda: filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
    hang(t, 1, "File doi_chieu_*.xlsx", v_dc, ("...", chon_xl))
    hang(t, 2, "File BBT_distinct.xlsx", v_bbt, ("...", chon_xl))
    hang(t, 3, "Thư mục đích", v_dich, ("...", filedialog.askdirectory))
    hang(t, 4, "Thư mục kết quả", v_ra, ("...", filedialog.askdirectory))
    o = ttk.Frame(t); o.grid(row=5, column=0, columnspan=4, sticky="w", pady=4)
    v_ten, v_move = tk.StringVar(value=ch.get("theo_ten", "excel")), tk.BooleanVar()
    v_khac = tk.BooleanVar(value=ch.get("bo_khac", True))
    v_lech, v_scan = tk.BooleanVar(value=ch.get("lay_lech", True)), tk.BooleanVar(value=ch.get("lay_scan", True))
    v_sl, v_to = tk.StringVar(value=str(ch.get("so_luong", 2))), tk.StringVar(value=str(ch.get("timeout", TIMEOUT_MAC_DINH)))
    ttk.Label(o, text="Tên sách theo").pack(side="left")
    ttk.Radiobutton(o, text="OCR", value="ocr", variable=v_ten).pack(side="left")
    ttk.Radiobutton(o, text="Excel", value="excel", variable=v_ten).pack(side="left")
    ttk.Label(o, text="Song song").pack(side="left", padx=(12, 2))
    ttk.Spinbox(o, from_=1, to=8, width=3, textvariable=v_sl).pack(side="left")
    ttk.Label(o, text="Chờ tối đa (s)").pack(side="left", padx=(12, 2))
    ttk.Entry(o, width=6, textvariable=v_to).pack(side="left")
    ttk.Checkbutton(o, text="Di chuyển file (mặc định: copy)", variable=v_move).pack(side="left", padx=12)
    chk = ttk.Frame(t); chk.grid(row=6, column=0, columnspan=4, sticky="w")
    ttk.Checkbutton(chk, text="Chỉ chép sách có tên OCR KHỚP tên Excel (không khớp thì OCR thêm bìa; vẫn không khớp: CẦN XEM)",
                    variable=v_khac).pack(anchor="w")
    ttk.Checkbutton(chk, text=f"Xử lý sách LỆCH NĂM (sheet CAN_XU_LY): thử tối đa {TOI_DA_UNG_VIEN} cặp file đến khi tên KHỚP",
                    variable=v_lech).pack(anchor="w")
    ttk.Checkbutton(chk, text="Xử lý PDF scan (sheet CAN_OCR): OCR, ghép vào sách theo tên; không khớp sách nào: "
                              "copy vào CAN_KIEM_TRA/KHONG_KHOP_SACH", variable=v_scan).pack(anchor="w")
    nut = ttk.Frame(t); nut.grid(row=7, column=1, sticky="w", pady=6)
    b_thu = ttk.Button(nut, text="Chạy thử"); b_thu.pack(side="left")
    b_chay = ttk.Button(nut, text="Thực hiện"); b_chay.pack(side="left", padx=6)
    b_dung = ttk.Button(nut, text="Dừng", state="disabled"); b_dung.pack(side="left")
    b_kn = ttk.Button(nut, text="Kiểm tra kết nối"); b_kn.pack(side="left", padx=6)
    b_ht = ttk.Button(nut, text="Hoàn tác..."); b_ht.pack(side="left")
    pb = ttk.Progressbar(t, maximum=1000); pb.grid(row=8, column=0, columnspan=4, sticky="we")
    lb = ttk.Label(t, text=""); lb.grid(row=9, column=0, columnspan=4, sticky="w")
    txt = tk.Text(t, height=12, wrap="none"); txt.grid(row=10, column=0, columnspan=4, sticky="nsew", pady=6)
    t.columnconfigure(1, weight=1); t.rowconfigure(10, weight=1)
    log = lambda s: root.after(0, lambda: (txt.insert("end", s + "\n"), txt.see("end")))
    t0 = [0.0]

    # Tiến độ: giờ tính riêng từng bước (OCR, OCR lệch năm, OCR PDF scan, OCR bìa, chép...) để ước tính còn lại đúng
    bat_dau_buoc = {}

    def bao(buoc, xong, tong):
        bd = bat_dau_buoc.setdefault(buoc, time.time())
        da, tong_da = time.time() - bd, time.time() - t0[0]
        con = f" - còn khoảng {sxb_tool._gio(da / xong * (tong - xong))}" if 0 < xong < tong else ""
        root.after(0, lambda: (pb.config(value=1000 * xong / max(tong, 1)),
                               lb.config(text=f"{buoc}: {xong}/{tong} - bước này {sxb_tool._gio(da)}{con}"
                                              f" (cả lần chạy {sxb_tool._gio(tong_da)})")))

    def ket_noi():
        def run():
            try:
                log(f"Kết nối máy chủ OK: {kiem_tra_ket_noi(API_MAC_DINH)}")
            except Exception as e:
                log(f"Không kết nối được máy chủ {API_MAC_DINH}: {type(e).__name__}")
        threading.Thread(target=run, daemon=True).start()

    def bat_dau(thu):
        dc, dich, ra = v_dc.get().strip(), v_dich.get().strip(), v_ra.get().strip()
        if not os.path.isfile(dc):
            messagebox.showerror("Lỗi", "Chưa chọn file doi_chieu_*.xlsx"); return
        if not os.path.isdir(dich):
            messagebox.showerror("Lỗi", "Chưa chọn thư mục đích"); return
        try:
            so_luong, timeout = int(v_sl.get()), int(v_to.get())
        except ValueError:
            messagebox.showerror("Lỗi", "Song song / Chờ tối đa phải là số"); return
        bbt = v_bbt.get().strip() or tim_bbt(dc)
        ch.update(so_luong=so_luong, timeout=timeout, dich_copy=dich, theo_ten=v_ten.get(), bo_khac=v_khac.get(),
                  lay_lech=v_lech.get(), lay_scan=v_scan.get())
        ghi_cau_hinh(ch)
        sxb_tool.DUNG.clear(); t0[0] = time.time(); bat_dau_buoc.clear()
        for b in (b_thu, b_chay):
            b.config(state="disabled")
        b_dung.config(state="normal")
        os.makedirs(ra, exist_ok=True)

        def run():
            try:
                chay_ocr_copy(dc, dich, ra, API_MAC_DINH, so_luong, timeout, v_ten.get(), v_move.get(), thu,
                              cache=os.path.join(ra, "ocr_cache.jsonl"), bbt=bbt, log=log, bao=bao, bo_khac=v_khac.get(),
                              lay_lech=v_lech.get(), lay_scan=v_scan.get())
            except Exception as e:
                log(f"LỖI: {type(e).__name__}: {e}")
            finally:
                root.after(0, lambda: ([b.config(state="normal") for b in (b_thu, b_chay)], b_dung.config(state="disabled")))
        threading.Thread(target=run, daemon=True).start()

    def dung_lai():
        sxb_tool.DUNG.set(); log("Đang dừng: chờ các cuốn đang OCR / file đang chép xong...")

    # Hoàn tác một lần chạy: chọn nhật ký chép (hoặc Excel kết quả nếu lần đó chưa có nhật ký), xem trước rồi mới làm
    def hoan_tac_ui():
        from collections import Counter
        ra = v_ra.get().strip()
        p = filedialog.askopenfilename(initialdir=ra if os.path.isdir(ra) else None, title="Chọn lần chạy cần hoàn tác",
                                       filetypes=[("Nhật ký chép / Excel kết quả", "nhat_ky_chep_*.csv ocr_copy_*.xlsx"),
                                                  ("Tất cả", "*.*")])
        if not p:
            return
        try:
            kh = ke_hoach_hoan_tac(p)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không đọc được {os.path.basename(p)}: {e}"); return
        dem = Counter(v for v, *_ in kh)
        ly_do = Counter(x[3] for x in kh if x[0] == "bỏ qua")
        log(f"Hoàn tác {os.path.basename(p)}: {dem['xoá']} bản chép sẽ xoá, {dem['trả về']} file trả về chỗ cũ, "
            f"{dem['bỏ qua']} bỏ qua" + "".join(f"\n  bỏ qua {n}: {k}" for k, n in ly_do.most_common()))
        if not dem["xoá"] and not dem["trả về"]:
            messagebox.showinfo("Hoàn tác", "Không có file nào cần hoàn tác"); return
        canh_bao = ("\n\nĐây là file Excel (lần chạy chưa có nhật ký): chỉ đúng khi lần đó COPY file. "
                    "Nếu lần đó DI CHUYỂN file thì chọn Không." if p.lower().endswith(".xlsx") else "")
        if not messagebox.askyesno("Hoàn tác", f"Xoá {dem['xoá']} bản chép ở thư mục đích, trả {dem['trả về']} file "
                                               f"đã di chuyển về chỗ cũ, xoá thư mục rỗng?{canh_bao}\n\nTiếp tục?"):
            return
        b_ht.config(state="disabled")

        def run():
            try:
                hoan_tac(kh, log)
            except Exception as e:
                log(f"LỖI: {type(e).__name__}: {e}")
            finally:
                root.after(0, lambda: b_ht.config(state="normal"))
        threading.Thread(target=run, daemon=True).start()

    b_thu.config(command=lambda: bat_dau(True)); b_chay.config(command=lambda: bat_dau(False))
    b_dung.config(command=dung_lai); b_kn.config(command=ket_noi); b_ht.config(command=hoan_tac_ui)


# Tab 5. Cần kiểm tra: chọn Excel tab 4 + đối chiếu -> chép file các sách cần xem vào CAN_KIEM_TRA/<loại>/
def tab_can_kiem_tra(root, nb, ch, hang, vi_tri=None):
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    goc = thu_muc_ketqua()
    t = ttk.Frame(nb, padding=10)
    if vi_tri is not None and vi_tri < len(nb.tabs()):
        nb.insert(vi_tri, t, text="5. Cần kiểm tra")
    else:
        nb.add(t, text="5. Cần kiểm tra")
    ttk.Label(t, text="Từ Excel kết quả tab 4: chép toàn bộ file liên quan (Word, PDF, ảnh, bìa...) của các sách cần xem "
                      f"vào <thư mục>/{THU_MUC_KIEM_TRA}/<loại>/<Năm>/<Năm Excel>_<Dòng Excel>_<tên sách>/, giữ cấu trúc thư mục con, mỗi sách có "
                      "_thong_tin.txt, kèm Excel danh sách. Chỉ copy, không đụng file gốc.",
              foreground="#555", wraplength=760).grid(row=0, column=0, columnspan=4, sticky="w")
    v_kq = tk.StringVar(value=sxb_tool.moi_nhat(os.path.join(goc, LOAI["copy"][3], "ocr_copy_*.xlsx")))
    v_dc = tk.StringVar(value=sxb_tool.moi_nhat(os.path.join(goc, "2_doi_chieu", "doi_chieu_*.xlsx")))
    v_tm = tk.StringVar(value=ch.get("thu_muc_kiem_tra") or goc)
    chon_xl = lambda: filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
    hang(t, 1, "Kết quả ocr_copy_*.xlsx", v_kq, ("...", chon_xl))
    hang(t, 2, "File doi_chieu_*.xlsx", v_dc, ("...", chon_xl))
    hang(t, 3, f"Nơi tạo {THU_MUC_KIEM_TRA}", v_tm, ("...", filedialog.askdirectory))
    lf = ttk.LabelFrame(t, text="Loại cần chép (một sách chỉ vào loại đầu tiên khớp)", padding=6)
    lf.grid(row=4, column=0, columnspan=4, sticky="we", pady=4)
    da_chon = ch.get("loai_kiem_tra")
    if da_chon is not None and not set(da_chon) <= {m for m, *_ in LOAI_KIEM_TRA}:
        da_chon = None  # cấu hình lưu từ bản cũ (tên thư mục đã đổi): dùng mặc định
    v_loai = {}
    for i, (ma, ten, giai_thich, _, md) in enumerate(LOAI_KIEM_TRA):
        v_loai[ma] = tk.BooleanVar(value=(ma in da_chon) if da_chon is not None else md)
        ttk.Checkbutton(lf, text=f"{ma}: {giai_thich}", variable=v_loai[ma]).grid(row=i, column=0, sticky="w")
    v_khac = tk.BooleanVar(value=False)
    ttk.Checkbutton(t, text="Chép cả file khác năm trong 'Tất cả file tìm được' (mặc định: chỉ file cùng năm + Word / PDF đã chọn)",
                    variable=v_khac).grid(row=5, column=0, columnspan=4, sticky="w")
    nut = ttk.Frame(t); nut.grid(row=6, column=1, sticky="w", pady=6)
    b_thu = ttk.Button(nut, text="Chạy thử"); b_thu.pack(side="left")
    b_chay = ttk.Button(nut, text="Thực hiện"); b_chay.pack(side="left", padx=6)
    b_dung = ttk.Button(nut, text="Dừng", state="disabled"); b_dung.pack(side="left")
    pb = ttk.Progressbar(t, maximum=1000); pb.grid(row=7, column=0, columnspan=4, sticky="we")
    lb = ttk.Label(t, text=""); lb.grid(row=8, column=0, columnspan=4, sticky="w")
    txt = tk.Text(t, height=10, wrap="none"); txt.grid(row=9, column=0, columnspan=4, sticky="nsew", pady=6)
    t.columnconfigure(1, weight=1); t.rowconfigure(9, weight=1)
    log = lambda s: root.after(0, lambda: (txt.insert("end", s + "\n"), txt.see("end")))
    t0 = [0.0]

    def bao(buoc, xong, tong):
        da = time.time() - t0[0]
        root.after(0, lambda: (pb.config(value=1000 * xong / max(tong, 1)),
                               lb.config(text=f"{buoc}: {xong}/{tong} - đã chạy {sxb_tool._gio(da)}")))

    def bat_dau(thu):
        kq, dc, tm = v_kq.get().strip(), v_dc.get().strip(), v_tm.get().strip()
        if not os.path.isfile(kq):
            messagebox.showerror("Lỗi", "Chưa chọn file kết quả ocr_copy_*.xlsx của tab 4"); return
        if not os.path.isfile(dc):
            messagebox.showerror("Lỗi", "Chưa chọn file doi_chieu_*.xlsx (cần để biết các file của từng sách)"); return
        if not tm:
            messagebox.showerror("Lỗi", f"Chưa chọn nơi tạo {THU_MUC_KIEM_TRA}"); return
        loai = [m for m, v in v_loai.items() if v.get()]
        if not loai:
            messagebox.showerror("Lỗi", "Chưa chọn loại nào"); return
        ch.update(thu_muc_kiem_tra=tm, loai_kiem_tra=loai)
        ghi_cau_hinh(ch)
        sxb_tool.DUNG.clear(); t0[0] = time.time()
        for b in (b_thu, b_chay):
            b.config(state="disabled")
        b_dung.config(state="normal")

        def run():
            try:
                chay_can_kiem_tra(kq, dc, tm, loai, v_khac.get(), thu, log=log, bao=bao)
            except Exception as e:
                log(f"LỖI: {type(e).__name__}: {e}")
            finally:
                root.after(0, lambda: ([b.config(state="normal") for b in (b_thu, b_chay)], b_dung.config(state="disabled")))
        threading.Thread(target=run, daemon=True).start()

    def dung_lai():
        sxb_tool.DUNG.set(); log("Đang dừng: chờ file đang chép xong...")

    b_thu.config(command=lambda: bat_dau(True)); b_chay.config(command=lambda: bat_dau(False)); b_dung.config(command=dung_lai)


# Thêm các tab OCR
def them_tab_ocr(root, nb):
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    ch = doc_cau_hinh()
    root.title("Tool SXB - quét kho, đối chiếu, kiểm tra nội dung, OCR & copy, cần kiểm tra, metadata, thông tin trang cuối")
    root.geometry("940x660")

    def hang(parent, r, nhan, var, *nut):
        ttk.Label(parent, text=nhan, width=20).grid(row=r, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=var, width=62).grid(row=r, column=1, sticky="we")
        for j, (chu, ham) in enumerate(nut):
            ttk.Button(parent, text=chu, width=8, command=lambda h=ham: var.set(h() or var.get())).grid(row=r, column=2 + j)

    def tao_tab(loai, tieu_de, mo_ta):
        t = ttk.Frame(nb, padding=10); nb.add(t, text=tieu_de)
        ttk.Label(t, text=mo_ta, foreground="#555", wraplength=720).grid(row=0, column=0, columnspan=4, sticky="w")
        ra_md = os.path.join(thu_muc_ketqua(), LOAI[loai][3])
        v_vao, v_ra, v_xl = tk.StringVar(value=ch.get(f"vao_{loai}", "")), tk.StringVar(value=ra_md), tk.StringVar()
        v_sl, v_to = tk.StringVar(value=str(ch.get("so_luong", 2))), tk.StringVar(value=str(ch.get("timeout", TIMEOUT_MAC_DINH)))
        hang(t, 1, "Excel đường dẫn", v_xl,
             ("...", lambda: filedialog.askopenfilename(initialdir=thu_muc_ketqua() if os.path.isdir(thu_muc_ketqua()) else None,
                                                        filetypes=[("Excel", "*.xlsx")])),
             ("Mới nhất", lambda: excel_moi_nhat() or (log("Chưa có file kết quả nào của tab 1-4 trong ketqua/") and "")),
             ("Bỏ", lambda: v_xl.set("")))
        ttk.Label(t, text="Quét theo đường dẫn có trong Excel", foreground="#555", wraplength=640)\
            .grid(row=2, column=1, columnspan=3, sticky="w")
        hang(t, 3, "Thư mục / file sách", v_vao, ("Thư mục", filedialog.askdirectory),
             ("File", lambda: filedialog.askopenfilename(filetypes=[("PDF / Word", "*.pdf *.docx *.doc")])))
        hang(t, 4, "Thư mục kết quả", v_ra, ("...", filedialog.askdirectory))
        o = ttk.Frame(t); o.grid(row=5, column=0, columnspan=4, sticky="w", pady=4)
        ttk.Label(o, text="Song song").pack(side="left", padx=(0, 2))
        ttk.Spinbox(o, from_=1, to=8, width=3, textvariable=v_sl).pack(side="left")
        ttk.Label(o, text="Chờ tối đa (s)").pack(side="left", padx=(12, 2))
        ttk.Entry(o, width=6, textvariable=v_to).pack(side="left")
        nut = ttk.Frame(t); nut.grid(row=6, column=1, sticky="w", pady=6)
        b_chay = ttk.Button(nut, text="Bắt đầu quét"); b_chay.pack(side="left")
        b_dung = ttk.Button(nut, text="Dừng", state="disabled"); b_dung.pack(side="left", padx=6)
        b_kn = ttk.Button(nut, text="Kiểm tra kết nối"); b_kn.pack(side="left")
        txt = tk.Text(t, height=12, wrap="none"); txt.grid(row=7, column=0, columnspan=4, sticky="nsew", pady=8)
        t.columnconfigure(1, weight=1); t.rowconfigure(7, weight=1)
        log = lambda s: root.after(0, lambda: (txt.insert("end", s + "\n"), txt.see("end")))
        dung = threading.Event()

        def ket_noi():
            def run():
                try:
                    log(f"Kết nối máy chủ OK: {kiem_tra_ket_noi(API_MAC_DINH)}")
                except Exception as e:
                    log(f"Không kết nối được máy chủ: {type(e).__name__}")
            threading.Thread(target=run, daemon=True).start()

        def bat_dau():
            vao, xl = v_vao.get().strip(), v_xl.get().strip()
            if xl and not os.path.isfile(xl):
                messagebox.showerror("Lỗi", f"Không thấy file Excel đường dẫn: {xl}"); return
            if not xl and not os.path.exists(vao):
                messagebox.showerror("Lỗi", "Chưa chọn Excel đường dẫn hoặc thư mục / file sách"); return
            try:
                so_luong, timeout = int(v_sl.get()), int(v_to.get())
            except ValueError:
                messagebox.showerror("Lỗi", "Song song / Chờ tối đa phải là số"); return
            bo_qua = []
            if xl:
                try:
                    sach = liet_ke_tu_excel(xl, bo_qua)
                except Exception as e:
                    messagebox.showerror("Lỗi", f"Không đọc được Excel đường dẫn: {e}"); return
                log(f"Excel đường dẫn {os.path.basename(xl)}: {len(sach)} cuốn")
            else:
                sach = liet_ke_sach(vao, bo_qua=bo_qua)
            if bo_qua:
                log(f"Bỏ qua {len(bo_qua)} file:")
                for p, ly_do in bo_qua[:200]:
                    log(f"   - {os.path.basename(p)}  ({ly_do})")
                if len(bo_qua) > 200:
                    log(f"   ... và {len(bo_qua) - 200} file khác")
            if not sach:
                messagebox.showerror("Lỗi", "Không có cuốn nào (PDF / Word) " + ("trong Excel đường dẫn" if xl else "trong thư mục")); return
            ra = v_ra.get().strip() or ra_md
            try:
                os.makedirs(ra, exist_ok=True)
            except OSError as e:
                messagebox.showerror("Lỗi", f"Không tạo được thư mục kết quả: {e}"); return
            xlsx = os.path.join(ra, f"{LOAI[loai][2]}_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx")
            ch.update({"so_luong": so_luong, "timeout": timeout}, **({} if xl else {f"vao_{loai}": vao}))
            ghi_cau_hinh(ch)
            dung.clear()
            b_chay.config(state="disabled"); b_dung.config(state="normal")
            log(f"{len(sach)} cuốn, upload lên máy chủ (song song {so_luong}) → {xlsx}")

            def run():
                t0 = time.time()
                try:
                    dong = quet_api(API_MAC_DINH, loai, sach, xlsx, so_luong, timeout, log, dung)
                    ok = sum(d["trang_thai"] == "OK" for d in dong)
                    log(f"Xong {len(dong)}/{len(sach)} cuốn ({ok} OK) sau {time.time() - t0:.0f}s → {xlsx}")
                except Exception as e:
                    log(f"Lỗi: {e}")
                finally:
                    root.after(0, lambda: (b_chay.config(state="normal"), b_dung.config(state="disabled")))
            threading.Thread(target=run, daemon=True).start()

        def dung_lai():
            dung.set(); log("Đang dừng: chờ các cuốn đang gửi xong, không gửi cuốn mới...")

        b_chay.config(command=bat_dau); b_dung.config(command=dung_lai); b_kn.config(command=ket_noi)

    # Tab 4 "Đổi tên & lưu" của tool SXB thay bằng "OCR & copy" (gộp đổi tên & lưu + quét tên sách)
    vi_tri = None
    for t in nb.tabs():
        if nb.tab(t, "text").endswith("Đổi tên & lưu"):
            vi_tri = nb.index(t)
            nb.forget(t)
            break
    tab_ocr_copy(root, nb, ch, hang, vi_tri)
    tab_can_kiem_tra(root, nb, ch, hang, None if vi_tri is None else vi_tri + 1)
    tao_tab("metadata", "6. Quét metadata",
            "Trích xuất tháng, năm in xong và nộp lưu chiểu tại trang cuối. Kết quả được ghi vào Excel.")
    tao_tab("info", "7. Quét thông tin trang cuối",
            "Trích xuất các trường thông tin xuất bản tại trang cuối. Kết quả được ghi vào Excel.")


# Mở giao diện: tab 1-3 tool SXB + OCR & copy, metadata, thông tin trang cuối
def gui():
    sxb_tool.gui(4, them_tab=them_tab_ocr)


# Tự kiểm tra
def _test():
    import tempfile, shutil, openpyxl
    from email.parser import BytesParser
    from email.policy import default
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
    from urllib.parse import parse_qs
    sxb_tool._test()

    tmp = tempfile.mkdtemp()
    a = os.path.join(tmp, "a"); b = os.path.join(tmp, "b"); c = os.path.join(tmp, "a", "con")
    for d in (a, b, c):
        os.makedirs(d)
    for p in (os.path.join(a, "Sach 1.pdf"), os.path.join(a, "Sach 1.docx"), os.path.join(a, "Sach 2.pdf"),
              os.path.join(a, "~$Sach 1.docx"), os.path.join(b, "chi word.docx"), os.path.join(c, "Con.pdf"),
              os.path.join(a, "Bìa Sach 1.pdf"), os.path.join(a, "A bia.docx"), os.path.join(a, "anh.jpg"),
              os.path.join(a, "Tư tưởng Hồ Chí Minh - 01.pdf"), os.path.join(b, "BIA_truoc.pdf"),
              os.path.join(b, "chi word.doc"), os.path.join(c, "Con.doc")):
        open(p, "wb").write(b"%PDF-1 test")
    s1, s2, w1 = os.path.join(a, "Sach 1.pdf"), os.path.join(a, "Sach 2.pdf"), os.path.join(a, "Sach 1.docx")
    assert tim_word(s1) == w1 and tim_word(s2) == w1
    assert tim_word(os.path.join(c, "Con.pdf")) == os.path.join(c, "Con.doc")
    assert ly_do_bo_qua("Bìa sách.pdf") and ly_do_bo_qua("x - BIA.docx") and ly_do_bo_qua("bia_sau.pdf")
    assert ly_do_bo_qua("a.JPG") and ly_do_bo_qua("Tư tưởng Hồ Chí Minh - 01.pdf") and ly_do_bo_qua("Sach -3 .pdf")
    assert not ly_do_bo_qua("Biên bản.pdf") and not ly_do_bo_qua("Lịch sử - Tập 1.pdf") and not ly_do_bo_qua("Sach - 01.docx")
    assert not ly_do_bo_qua("Sach 2020.pdf") and not ly_do_bo_qua("Sach-2.pdf.docx")
    bo = []
    ds = liet_ke_sach(tmp, bo_qua=bo)
    assert sorted(os.path.basename(p) for p, _ in bo) == ["A bia.docx", "BIA_truoc.pdf", "Bìa Sach 1.pdf",
                                                          "Tư tưởng Hồ Chí Minh - 01.pdf", "anh.jpg"], bo
    assert ds == [(s1, w1), (s2, w1), (os.path.join(c, "Con.pdf"), os.path.join(c, "Con.doc")),
                  ("", os.path.join(b, "chi word.docx"))], ds
    assert liet_ke_sach(tmp, can_pdf=True) == [(s1, ""), (s2, ""), (os.path.join(c, "Con.pdf"), "")]
    assert len(liet_ke_sach(a, de_quy=False)) == 2 and liet_ke_sach(s1) == [(s1, w1)]

    def tao_xl(ten, hdr, rows):
        wb = openpyxl.Workbook(); ws = wb.active; ws.append(hdr)
        for r in rows:
            ws.append(r)
        p = os.path.join(tmp, ten); wb.save(p)
        return p
    tat_ca = [os.path.join(g, f) for g, _, fs in os.walk(tmp) for f in fs] + [os.path.join(tmp, "x", "Mat.pdf")]
    bo = []
    kho = tao_xl("kho_1.xlsx", sxb_tool.COT_KHO, [[i, os.path.basename(x), "", "", "", "", x] for i, x in enumerate(tat_ca, 1)])
    assert liet_ke_tu_excel(kho, bo) == ds, liet_ke_tu_excel(kho)
    assert (os.path.join(tmp, "x", "Mat.pdf"), "không còn file") in bo and any(p.endswith("anh.jpg") for p, _ in bo)
    c_pdf = os.path.join(c, "Con.pdf")
    dc = tao_xl("doi_chieu_1.xlsx", ["Năm", "Tên sách", "File Word chọn", "File PDF chọn", "Đường dẫn đích"],
                [[2021, "Sách 1", w1, s1, None], [2021, "Con", None, c_pdf, None], [2021, "Trống", None, None, None]])
    assert liet_ke_tu_excel(dc) == [(s1, w1), (c_pdf, "")]
    dich = tao_xl("doi_chieu_2.xlsx", ["Tên sách", "File PDF chọn", "Đường dẫn đích"],
                  [["Sách 1", "C:/cu/Sach 1.pdf", "\n".join((s1, w1, os.path.join(a, "Bia", "Sach 1_B.pdf")))]])
    assert liet_ke_tu_excel(dich) == [(s1, w1)]

    nhan = []

    class H(BaseHTTPRequestHandler):
        def log_message(self, *a): pass

        def _tra(self, code, obj):
            body = json.dumps(obj).encode()
            self.send_response(code); self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body))); self.end_headers(); self.wfile.write(body)

        def do_GET(self):
            if self.path == "/health":
                return self._tra(200, {"status": "healthy"})
            if self.path.startswith("/api/request-status/"):
                return self._tra(200, {"request_id": "r9", "status": "SUCCESS",
                                       "result": {"job_id": "j9", "status": "completed", "error": "",
                                                  "result": {"ten_sach": "SÁCH CHỜ", "text_source": "ocr"}}})
            self._tra(404, {})

        def do_POST(self):
            body = self.rfile.read(int(self.headers["Content-Length"]))
            if self.headers["Content-Type"].startswith("multipart/"):
                msg = BytesParser(policy=default).parsebytes(
                    f"Content-Type: {self.headers['Content-Type']}\r\n\r\n".encode() + body)
                f = {p.get_param("name", header="content-disposition"):
                     p.get_filename() or p.get_content().strip() for p in msg.iter_parts()}
            else:
                f = {k: v[0] for k, v in parse_qs(body.decode()).items()}
            nhan.append((self.path, f))
            ten = os.path.basename(f.get("pdf") or f.get("pdf_path") or f.get("word") or "")
            if f.get("bia") and not ten:  # chỉ gửi bìa
                return self._tra(200, {"job_id": "j", "status": "completed", "error": "", "result": {
                    "ten_sach": "TÊN TỪ BÌA", "text_source": "bia"}})
            if "Trong" in ten and self.path.endswith("/title"):  # Word / PDF không ra tên
                return self._tra(200, {"job_id": "j", "status": "completed", "error": "", "result": {
                    "ten_sach": "", "text_source": "ocr"}})
            if "Sach 2" in ten:
                return self._tra(202, {"error_code": 202, "error_message": "cho", "request_id": "r9"})
            if "Con" in ten:
                return self._tra(500, {"job_id": "j3", "status": "failed", "error": "Lỗi OCR: x", "result": {}})
            if ten.startswith("Ten_") and self.path.endswith("/title"):  # file 'Ten_<tên>.pdf': OCR ra đúng <tên>
                return self._tra(200, {"job_id": "j", "status": "completed", "error": "", "result": {
                    "ten_sach": os.path.splitext(ten)[0][4:].replace("_", " "), "text_source": "ocr"}})
            if "Cu" in ten and self.path.endswith("/metadata"):  # sách OCR ra năm trước 2006
                return self._tra(200, {"job_id": "j", "status": "completed", "error": "", "result": {
                    "metadata": "In xong và nộp lưu chiểu tháng 5 năm 1951", "luu_chieu_thang": 5, "luu_chieu_nam": 1951}})
            if self.path.endswith("/info"):
                return self._tra(200, {"job_id": "j", "status": "completed", "error": "", "result": {
                    "thong_tin": {"nha_xuat_ban": "NXB Trẻ", "isbn": "978-604-1", "luu_chieu_thang": 3,
                                  "luu_chieu_nam": 2021, "fax_nxb": "024.1"},
                    "thong_tin_text": "NXB Trẻ\nISBN 978-604-1", "thong_tin_source": "ocr", "thong_tin_page": 120,
                    "page_count": 120}})
            if self.path.endswith("/metadata"):
                return self._tra(200, {"job_id": "j", "status": "completed", "error": "", "result": {
                    "metadata": "ISBN 978\nNXB Trẻ", "metadata_raw": "ISBN 978\nNXB Tre", "metadata_score": 0.95,
                    "metadata_box": [1, 2, 3, 4], "page_count": 120, "luu_chieu_thang": 3, "luu_chieu_nam": 2021}})
            self._tra(200, {"job_id": "j", "status": "completed", "error": "", "result": {
                "ten_sach": "LỊCH SỬ", "ten_sach_raw": "LICH SU", "text_source": "word" if f.get("word") else "ocr",
                "page_count": 120}})

    sv = ThreadingHTTPServer(("127.0.0.1", 0), H)
    threading.Thread(target=sv.serve_forever, daemon=True).start()
    api = f"http://127.0.0.1:{sv.server_address[1]}"
    try:
        assert kiem_tra_ket_noi(api)["status"] == "healthy"
        xl = os.path.join(tmp, "ten.xlsx")
        dong = quet_api(api, "title", ds, xl, so_luong=2, log=lambda s: None)
        assert [d["trang_thai"] for d in dong] == ["OK", "OK", "LỖI", "OK"], [d["trang_thai"] for d in dong]
        assert dong[0]["ten_sach"] == "LỊCH SỬ" and dong[0]["nguon"] == "Word" and dong[1]["ten_sach"] == "SÁCH CHỜ"
        assert dong[2]["loi"] == "Lỗi OCR: x" and dong[3]["nguon"] == "Word"
        up = dict((os.path.basename(f.get("pdf", "") or f.get("word", "")), f) for _, f in nhan)
        assert up["Sach 1.pdf"]["word"] == "Sach 1.docx" and up["Con.pdf"]["word"] == "Con.doc" and "pdf" not in up["chi word.docx"]
        ws = openpyxl.load_workbook(xl).active
        assert [c.value for c in ws[1]][:2] == ["STT", "Tên sách"] and ws.max_row == 5 and ws["B2"].value == "LỊCH SỬ"
        assert liet_ke_tu_excel(xl) == ds

        nhan.clear()
        xl2 = os.path.join(tmp, "meta.xlsx")
        dong = quet_api(api, "metadata", [(s1, "")], xl2, log=lambda s: None)
        assert dong[0]["metadata"] == "ISBN 978\nNXB Trẻ" and dong[0]["vung"] == "1,2,3,4" and dong[0]["diem"] == 0.95
        assert nhan[0][0] == "/api/sort/metadata" and nhan[0][1]["pdf"] == "Sach 1.pdf" and "word" not in nhan[0][1]
        assert dong[0]["thang"] == 3 and dong[0]["nam"] == 2021
        ws2 = openpyxl.load_workbook(xl2).active
        assert [c.value for c in ws2[1]][:7] == ["STT", "Tháng", "Năm", "Nguồn", "Trang", "Dòng In xong", "METADATA"]
        assert (ws2["B2"].value, ws2["C2"].value, ws2["G2"].value) == (3, 2021, "ISBN 978\nNXB Trẻ")

        nhan.clear()
        xl3 = os.path.join(tmp, "info.xlsx")
        dong = quet_api(api, "info", [(s1, w1)], xl3, log=lambda s: None)
        assert nhan[0][0] == "/api/sort/info" and nhan[0][1]["word"] == "Sach 1.docx"
        assert (dong[0]["nha_xuat_ban"], dong[0]["isbn"], dong[0]["so_truong"], dong[0]["nguon"]) == ("NXB Trẻ", "978-604-1", 3, "OCR")
        ws3 = openpyxl.load_workbook(xl3).active
        hdr = [c.value for c in ws3[1]]
        assert ws3.title == "Thông tin trang cuối" and len(hdr) == len(COT_THONG_TIN) and hdr[1] == "Nhà xuất bản"
        assert ws3.cell(2, hdr.index("Fax NXB") + 1).value == "024.1" and ws3.cell(2, hdr.index("File Word") + 1).value == w1
        assert dong_ket_qua("info", 1, "x.pdf", "", {"status": "completed", "result": {"thong_tin": {}}}, 1)["trang_thai"] == "KHÔNG RA THÔNG TIN"

        d = dong_ket_qua("title", 1, "x.pdf", "", None, 1.0, loi="ConnectionError: x")
        assert d["trang_thai"] == "LỖI" and d["ten_sach"] == ""
        assert dong_ket_qua("title", 1, "x.pdf", "", {"status": "completed", "result": {"ten_sach": ""}}, 1)["trang_thai"] == "KHÔNG RA TÊN"

        # OCR & copy: Sach 1 OCR đủ tên + 3/2021; Sach 2 chỉ ra tên (202 rồi chờ); Con lỗi -> dùng Excel
        ngay = "2024-01-01 08:00:00"
        hdr = [h for _, h in sxb_tool.COT_KQ]
        dong_kq = lambda **k: [k.get(key, "") for key, _ in sxb_tool.COT_KQ]
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "KHOP"; ws.append(hdr)
        ws.append(dong_kq(nam=2020, dong=5, ten="Sách một", bbt="Đảng", qdxb="01/2020", tinh_trang="KHỚP",
                          file_word=w1, file_pdf=s1, cac_file=sxb_tool._ds_file(
                              [(ngay, s1), (ngay, w1), (ngay, os.path.join(a, "anh.jpg"))])))
        ws.append(dong_kq(nam=2020, dong=6, ten="Sách một bản 2", qdxb="02/2020", tinh_trang="KHỚP",
                          file_word=w1, file_pdf=s1))
        ws.append(dong_kq(nam=2019, dong=7, ten="Con", qdxb="05/2019", tinh_trang="KHỚP", file_pdf=c_pdf))
        ws.append(dong_kq(nam=2018, dong=8, ten="Sách hai", tinh_trang="KHỚP", file_pdf=s2))
        ws.append(dong_kq(nam=2018, dong=9, ten="Mất file", tinh_trang="KHỚP", file_pdf=os.path.join(tmp, "x", "Mat.pdf")))
        ws2 = wb.create_sheet("CAN_XU_LY"); ws2.append(hdr)
        ws2.append(dong_kq(nam=2018, dong=10, ten="Lệch", tinh_trang="LỆCH NĂM - cần mở file kiểm tra", file_pdf=s2))
        dc_oc = os.path.join(tmp, "doi_chieu_oc.xlsx"); wb.save(dc_oc)
        bbt = tao_xl("BBT_distinct.xlsx", ["STT", "BBT", "", "", "", "Mã"], [[1, "Đảng", None, None, None, "D"]])
        dich, ra = os.path.join(tmp, "dich"), os.path.join(tmp, "ra")
        os.makedirs(dich)
        cache, im = os.path.join(ra, "ocr_cache.jsonl"), lambda s: None

        nhan.clear(); sxb_tool.DUNG.clear()
        xl_thu, tk_thu = chay_ocr_copy(dc_oc, dich, ra, api, 2, 30, "ocr", thu=True, cache=cache, bbt=bbt, log=im, bo_khac=False, lay_lech=False)
        assert tk_thu["ocr"] == 3 and len(nhan) == 6, nhan  # 3 cặp file khác nhau x (tên + thời gian), dòng trùng file chỉ OCR 1 lần
        assert os.listdir(dich) == [] and len(doc_cache_ocr(cache)) == 2, (os.listdir(dich), doc_cache_ocr(cache))  # chạy thử không chép; cặp lỗi không lưu cache
        rows = sxb_tool._doc_sheet(xl_thu, "OCR & copy", [(k, h) for k, h, _ in COT_OCR_COPY])
        # Excel sắp theo thời gian mới: 2018/thang 00 -> 2019/thang 05 -> 2021/thang 03
        assert [r["ten_excel"] for r in rows] == ["Mất file", "Sách hai", "Con", "Sách một", "Sách một bản 2"], rows
        # Sách hai (chỉ ra tên), Con (OCR lỗi) không ra ngày: không dùng ngày Excel, không chép vào thư mục năm
        assert [r["trang_thai"] for r in rows] == ["KHÔNG CÓ FILE"] + [KHONG_RA_NGAY] * 2 + ["CHẠY THỬ"] * 2, rows
        assert rows[3]["ten_moi"] == "lich_su_0321_R_D_M00" and rows[3]["nguon_ten"] == "OCR" and rows[3]["nguon_ngay"] == "OCR"

        nhan.clear()
        xl_oc, tk_oc = chay_ocr_copy(dc_oc, dich, ra, api, 2, 30, "ocr", cache=cache, bbt=bbt, log=im, bo_khac=False, lay_lech=False)
        assert len(nhan) == 2 and tk_oc["ocr"] == 1, nhan  # chạy lại: chỉ OCR lại cặp lỗi
        moi = os.path.join(dich, "2021", "thang 03", "lich_su_0321_R_D_M00")
        assert sorted(os.listdir(moi)) == ["Bia", "lich_su_0321_R_D_M00.docx", "lich_su_0321_R_D_M00.pdf"], os.listdir(moi)
        assert os.listdir(os.path.join(moi, "Bia")) == ["lich_su_0321_B_D_M00.jpg"]
        assert os.path.isdir(os.path.join(dich, "2021", "thang 03", "lich_su_0321_R_XX_M00"))  # dòng trùng file: BBT trống
        nam_dich = lambda: sorted(x for x in os.listdir(dich) if x != THU_MUC_KIEM_TRA)
        assert sorted(os.listdir(dich)) == ["2021", THU_MUC_KIEM_TRA]  # không ra ngày: không chép theo ngày Excel
        # Sách hai: Word / PDF không ra dòng In xong -> copy vào CAN_KIEM_TRA/KHONG_NGAY/<năm Excel>/; Con OCR lỗi: không copy
        kn = os.path.join(dich, THU_MUC_KIEM_TRA, MA_KHONG_NGAY)
        assert os.listdir(kn) == ["2018"] and sorted(os.listdir(os.path.join(kn, "2018", "2018_0008_sach_hai"))) == \
            ["Sach 2.pdf", "_thong_tin.txt"], os.listdir(kn)
        rows = sxb_tool._doc_sheet(xl_oc, "OCR & copy", [(k, h) for k, h, _ in COT_OCR_COPY])
        assert [r["trang_thai"] for r in rows] == ["KHÔNG CÓ FILE"] + [KHONG_RA_NGAY] * 2 + ["ĐÃ LƯU"] * 2, rows
        r_con = rows[2]
        assert (r_con["nguon_ten"], r_con["nguon_ngay"], r_con["ten_excel"], r_con["dich"]) == ("Excel", KHONG_RA_NGAY, "Con", "")
        assert "OCR lỗi" in r_con["ghi_chu"] and "không còn file Mat.pdf" in rows[0]["ghi_chu"]
        assert "KHONG_NGAY" in rows[1]["ghi_chu"] and "KHONG_NGAY" in rows[1]["dich"], rows[1]
        assert (rows[1]["ten_ocr"], rows[1]["nguon_ngay"], rows[1]["thang_lc"], rows[1]["nam_lc"]) == ("SÁCH CHỜ", KHONG_RA_NGAY, "", "")

        # Hoàn tác theo nhật ký (copy): xoá 5 bản chép + 1 bản trong KHONG_NGAY, xoá thư mục rỗng, file gốc còn nguyên
        nk = xl_oc.replace("ocr_copy_", "nhat_ky_chep_").replace(".xlsx", ".csv")
        kh_ht = ke_hoach_hoan_tac(nk)
        assert [v for v, *_ in kh_ht] == ["xoá"] * 6, kh_ht
        assert hoan_tac(kh_ht, im)[:2] == (6, 0) and nam_dich() == [], os.listdir(dich)
        assert all(os.path.isfile(p) for p in (s1, w1, s2, c_pdf, os.path.join(a, "anh.jpg")))
        # Hoàn tác theo Excel kết quả (lần chạy chưa có nhật ký, copy)
        xl_c, _ = chay_ocr_copy(dc_oc, dich, ra, api, 2, 30, "ocr", cache=cache, bbt=bbt, log=im, bo_khac=False, lay_lech=False)
        assert [v for v, *_ in ke_hoach_hoan_tac(xl_c)] == ["xoá"] * 6
        hoan_tac(ke_hoach_hoan_tac(xl_c), im)
        assert nam_dich() == []
        # Di chuyển rồi hoàn tác theo nhật ký: file về đúng chỗ cũ; Excel của lần di chuyển thì không xoá gì
        xl_m, _ = chay_ocr_copy(dc_oc, dich, ra, api, 2, 30, "ocr", di_chuyen=True, cache=cache, bbt=bbt, log=im, bo_khac=False, lay_lech=False)
        assert not os.path.exists(s1) and not os.path.exists(w1)
        assert {v for v, *_ in ke_hoach_hoan_tac(xl_m)} == {"bỏ qua", "xoá"}  # chỉ xoá bản copy trong KHONG_NGAY
        kh_m = ke_hoach_hoan_tac(xl_m.replace("ocr_copy_", "nhat_ky_chep_").replace(".xlsx", ".csv"))
        # 2 dòng Excel cùng trỏ một file: dòng đầu đã di chuyển file đi, dòng sau không có gì để trả về
        assert sorted(v for v, *_ in kh_m) == ["bỏ qua"] * 2 + ["trả về"] * 3 + ["xoá"], kh_m  # KHONG_NGAY luôn copy
        hoan_tac(kh_m, im)
        assert nam_dich() == [] and all(os.path.isfile(p) for p in (s1, w1, s2, c_pdf, os.path.join(a, "anh.jpg")))

        # Mặc định: tên theo Excel, sách tên OCR KHÁC tên Excel không chép (Sách một / LỊCH SỬ, Sách hai / SÁCH CHỜ)
        xl_k, _ = chay_ocr_copy(dc_oc, os.path.join(tmp, "dich_k"), ra, api, 2, 30, thu=True, cache=cache, bbt=bbt, log=im, lay_lech=False)
        rows = {r["ten_excel"]: r for r in sxb_tool._doc_sheet(xl_k, "OCR & copy", [(k, h) for k, h, _ in COT_OCR_COPY])}
        assert [rows[k]["trang_thai"] for k in ("Sách một", "Sách hai", "Con")] == ["CẦN XEM", KHONG_RA_NGAY, KHONG_RA_NGAY], rows
        assert (rows["Sách một"]["ten_khop"], rows["Con"]["ten_khop"], rows["Con"]["ten_moi"]) == ("KHÁC", "", "con_R_XX_M00")
        assert rows["Sách một"]["ten_moi"] == "sach_mot_0321_R_D_M00" and rows["Sách một"]["nguon_ten"] == "Excel"

        # Tab 5 Cần kiểm tra: chép mọi file của sách cần xem vào CAN_KIEM_TRA/<loại>/, mỗi sách một loại
        kt = os.path.join(tmp, "kt")
        xl_t, tk_t = chay_can_kiem_tra(xl_k, dc_oc, kt, thu=True, log=im)
        assert not os.path.exists(kt) and tk_t["chep"] == 0
        xl_kt, tk_kt = chay_can_kiem_tra(xl_k, dc_oc, kt, log=im)
        goc_kt = os.path.join(kt, THU_MUC_KIEM_TRA)
        assert os.path.dirname(xl_kt) == goc_kt and sorted(os.listdir(goc_kt)) == \
            sorted(["1_TEN_OCR_KHAC", "2_KHONG_CO_FILE_RUOT", "4_OCR_LOI", "KHONG_NGAY", os.path.basename(xl_kt)]), os.listdir(goc_kt)
        ds_kt = {r["ten_excel"]: r for r in sxb_tool._doc_sheet(xl_kt, "1_TEN_OCR_KHAC", [(k, h) for k, h, _ in COT_KIEM_TRA])}
        assert sorted(ds_kt) == ["Sách một", "Sách một bản 2"], ds_kt  # Sách hai không ra ngày: vào KHONG_NGAY
        mot = ds_kt["Sách một"]
        co = sorted(f for _, _, fs in os.walk(mot["thu_muc"]) for f in fs)
        assert co == sorted(["_thong_tin.txt", "anh.jpg", os.path.basename(s1), os.path.basename(w1)]), co
        assert mot["trang_thai"] == "ĐÃ CHÉP" and mot["so_file"] == 3 and os.path.isfile(s1)  # chỉ copy
        con = sxb_tool._doc_sheet(xl_kt, "4_OCR_LOI", [(k, h) for k, h, _ in COT_KIEM_TRA])[0]
        assert con["ten_excel"] == "Con" and con["ly_do_khac"] == "OCR không ra tên", con  # OCR lỗi không tính KHONG_NGAY
        hai = sxb_tool._doc_sheet(xl_kt, "KHONG_NGAY", [(k, h) for k, h, _ in COT_KIEM_TRA])[0]
        assert hai["ten_excel"] == "Sách hai" and os.path.basename(os.path.dirname(hai["thu_muc"])) == "2018", hai
        mat = sxb_tool._doc_sheet(xl_kt, "2_KHONG_CO_FILE_RUOT", [(k, h) for k, h, _ in COT_KIEM_TRA])[0]
        assert mat["ten_excel"] == "Mất file" and "không còn file Mat.pdf" in mat["ghi_chu"], mat
        tong = sxb_tool._doc_sheet(xl_kt, "Tổng hợp", [("ma", "Thư mục"), ("so", "Số sách")])
        assert [(r["ma"], r["so"]) for r in tong[:7]] == [("TRUOC_2006", 0), ("KHONG_NGAY", 1), ("1_TEN_OCR_KHAC", 2),
                                                          ("2_KHONG_CO_FILE_RUOT", 1), ("3_LOI_CHEP", 0), ("4_OCR_LOI", 1),
                                                          ("5_OCR_KHONG_RA_TEN", 0)], tong
        assert chay_can_kiem_tra(xl_k, dc_oc, kt, log=im)[1]["chep"] == 0  # chạy lại: file đã có thì bỏ qua
        # Chọn loại khác: Con chỉ vào 5_OCR_KHONG_RA_TEN khi bỏ chọn 4_OCR_LOI
        _, kh_kt = ke_hoach_kiem_tra(doc_ket_qua_copy(xl_k), sxb_tool.doc_doi_chieu(dc_oc), "x", ["5_OCR_KHONG_RA_TEN"])
        assert {x["ten_excel"] for x, _, _ in kh_kt} == {"Con", "Mất file"}
        # Năm lưu chiểu OCR trước 2006: xét trước mọi loại; Excel cũ (ngày theo QĐXB Excel) thì năm Excel
        cu = [dict(ten_excel="OCR cũ", ten_ocr="A", nam_excel=2010, nam_lc=2005, dong=1, trang_thai="ĐÃ LƯU"),
              dict(ten_excel="Excel cũ", ten_ocr="B", nam_excel=2005, nam_lc="", nguon_ngay="Excel (QĐXB)", dong=2,
                   trang_thai="ĐÃ LƯU"),
              dict(ten_excel="Mới", ten_ocr="C", nam_excel=2006, nam_lc=2008, dong=3, trang_thai="ĐÃ LƯU"),
              dict(ten_excel="Mốc 2006", ten_ocr="F", nam_excel=2010, nam_lc=2006, dong=7, trang_thai="ĐÃ LƯU"),  # 2006 hợp lệ
              dict(ten_excel="Khác cũ", ten_ocr="D", nam_excel=2010, nam_lc=2004, dong=4, trang_thai="CẦN XEM"),
              dict(ten_excel="Không tên cũ", ten_ocr="", nam_excel=2023, nam_lc=1951, dong=5, trang_thai="ĐÃ LƯU"),
              dict(ten_excel="Không ngày", ten_ocr="E", nam_excel=2004, nam_lc="", nguon_ngay=KHONG_RA_NGAY, dong=6,
                   trang_thai=KHONG_RA_NGAY)]
        d_cu, _ = ke_hoach_kiem_tra(cu, [], "x")
        assert [(x["ten_excel"], x["ma"], x["ly_do_khac"]) for x in d_cu] == [
            ("OCR cũ", "TRUOC_2006", ""), ("Excel cũ", "TRUOC_2006", "Không ra ngày"),
            ("Khác cũ", "TRUOC_2006", "Tên OCR không khớp Excel"), ("Không tên cũ", "TRUOC_2006", "OCR không ra tên"),
            ("Không ngày", "KHONG_NGAY", "")], d_cu  # năm Excel 2004 không tính: không dùng ngày Excel
        assert [os.path.basename(os.path.dirname(x["thu_muc"])) for x in d_cu] == ["2005", "2005", "2004", "1951", "2004"], d_cu
        assert MA_TRUOC == "TRUOC_2006" and TRANG_THAI_TRUOC == "TRƯỚC 2006"

        # Không ra tên từ Word / PDF -> OCR bìa (ưu tiên bìa 1); sách chỉ có bìa cũng OCR bìa; không có bìa thì thôi
        tb = os.path.join(tmp, "bia_test"); os.makedirs(tb)
        tp = {n: os.path.join(tb, n) for n in ("Trong ten.pdf", "Trong ten bia 4.jpg", "Trong ten bia 1.jpg",
                                                "Chi bia.png", "Trong khong bia.pdf")}
        for p in tp.values():
            open(p, "wb").write(b"x")
        ds = lambda *ns: sxb_tool._ds_file([(ngay, tp[n]) for n in ns])
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "KHOP"; ws.append(hdr)
        ws.append(dong_kq(nam=2022, dong=1, ten="Có bìa", qdxb="01/2022", tinh_trang="KHỚP", file_pdf=tp["Trong ten.pdf"],
                          cac_file=ds("Trong ten.pdf", "Trong ten bia 4.jpg", "Trong ten bia 1.jpg")))
        ws.append(dong_kq(nam=2022, dong=2, ten="Sách lẻ", qdxb="02/2022", tinh_trang="KHỚP", cac_file=ds("Chi bia.png")))
        ws.append(dong_kq(nam=2022, dong=3, ten="Không bìa", qdxb="03/2022", tinh_trang="KHỚP",
                          file_pdf=tp["Trong khong bia.pdf"], cac_file=ds("Trong khong bia.pdf")))
        dc_bia = os.path.join(tmp, "doi_chieu_bia.xlsx"); wb.save(dc_bia)
        cache_b = os.path.join(tmp, "ra_bia", "ocr_cache.jsonl")
        nhan.clear()
        xl_b, _ = chay_ocr_copy(dc_bia, os.path.join(tmp, "dich_b"), os.path.join(tmp, "ra_bia"), api, 2, 30, "ocr",
                                thu=True, cache=cache_b, log=im)
        gui_bia = sorted(os.path.basename(f["bia"]) for _, f in nhan if f.get("bia"))
        assert gui_bia == ["Chi bia.png", "Trong ten bia 1.jpg"], nhan  # bìa 1 trước bìa 4, cuốn không bìa không gửi
        rows = {r["ten_excel"]: r for r in sxb_tool._doc_sheet(xl_b, "OCR & copy", [(k, h) for k, h, _ in COT_OCR_COPY])}
        assert (rows["Có bìa"]["ten_ocr"], rows["Có bìa"]["nguon_ten"]) == ("TÊN TỪ BÌA", "OCR (bìa)"), rows
        assert rows["Sách lẻ"]["ten_ocr"] == "TÊN TỪ BÌA" and rows["Không bìa"]["ten_ocr"] == ""
        assert rows["Có bìa"]["thang_lc"] == 3 and rows["Có bìa"]["nguon_ngay"] == "OCR"  # thời gian vẫn từ PDF
        nhan.clear()
        chay_ocr_copy(dc_bia, os.path.join(tmp, "dich_b"), os.path.join(tmp, "ra_bia"), api, 2, 30, "ocr",
                      thu=True, cache=cache_b, log=im)
        assert nhan == [], nhan  # chạy lại: OCR, OCR bìa đều lấy từ cache
        assert chon_bia([tp["Trong ten bia 4.jpg"], os.path.join(tb, "a.ai"), tp["Trong ten bia 1.jpg"]]) == tp["Trong ten bia 1.jpg"]

        # Bìa = tên có chữ Bìa (mọi định dạng) + mọi file .jpg + PDF tên kết thúc '- số'; file khác (.png .ai .psd...) thì không
        def thu_bia(ten_file, ten_sach="Sách", ruot="Ruot.pdf"):
            d = os.path.join(tmp, f"bia_chat_{len(os.listdir(tmp))}"); os.makedirs(d)
            ps = [os.path.join(d, n) for n in [ruot] + ten_file]
            for p in ps:
                os.makedirs(os.path.dirname(p), exist_ok=True)
                open(p, "wb").write(b"x")
            r = {"cac_file": sxb_tool._ds_file([(ngay, p) for p in ps]), "file_word": "", "file_pdf": ps[0], "ten": ten_sach}
            return sorted(os.path.relpath(p, d) for p in ds_bia_chat(r)), [os.path.join(d, n) for n in ten_file]
        assert thu_bia(["Bia.jpg", "1.jpg", "30-31.JPG", "Bang1.eps", "cover.png", "anh.jpeg"])[0] == \
            sorted(["Bia.jpg", "1.jpg", "30-31.JPG", "anh.jpeg"])
        assert thu_bia(["Tu tuong HCM - 01.pdf", "Tu tuong HCM -2.pdf", "Phan 2.pdf", "BiaSach-LTTT.png", "bia_sau.ai",
                        "Bìa trước.psd", "Colombia.png", "To che.pdf", os.path.join("con", "Bia 4.tif"), os.path.join("con", "x.jpg")])[0] \
            == sorted(["Tu tuong HCM - 01.pdf", "Tu tuong HCM -2.pdf", "BiaSach-LTTT.png", "bia_sau.ai", "Bìa trước.psd",
                       os.path.join("con", "Bia 4.tif")])
        # Tên sách có chữ Bia: tên file có 'bia' không còn là dấu hiệu bìa
        assert thu_bia(["Bia da van mieu phan 2.doc", "Bia 1.jpg"], ten_sach="Bia đá Văn Miếu")[0] == ["Bia 1.jpg"]
        _, ps = thu_bia(["a.jpg", "x - 01.pdf", "Bia 4.jpg", "Bia 1.jpg", "khac.ai"])
        assert os.path.basename(chon_bia(ps)) == "Bia 1.jpg"
        assert os.path.basename(chon_bia([p for p in ps if "Bia" not in p])) == "x - 01.pdf"
        # Đủ cả 3 loại bìa: chép hết vào Bia/, OCR chọn bìa 1 (kể cả khi có PDF bìa 4)
        bia3, ps3 = thu_bia(["Lich su.doc", "Bia 1.ai", "Bia 4.pdf", "Bia 1.jpg", "anh 3.jpg", "Lich su - 01.pdf"])
        assert bia3 == sorted(["Bia 1.ai", "Bia 4.pdf", "Bia 1.jpg", "anh 3.jpg", "Lich su - 01.pdf"]), bia3
        assert os.path.basename(chon_bia(ps3)) == "Bia 1.jpg"

        # Năm OCR trước 2006: không chép vào đích (kể cả chế độ di chuyển), COPY mọi file liên quan vào
        # <đích>/CAN_KIEM_TRA/TRUOC_2006/<năm OCR>/<năm Excel>_<dòng>_<tên>/; sách khác vẫn di chuyển bình thường
        tc = os.path.join(tmp, "truoc"); os.makedirs(os.path.join(tc, "cu")); os.makedirs(os.path.join(tc, "moi"))
        cu = {n: os.path.join(tc, "cu", n) for n in ("Cu.pdf", "Cu.docx", "anh.jpg")}
        moi_pdf = os.path.join(tc, "moi", "Moi.pdf")
        for p in list(cu.values()) + [moi_pdf]:
            open(p, "wb").write(b"x")
        n22 = "2022-01-01 08:00:00"
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "KHOP"; ws.append(hdr)
        ws.append(dong_kq(nam=2022, dong=11, ten="Sách cũ", tinh_trang="KHỚP", file_pdf=cu["Cu.pdf"], file_word=cu["Cu.docx"],
                          cac_file=sxb_tool._ds_file([(n22, p) for p in cu.values()])))
        ws.append(dong_kq(nam=2022, dong=12, ten="Lịch sử", tinh_trang="KHỚP", file_pdf=moi_pdf,
                          cac_file=sxb_tool._ds_file([(n22, moi_pdf)])))
        dc_tc = os.path.join(tmp, "doi_chieu_truoc.xlsx"); wb.save(dc_tc)
        dich_tc, ra_tc = os.path.join(tmp, "dich_tc"), os.path.join(tmp, "ra_tc")
        os.makedirs(dich_tc)
        xl_tc, _ = chay_ocr_copy(dc_tc, dich_tc, ra_tc, api, 2, 30, di_chuyen=True, cache=os.path.join(ra_tc, "c.jsonl"), log=im)
        rows = {r["ten_excel"]: r for r in sxb_tool._doc_sheet(xl_tc, "OCR & copy", [(k, h) for k, h, _ in COT_OCR_COPY])}
        assert (rows["Sách cũ"]["trang_thai"], rows["Lịch sử"]["trang_thai"]) == ("TRƯỚC 2006", "ĐÃ LƯU"), rows
        tm_cu = os.path.join(dich_tc, THU_MUC_KIEM_TRA, "TRUOC_2006", "1951", "2022_0011_sach_cu")
        assert sorted(os.listdir(tm_cu)) == ["Cu.docx", "Cu.pdf", "_thong_tin.txt", "anh.jpg"], os.listdir(tm_cu)
        assert all(os.path.isfile(p) for p in cu.values()) and not os.path.exists(moi_pdf)  # sách cũ chỉ copy
        assert sorted(os.listdir(dich_tc)) == ["2021", THU_MUC_KIEM_TRA] and "1951" not in os.listdir(dich_tc)
        assert any(f.startswith("can_kiem_tra_tab4_") for f in os.listdir(os.path.join(dich_tc, THU_MUC_KIEM_TRA)))
        # Tab 5 chạy lại trên kết quả này: cùng thư mục, file đã có thì bỏ qua
        assert chay_can_kiem_tra(xl_tc, dc_tc, dich_tc, [MA_TRUOC], log=im)[1]["chep"] == 0
        # Hoàn tác theo nhật ký: xoá bản copy trong CAN_KIEM_TRA, trả sách di chuyển về chỗ cũ
        kh_tc = ke_hoach_hoan_tac(xl_tc.replace("ocr_copy_", "nhat_ky_chep_").replace(".xlsx", ".csv"))
        assert sorted(v for v, *_ in kh_tc) == ["trả về"] + ["xoá"] * 3, kh_tc
        hoan_tac(kh_tc, im)
        assert os.path.isfile(moi_pdf) and all(os.path.isfile(p) for p in cu.values())
        assert not os.path.exists(os.path.join(dich_tc, "2021")) and os.listdir(tm_cu) == ["_thong_tin.txt"], os.listdir(tm_cu)

        # File đối chiếu đã tách (chỉ CAN_XU_LY + CAN_OCR): sách LỆCH NĂM thử tối đa 3 cặp file (thư mục năm gần năm Excel
        # trước) đến khi tên KHỚP; PDF scan ghép vào sách theo tên, trùng thì bỏ, không khớp -> CAN_KIEM_TRA/KHONG_KHOP_SACH
        lc = os.path.join(tmp, "lech_scan")
        tep = {}
        for tm, ten_file in (("x1", "Ten_Sach_lech.pdf"), ("x2", "Ten_Sach_sai.pdf"), ("x3", "Ten_Sach_lech_2.pdf"),
                             ("x4", "anh.jpg"), ("s1", "Ten_Sach_khong_file.pdf"), ("s2", "Ten_Sach_lech.pdf"),
                             ("s3", "Ten_Sach_la.pdf")):
            os.makedirs(os.path.join(lc, tm), exist_ok=True)
            tep[tm] = os.path.join(lc, tm, ten_file)
            open(tep[tm], "wb").write(b"x")
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "CAN_XU_LY"; ws.append(hdr)
        ws.append(dong_kq(nam=2020, dong=31, ten="Sách lệch", tinh_trang="LỆCH NĂM - cần mở file kiểm tra",
                          cac_file=sxb_tool._ds_file([("2015-01-01 00:00:00", tep["x1"]), ("2019-01-01 00:00:00", tep["x2"]),
                                                      ("2010-01-01 00:00:00", tep["x3"])])))
        ws.append(dong_kq(nam=2020, dong=32, ten="Sách không file", tinh_trang="LỆCH NĂM - tìm theo lời nói đầu",
                          cac_file=sxb_tool._ds_file([("2019-01-01 00:00:00", tep["x4"])])))
        ws.append(dong_kq(nam=2020, dong=33, ten="Không tìm", tinh_trang="KHÔNG TÌM ĐƯỢC"))
        ws_o = wb.create_sheet("CAN_OCR"); ws_o.append([h for _, h in sxb_tool.COT_OCR])
        for k, lq in (("s1", "2020 - dòng 32 - Sách không file"), ("s2", ""), ("s3", "2020 - dòng 31 - Sách lệch")):
            ws_o.append([tep[k], k, os.path.basename(tep[k]), 10, "PDF ảnh scan, không có lớp chữ", "CAO", lq])
        ws_o.append([os.path.join(lc, "mat", "Mat_scan.pdf"), "mat", "Mat_scan.pdf", 3, "PDF ảnh scan", "CAO", ""])
        dc_lc = os.path.join(tmp, "doi_chieu_tach.xlsx"); wb.save(dc_lc)
        dich_lc, ra_lc = os.path.join(tmp, "dich_lc"), os.path.join(tmp, "ra_lc")
        nhan.clear()
        xl_lc, _ = chay_ocr_copy(dc_lc, dich_lc, ra_lc, api, 2, 30, cache=os.path.join(ra_lc, "c.jsonl"), log=im)
        gui_title = [os.path.basename(f.get("pdf", "")) for p, f in nhan if p.endswith("/title")]
        assert "Ten_Sach_lech_2.pdf" not in gui_title and gui_title.count("Ten_Sach_sai.pdf") == 1, gui_title  # dừng ở cặp 2
        rows = {r["ten_excel"]: r for r in sxb_tool._doc_sheet(xl_lc, "OCR & copy", [(k, h) for k, h, _ in COT_OCR_COPY])}
        assert sorted(rows) == ["Sách không file", "Sách lệch"], rows  # 'KHÔNG TÌM ĐƯỢC' không xử lý
        lech_ = rows["Sách lệch"]
        assert (lech_["trang_thai"], lech_["file_pdf"], lech_["ten_khop"]) == ("ĐÃ LƯU", tep["x1"], "KHỚP"), lech_
        assert "thử 2/3 cặp file, chọn Ten_Sach_lech.pdf (tên KHỚP)" in lech_["ghi_chu"], lech_["ghi_chu"]
        kf = rows["Sách không file"]
        assert (kf["trang_thai"], kf["file_pdf"]) == ("ĐÃ LƯU", tep["s1"]) and "PDF scan (CAN_OCR)" in kf["ghi_chu"], kf
        scan = {os.path.basename(os.path.dirname(r["path"])): r for r in sxb_tool._doc_sheet(
            xl_lc, "PDF scan", [(k, h) for k, h, _ in COT_SCAN])}
        assert {k: v["trang_thai"] for k, v in scan.items()} == {"s1": "ĐÃ GHÉP", "s2": "TRÙNG", "s3": "KHÔNG KHỚP SÁCH",
                                                                 "mat": "KHÔNG CÒN FILE"}, scan
        assert scan["s1"]["ghep"] == "2020 - dòng 32 - Sách không file" and scan["s3"]["ten_ocr"] == "Sach la"
        kk = os.path.dirname(scan["s3"]["dich"])
        assert kk.startswith(os.path.join(dich_lc, THU_MUC_KIEM_TRA, MA_KHONG_KHOP_SACH, "2021", "ten_sach_la_"))
        assert sorted(os.listdir(kk)) == ["Ten_Sach_la.pdf", "_thong_tin.txt"] and os.path.isfile(tep["s3"]), os.listdir(kk)
        assert sorted(os.listdir(os.path.join(dich_lc, "2021", "thang 03"))) == ["sach_khong_file_0321_R_XX_M00",
                                                                                "sach_lech_0321_R_XX_M00"]
        # Hoàn tác theo nhật ký: xoá cả bản copy scan trong KHONG_KHOP_SACH
        kh_lc = ke_hoach_hoan_tac(xl_lc.replace("ocr_copy_", "nhat_ky_chep_").replace(".xlsx", ".csv"))
        assert sorted(v for v, *_ in kh_lc) == ["xoá"] * 3, kh_lc
        hoan_tac(kh_lc, im)
        assert not os.path.exists(os.path.join(kk, "Ten_Sach_la.pdf")) and not os.path.exists(os.path.join(dich_lc, "2021"))

        # Chỉ chép sách KHỚP: tên PDF ('LỊCH SỬ') không khớp -> OCR bìa; bìa ('TÊN TỪ BÌA') khớp thì chép theo bìa,
        # không có bìa / bìa cũng không khớp thì CẦN XEM; tên PDF khớp sẵn thì không OCR bìa
        kb = os.path.join(tmp, "khop_bia")
        ruot = {}
        for ten_tm, ten_file in (("a", "Ruot A.pdf"), ("b", "Ruot B.pdf"), ("c", "Ruot C.pdf")):
            os.makedirs(os.path.join(kb, ten_tm))
            ruot[ten_tm] = os.path.join(kb, ten_tm, ten_file)
            open(ruot[ten_tm], "wb").write(b"x")
        bia_a = os.path.join(kb, "a", "Bia 1.jpg"); open(bia_a, "wb").write(b"x")
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "KHOP"; ws.append(hdr)
        ws.append(dong_kq(nam=2021, dong=21, ten="Tên từ bìa", tinh_trang="KHỚP", file_pdf=ruot["a"],
                          cac_file=sxb_tool._ds_file([(ngay, ruot["a"]), (ngay, bia_a)])))
        ws.append(dong_kq(nam=2021, dong=22, ten="Lịch sử Việt Nam", tinh_trang="KHỚP", file_pdf=ruot["b"],
                          cac_file=sxb_tool._ds_file([(ngay, ruot["b"])])))
        ws.append(dong_kq(nam=2021, dong=23, ten="Lịch sử", tinh_trang="KHỚP", file_pdf=ruot["c"],
                          cac_file=sxb_tool._ds_file([(ngay, ruot["c"])])))
        dc_kb = os.path.join(tmp, "doi_chieu_khop_bia.xlsx"); wb.save(dc_kb)
        nhan.clear()
        xl_kb, _ = chay_ocr_copy(dc_kb, os.path.join(tmp, "dich_kb"), os.path.join(tmp, "ra_kb"), api, 2, 30, "ocr",
                                 cache=os.path.join(tmp, "ra_kb", "c.jsonl"), log=im)
        assert [os.path.basename(f["bia"]) for _, f in nhan if f.get("bia")] == ["Bia 1.jpg"], nhan  # B không có bìa, C khớp sẵn
        rows = {r["ten_excel"]: r for r in sxb_tool._doc_sheet(xl_kb, "OCR & copy", [(k, h) for k, h, _ in COT_OCR_COPY])}
        a_ = rows["Tên từ bìa"]
        assert (a_["ten_khop"], a_["ten_ocr"], a_["nguon_ten"], a_["trang_thai"]) == ("KHỚP", "TÊN TỪ BÌA", "OCR (bìa)", "ĐÃ LƯU"), a_
        assert "tên lấy theo bìa (Word / PDF ra 'LỊCH SỬ')" in a_["ghi_chu"]
        assert (rows["Lịch sử Việt Nam"]["ten_khop"], rows["Lịch sử Việt Nam"]["trang_thai"]) == ("KHÁC", "CẦN XEM")
        assert (rows["Lịch sử"]["ten_khop"], rows["Lịch sử"]["trang_thai"]) == ("KHỚP", "ĐÃ LƯU")
        # ap_dung_ocr: tên gần giống vẫn là KHÁC (không chép); tên bìa KHỚP thì lấy tên bìa
        r = {"ten": "Giáo trình kinh tế chính trị Mác Lênin", "nam": 2021, "qdxb": "", "bbt": "", "lan": ""}
        ap_dung_ocr(r, {"ten_sach": "Giáo trình Kinh tế chính trị Mác - Lê-nin", "thang": 1, "nam": 2021})
        assert r["ten_khop"] == "KHÁC"
        ap_dung_ocr(r, {"ten_sach": "Giáo trình Kinh tế chính trị Mác - Lê-nin", "ten_bia": "GIÁO TRÌNH KINH TẾ CHÍNH TRỊ MÁC LÊNIN",
                        "thang": 1, "nam": 2021})
        assert (r["ten_khop"], r["nguon_ten"]) == ("KHỚP", "OCR (bìa)")

        # chỉ KHỚP / KHÁC: viết thường, bỏ dấu, bỏ ký tự đặc biệt, bỏ chú thích rồi so bằng nhau
        assert so_ten("Văn kiện Đảng toàn tập- Tập43 (1982)", "VĂN KIỆN ĐẢNG TOÀN TẬP TẬP 43") == "KHÁC"
        assert so_ten("Văn kiện Đảng toàn tập- Tập43 (1982)", "Văn kiện Đảng toàn tập Tập 48") == "KHÁC"
        assert so_ten("Lịch sử Đảng bộ tỉnh Nam Định", "LỊCH SỬ ĐẢNG BỘ TỈNH NAM ĐỊNH") == "KHỚP"
        assert so_ten("Hồ Chí Minh với báo chí", "Hồ Chí Minh với báo chí (Sách chuyên khảo)") == "KHỚP"
        assert so_ten("Một số luận cứ phản bác quan điểm sai trái, thù địch", "MỘT SỐ LUẬN CỨ PHẢN BÁC QUAN ĐIỂM SAI TRÁI THỦ ĐỊCH") == "KHỚP"
        assert so_ten("Trường - Chinh tuyển tập", "Trường-Chinh Tuyển tập") == "KHỚP"
        assert so_ten("Kinh tế số Việt Nam", "Kinh tế số Việt Nam: Thực trạng và giải pháp") == "KHÁC"
        assert so_ten("Kinh tế số Việt Nam: Thực trạng và giải pháp", "KINH TẾ SỐ VIỆT NAM") == "KHÁC"
        assert so_ten("Lịch sử Đảng bộ và nhân dân xã Nam Vân", "LỊCH SỬ ĐẢNG BỘ VÀ NHÂN DÂN XÃ AN LĨNH") == "KHÁC"
        assert so_ten("Giáo trình kinh tế chính trị Mác Lênin", "Giáo trình Kinh tế chính trị Mác - Lê-nin") == "KHÁC"
        assert so_ten("Sách bất kỳ", "") == ""

        r = {"ten": "Sách Excel", "nam": 2020, "qdxb": "01/2020", "bbt": "", "lan": "TB"}
        ap_dung_ocr(r, {"ten_sach": "Sách OCR", "thang": 7, "nam": 2022}, theo_ten="excel")
        assert (r["ten_moi"], r["nam"], r["nguon_ten"]) == ("sach_excel_0722_R_XX_T01", 2022, "Excel")
        r = {"ten": "A B", "nam": 2020, "qdxb": "01/2020", "bbt": "", "lan": ""}
        ap_dung_ocr(r, {"ten_sach": "", "thang": 7, "nam": 1802})  # năm OCR vô lý -> không ra ngày, không dùng QĐXB Excel
        assert (r["nam"], r["nguon_ngay"], r["qdxb"], r["ten_moi"]) == (2020, KHONG_RA_NGAY, "", "a_b_R_XX_M00")
        ap_dung_ocr(r, {"ten_sach": "C D", "nam": 2023})  # chỉ có năm -> thang 00, tên không có mmyy
        assert (r["nam"], sxb_tool.thang_qdxb(r["qdxb"]), r["ten_moi"]) == (2023, "00", "c_d_R_XX_M00")
    finally:
        sv.shutdown()
        shutil.rmtree(tmp, ignore_errors=True)
    print("OK OCR")


if __name__ == "__main__":
    import multiprocessing; multiprocessing.freeze_support()  # exe: tiến trình con đọc PDF song song
    _test() if "--test" in sys.argv else gui()
