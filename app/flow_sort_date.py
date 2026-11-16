# Luồng lấy tháng, năm lưu chiểu ở trang cuối
import os
import re
import json
import time
import base64
import logging
import tempfile
import subprocess
import unicodedata

import cv2
import numpy as np
import requests
from pdf2image import convert_from_path, pdfinfo_from_path
from dotenv import load_dotenv

from .logger import get_logger

load_dotenv()


# Cấu hình Chandra, Spellfix, PDF

CHANDRA_URL        = os.getenv("CHANDRA_URL",   "http://localhost:8101/v1/chat/completions")
CHANDRA_MODEL      = os.getenv("CHANDRA_MODEL", "datalab-to/chandra-ocr-2")
CHANDRA_MAX_TOKENS = int(os.getenv("SORT_CHANDRA_MAX_TOKENS", "2400"))
CHANDRA_WORKERS    = int(os.getenv("SORT_CHANDRA_WORKERS", "4"))
CHANDRA_TIMEOUT    = int(os.getenv("SORT_CHANDRA_TIMEOUT", "600"))
MAX_LONG_SIDE      = int(os.getenv("SORT_CHANDRA_MAX_SIDE", "1800"))

SPELLFIX_API_URL = os.getenv("SPELLFIX_API_URL", "http://localhost:8006/spellfix/correct")
SPELLFIX_ENABLED = os.getenv("SPELLFIX_ENABLED", "1").lower() in ("1", "true", "t", "yes")

PDF_DPI = int(os.getenv("SORT_PDF_DPI", "200"))


# Gọi POST JSON, ghi log khi lỗi
def _post(service: str, url: str, payload: dict, timeout: int, log, job_id: str = ""):
    t0 = time.time()
    try:
        resp = requests.post(url, json=payload, timeout=timeout)
    except (requests.ConnectionError, requests.Timeout) as e:
        get_logger("connection", job_id or None).error(
            f"{service} {url}: không kết nối được sau {time.time() - t0:.1f}s: {e}")
        raise
    if resp.status_code >= 400:
        log.error(f"{service} HTTP {resp.status_code} sau {time.time() - t0:.1f}s: {resp.text[:300]}")
    resp.raise_for_status()
    return resp


# Đếm số trang PDF
def pdf_page_count(pdf_path: str) -> int:
    return int(pdfinfo_from_path(pdf_path)["Pages"])


# Đọc Word, đổi TCVN3

_TCVN3_PAIRS = [
    ("µ¸¶·¹",   "àáảãạ"),
    ("¨»¾¼½Æ",  "ăằắẳẵặ"),
    ("©ÇÊÈÉË",  "âầấẩẫậ"),
    ("®",        "đ"),
    ("ÌÐÎÏÑ",   "èéẻẽẹ"),
    ("ªÒÕÓÔÖ",  "êềếểễệ"),
    ("×ÝØÜÞ",   "ìíỉĩị"),
    ("ßãáâä",   "òóỏõọ"),
    ("«åèæçé",  "ôồốổỗộ"),
    ("¬êíëìî",  "ơờớởỡợ"),
    ("ïóñòô",   "ùúủũụ"),
    ("\xadõøö÷ù", "ưừứửữự"),
    ("úýûüþ",   "ỳýỷỹỵ"),
    ("¡¢£¤¥¦§",  "ĂÂÊÔƠƯĐ"),
]
TCVN3_CHAR_MAP = {t: u for ts, us in _TCVN3_PAIRS for t, u in zip(ts, us)}

_VN_LATIN1 = set("ÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚÝàáâãèéêìíòóôõùúý")
TCVN3_MARKERS = {c for c in TCVN3_CHAR_MAP if c not in _VN_LATIN1} - {"\xad"}

_VN_UNICODE_ONLY = re.compile("[ĂăĐđĨĩŨũƠơƯư\u1EA0-\u1EF9]")

_TCVN3_FONT_RE = re.compile(r"^\.vn|tcvn", re.I)


# Kiểm tra text có phải mã TCVN3 không
def looks_like_tcvn3(text: str) -> bool:
    if not text:
        return False
    letters = sum(c.isalpha() for c in text)
    markers = sum(c in TCVN3_MARKERS for c in text)
    unicode_vn = len(_VN_UNICODE_ONLY.findall(text))
    return markers >= 2 and markers >= 0.02 * letters and unicode_vn * 5 <= markers


# Đổi text TCVN3 sang Unicode
def convert_tcvn3_text(text: str, assume=None) -> str:
    if not text:
        return text
    if assume is None:
        assume = looks_like_tcvn3(text)
    if not assume:
        return text
    out = []
    for line in text.split("\n"):
        if _VN_UNICODE_ONLY.search(line):
            out.append(line)
        else:
            out.append("".join(TCVN3_CHAR_MAP.get(c, c) for c in line))
    return "\n".join(out)

# Âm tiết tiếng Việt hợp lệ (đã bỏ dấu thanh): chấm text đổi mã có ra tiếng Việt thật không
_DAU_THANH = {"\u0300", "\u0301", "\u0303", "\u0309", "\u0323"}
_VAN = ("a|ă|â|e|ê|i|o|ô|ơ|u|ư|y|ai|ao|au|ay|âu|ây|eo|êu|ia|iê|iu|oa|oă|oe|oi|ôi|ơi|oo|ua|uâ|uê|ui|uô|uơ|uy"
        "|ưa|ưi|ưu|ươ|iêu|oai|oay|oeo|uây|uôi|uya|uyê|uyu|ươi|ươu|yê|yêu")
_AM_TIET_RE = re.compile(rf"^(ngh|ng|nh|ch|gh|gi|kh|ph|qu|th|tr|[bcdđghklmnprstvx])?({_VAN})(ch|ng|nh|[cmnpt])?$")
_TU_RE = re.compile(r"[^\W\d_]+")
VN_MIN_SCORE = 0.5  # tỷ lệ âm tiết hợp lệ tối thiểu để coi là tiếng Việt đọc được
# Chữ Latin-1 không có trong tiếng Việt: còn nhiều sau khi đọc là phông cũ (TCVN3, VNI, VPS...) chưa đổi được
_KY_TU_MA_CU = ({chr(c) for c in range(0xC0, 0x100)} - set("ÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚÝàáâãèéêìíòóôõùúý×÷")) | TCVN3_MARKERS


def _am_tiet_hop_le(tu: str) -> bool:
    d = unicodedata.normalize("NFD", tu.lower())
    if sum(c in _DAU_THANH for c in d) > 1:
        return False
    return bool(_AM_TIET_RE.match(unicodedata.normalize("NFC", "".join(c for c in d if c not in _DAU_THANH))))


# Tỷ lệ từ là âm tiết tiếng Việt hợp lệ
def vn_score(text: str) -> float:
    tu = _TU_RE.findall(text)
    return sum(map(_am_tiet_hop_le, tu)) / len(tu) if tu else 0.0


# Chỉ nhận bản đổi TCVN3 khi nó ra tiếng Việt hơn bản gốc (file mã khác như VPS, VietWare bị nhận nhầm là TCVN3)
def _nen_doi_tcvn3(goc: str, moi: str) -> bool:
    if moi == goc:
        return False
    diem = vn_score(moi)
    return diem >= VN_MIN_SCORE and diem > vn_score(goc)


# Đổi phông cũ (TCVN3 / VNI) nếu kết quả là tiếng Việt thật, không thì giữ nguyên
def fix_tcvn3_text(text: str) -> str:
    f = chon_bang_ma(text)
    return f(text) if f else text


# Bảng mã VNI-Windows: chữ gốc + ký tự dấu đứng sau ('HÖÔÙNG DAÃN' = 'HƯỚNG DẪN', 'ñ' = 'đ', 'ô' = 'ơ', 'ö' = 'ư')
_VNI_THANH = ("\u0301", "\u0300", "\u0309", "\u0303", "\u0323")  # sắc, huyền, hỏi, ngã, nặng


def _vni_map() -> dict:
    m = {}

    # goc: chữ gốc VNI; mu: ký tự mũ / trăng (không dấu thanh); duoi: ký tự dấu đứng sau theo thứ tự
    # sắc, huyền, hỏi, ngã, nặng (với chữ có mũ / trăng, một ký tự gộp cả mũ lẫn dấu thanh)
    def them(goc, chu, duoi, mu=""):
        for g, c, d, u in ((goc, chu, duoi, mu), (goc.upper(), chu.upper(), duoi.upper(), mu.upper()),
                           (goc.upper(), chu.upper(), duoi, mu)):  # chữ hoa đầu từ, dấu viết thường
            if g + u != c:
                m[g + u] = c
            for k, t in zip(d, _VNI_THANH):
                m[g + k] = unicodedata.normalize("NFC", c + t)
    for v in "aeou":
        them(v, v, "ùøûõï")
    them("y", "y", "ùøûõ")
    them("a", "â", "áàåãä", "â")
    them("e", "ê", "áàåãä", "â")
    them("o", "ô", "áàåãä", "â")
    them("a", "ă", "éèúüë", "ê")
    them("ô", "ơ", "ùøûõï")
    them("ö", "ư", "ùøûõï")
    m.update({"ñ": "đ", "Ñ": "Đ", "æ": "ỉ", "Æ": "Ỉ", "ó": "ĩ", "Ó": "Ĩ", "ò": "ị", "Ò": "Ị", "î": "ỵ", "Î": "Ỵ"})
    return m


VNI_MAP = _vni_map()
_VNI_RE = re.compile("|".join(map(re.escape, sorted(VNI_MAP, key=len, reverse=True))))


# Đổi text VNI sang Unicode (bỏ qua dòng đã có chữ Unicode tiếng Việt)
def convert_vni_text(text: str) -> str:
    if not text:
        return text
    return "\n".join(line if _VN_UNICODE_ONLY.search(line) else _VNI_RE.sub(lambda x: VNI_MAP[x.group(0)], line)
                     for line in text.split("\n"))


# Chọn cách đổi phông cũ (TCVN3 hoặc VNI) cho ra tiếng Việt thật nhất; None nếu text đã ổn / không đổi được
def chon_bang_ma(text: str):
    tot, diem = None, 0.0
    for f in (lambda t: convert_tcvn3_text(t, assume=True), convert_vni_text):
        moi = f(text)
        if _nen_doi_tcvn3(text, moi) and vn_score(moi) > diem:
            tot, diem = f, vn_score(moi)
    return tot


try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:
    Document = None

VIRTUAL_PAGE_CHARS = 2500
MAX_CHARS_PER_PAGE_READ = 6000


# Chia text thành các trang ảo
def _split_virtual_pages(text: str, n_pages: int, from_end: bool = False) -> list:
    n_pages = max(int(n_pages), 1)
    text = text.strip()
    if from_end:
        tail = text[-n_pages * VIRTUAL_PAGE_CHARS:]
        pages = [tail[max(0, i - VIRTUAL_PAGE_CHARS):i].strip() for i in range(len(tail), 0, -VIRTUAL_PAGE_CHARS)]
        return [p for p in pages if p][:n_pages][::-1]
    head = text[: n_pages * VIRTUAL_PAGE_CHARS]
    pages = [head[i:i + VIRTUAL_PAGE_CHARS].strip() for i in range(0, len(head), VIRTUAL_PAGE_CHARS)]
    return [p for p in pages if p][:n_pages]


# Lấy tên thẻ XML bỏ namespace
def _local(el) -> str:
    return el.tag.rsplit("}", 1)[-1] if isinstance(el.tag, str) else ""


# Duyệt đoạn văn và bảng theo thứ tự
def _iter_blocks(parent):
    for child in parent.iterchildren():
        tag = _local(child)
        if tag in ("p", "tbl"):
            yield tag, child
        elif tag == "sdt":
            content = child.find(qn("w:sdtContent"))
            if content is not None:
                yield from _iter_blocks(content)


# Xác định vị trí ngắt trang trong đoạn
def _page_break_mode(p_el):
    has_break = False
    pbb = p_el.find(f"./{qn('w:pPr')}/{qn('w:pageBreakBefore')}")
    if pbb is not None and pbb.get(qn("w:val"), "1").lower() not in ("0", "false", "off"):
        has_break = True
    text_before = text_after = False
    for el in p_el.iter():
        tag = el.tag
        if tag == qn("w:lastRenderedPageBreak") or (tag == qn("w:br") and el.get(qn("w:type")) == "page"):
            has_break = True
        elif tag == qn("w:t") and el.text and el.text.strip():
            if has_break:
                text_after = True
            else:
                text_before = True
    if not has_break:
        return None
    return "after" if (text_before and not text_after) else "before"


# Đoạn kết thúc section sang trang mới
def _ends_with_page_section(p_el) -> bool:
    sect = p_el.find(f"./{qn('w:pPr')}/{qn('w:sectPr')}")
    if sect is None:
        return False
    typ = sect.find(qn("w:type"))
    val = typ.get(qn("w:val")) if typ is not None else "nextPage"
    return val in ("nextPage", "oddPage", "evenPage")


# Đoạn có dùng font TCVN3 không
def _para_uses_tcvn3_font(para) -> bool:
    try:
        names = [r.font.name for r in para.runs if r.font.name]
        if not names and para.style is not None and para.style.font is not None and para.style.font.name:
            names = [para.style.font.name]
    except Exception:
        return False
    return any(_TCVN3_FONT_RE.search(n) for n in names)


# Lấy text một đoạn văn
def _para_text(para) -> str:
    text = para.text.strip()
    if text and _para_uses_tcvn3_font(para):
        return convert_tcvn3_text(text, assume=True)
    return text


# Lấy các đoạn trong text box
def _textbox_paragraphs(p_el, doc) -> list:
    found = p_el.xpath('.//w:txbxContent[not(ancestor::*[local-name()="Fallback"])]//w:p')
    return [Paragraph(e, doc) for e in found]


# Lấy text các dòng của bảng
def _table_lines(tbl_el, doc) -> list:
    lines = []
    for row in Table(tbl_el, doc).rows:
        seen, cells = set(), []
        for c in row.cells:
            if c._tc in seen:
                continue
            seen.add(c._tc)
            t = " ".join(x for x in (_para_text(p) for p in c.paragraphs) if x)
            if t:
                cells.append(t)
        if cells:
            lines.append(" | ".join(cells))
    return lines


# Đọc N trang đầu/cuối file .docx
def _extract_word_pages(docx_path: str, n_pages: int, from_end: bool = False) -> tuple:
    ext = os.path.splitext(docx_path)[1].lower()
    if ext == ".doc":
        return _extract_doc_text_as_pages(docx_path, n_pages, from_end)

    if Document is None:
        return [], "python-docx chưa được cài (pip install python-docx)"
    try:
        doc = Document(docx_path)
    except Exception as e:
        return [], f"Không mở được file Word: {e}"

    n_pages = max(int(n_pages), 1)
    pages = []
    current = []
    read_chars = 0
    char_cap = n_pages * MAX_CHARS_PER_PAGE_READ

    def _flush():
        if any(current):
            pages.append(list(current))
        current.clear()

    def _add(lines):
        nonlocal read_chars
        for ln in lines:
            if ln:
                current.append(ln)
                read_chars += len(ln)

    for kind, el in _iter_blocks(doc.element.body):
        if not from_end and (len(pages) >= n_pages or read_chars >= char_cap):
            break
        if kind == "tbl":
            _add(_table_lines(el, doc))
            continue
        mode = _page_break_mode(el)
        if mode == "before":
            _flush()
            if not from_end and len(pages) >= n_pages:
                break
        para = Paragraph(el, doc)
        _add([_para_text(para)] + [_para_text(tb) for tb in _textbox_paragraphs(el, doc)])
        if mode == "after" or _ends_with_page_section(el):
            _flush()
    _flush()

    page_texts = ["\n".join(pg) for pg in pages]
    page_texts = page_texts[-n_pages:] if from_end else page_texts[:n_pages]

    if n_pages > 1 and len(page_texts) == 1 and len(page_texts[0]) > VIRTUAL_PAGE_CHARS:
        page_texts = _split_virtual_pages(page_texts[0], n_pages, from_end)

    f = chon_bang_ma("\n".join(page_texts))  # phông cũ TCVN3 / VNI: chọn theo cả đoạn, đổi từng trang
    if f:
        page_texts = [f(t) for t in page_texts]
    return page_texts, ""


# Đọc file .doc bằng công cụ ngoài
def _extract_doc_text_as_pages(doc_path: str, n_pages: int, from_end: bool = False) -> tuple:
    text = ""
    for cmd in (["antiword", "-m", "UTF-8.txt", doc_path], ["catdoc", "-w", "-d", "utf-8", doc_path]):
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                               errors="replace", timeout=30)
            if r.returncode == 0 and r.stdout.strip():
                text = r.stdout
                break
        except (OSError, subprocess.SubprocessError):
            pass

    if not text.strip():
        try:
            with tempfile.TemporaryDirectory() as tmp:
                subprocess.run(
                    ["soffice", "--headless", "--convert-to", "txt:Text (encoded):UTF8",
                     "--outdir", tmp, doc_path],
                    capture_output=True, text=True, timeout=60,
                )
                out_txt = os.path.join(tmp, os.path.splitext(os.path.basename(doc_path))[0] + ".txt")
                if os.path.exists(out_txt):
                    with open(out_txt, encoding="utf-8", errors="replace") as f:
                        text = f.read()
        except (OSError, subprocess.SubprocessError):
            pass

    if not text.strip():
        return [], ("Không đọc được file .doc: server cần antiword, catdoc hoặc libreoffice (soffice), "
                    "hoặc hãy chuyển file sang .docx")

    pages = _split_virtual_pages(fix_tcvn3_text(text.strip()), n_pages, from_end)
    if not pages:
        return [], "File .doc không có text"
    return pages, ""


# Kiểm tra text Word có dùng được không
def word_quality_problem(page_texts: list) -> str:
    text = "\n".join(page_texts)
    visible = sum(not c.isspace() for c in text)
    if not visible:
        return "Word không có chữ"
    letters = sum(c.isalpha() for c in text)
    if letters / visible < 0.5:
        return f"chỉ {letters / visible:.0%} ký tự không-trắng là chữ cái → nghi là rác / lỗi mã hoá"
    la, diem = sum(c in _KY_TU_MA_CU for c in text), vn_score(text)
    if la >= 3 and letters and la / letters >= 0.02 and diem < VN_MIN_SCORE:
        return (f"còn {la} ký tự lạ, chỉ {diem:.0%} âm tiết tiếng Việt hợp lệ → nghi lỗi bảng mã "
                f"(phông cũ TCVN3 / VNI / VPS... chưa đổi được)")
    return ""


# Đọc N trang đầu/cuối file Word
def read_word_pages(word_path: str, n_pages: int, job_id: str = "", from_end: bool = False) -> tuple[list[str], str]:
    log = get_logger("input", job_id or None)
    t0 = time.time()
    try:
        pages, error = _extract_word_pages(word_path, n_pages, from_end)
    except Exception as e:
        pages, error = [], f"{type(e).__name__}: {e}"
    problem = f"đọc Word lỗi: {error}" if error else word_quality_problem(pages)
    log.log(logging.WARNING if problem else logging.INFO,
            f"Word {os.path.basename(word_path)}: {len(pages)} trang {'cuối' if from_end else 'đầu'}, "
            f"{sum(len(p) for p in pages)} ký tự, "
            f"{time.time() - t0:.2f}s" + (f" → không dùng được: {problem}" if problem else ""))
    return pages, problem


# Chandra OCR

# Đổi ảnh sang base64
def _img_to_b64(img: np.ndarray, max_side: int | None = MAX_LONG_SIDE) -> str:
    h, w = img.shape[:2]
    if max_side and max(h, w) > max_side:
        scale = max_side / max(h, w)
        img = cv2.resize(img, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_AREA)
    _, buf = cv2.imencode(".jpg", img, [cv2.IMWRITE_JPEG_QUALITY, 90])
    return base64.b64encode(buf.tobytes()).decode("utf-8")


# Đổi HTML của Chandra sang text
def strip_html(text: str) -> str:
    text = re.sub(r'<div[^>]*data-label="(?:Figure|Image)"[^>]*>.*?(?:</div>|$)', " ", text,
                  flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<img[^>]*/?>", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"<su[pb][^>]*>(.*?)</su[pb]>", r"\1", text, flags=re.IGNORECASE)
    text = re.sub(r"</t[dh]>\s*<t[dh][^>]*>", " | ", text, flags=re.IGNORECASE)
    text = re.sub(r"</?t[dh][^>]*>", "", text, flags=re.IGNORECASE)
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</?(?:p|div|h[1-6]|li|tr)[^>]*>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"[*_#`]{1,3}", "", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return unicodedata.normalize("NFC", text.replace("\xa0", " ")).strip()


_QUOTED_RES = (
    re.compile(r"(?<![A-Za-zÀ-ỹ0-9])'([^'\n]{2,200}?)'(?![A-Za-zÀ-ỹ0-9])"),
    re.compile(r"‘([^’\n]{2,200}?)’"),
    re.compile(r"“([^”\n]{2,200}?)”"),
)


# Lấy các chuỗi trong JSON
def _json_strings(obj) -> list[str]:
    if isinstance(obj, str):
        return [obj]
    if isinstance(obj, dict):
        return [x for v in obj.values() for x in _json_strings(v)]
    if isinstance(obj, list):
        return [x for v in obj for x in _json_strings(v)]
    return []


# Lấy chữ khi Chandra trả JSON
def strip_chandra_json(text: str) -> str | None:
    stripped = text.strip()
    if not stripped.startswith(("[{", "{", "[[")):
        return None
    strings = []
    for line in stripped.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            strings.extend(_json_strings(json.loads(line)))
        except json.JSONDecodeError:
            strings.extend(re.findall(r'"text"\s*:\s*"((?:[^"\\]|\\.)*)"', line))
            strings.extend(m for rx in _QUOTED_RES for m in rx.findall(line))
    phrases = []
    for value in strings:
        found = [m for rx in _QUOTED_RES for m in rx.findall(value)]
        if not found and len(value) <= 200 and not re.search(r"\b(?:the|is|of|and|with|in)\b", value, re.I):
            found = [value]
        phrases.extend(f.strip() for f in found if re.search(r"\w", f))
    seen, out = set(), []
    for ph in phrases:
        if ph.lower() not in seen:
            seen.add(ph.lower())
            out.append(ph)
    if not out:
        return None
    return unicodedata.normalize("NFC", "\n".join(out))


# OCR một ảnh bằng Chandra
def chandra_ocr(img: np.ndarray, max_tokens: int = CHANDRA_MAX_TOKENS, label: str = "",
                category: str = "first_8_pages", job_id: str = "") -> str:
    log = get_logger(category, job_id or None)
    payload = {
        "model": CHANDRA_MODEL,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{_img_to_b64(img)}"}},
                {"type": "text", "text": "<image>"},
            ],
        }],
        "max_tokens": max_tokens,
        "temperature": 0.0,
    }
    t0 = time.time()
    resp = _post("Chandra", CHANDRA_URL, payload, CHANDRA_TIMEOUT, log, job_id)
    raw = resp.json()["choices"][0]["message"]["content"].strip()
    text = strip_chandra_json(raw)
    if text is None:
        text = strip_html(raw)
    else:
        log.warning(f"Chandra {label} trả JSON mô tả ảnh thay vì chữ ({len(raw)} ký tự) → lấy cụm trong nháy: "
                    f"{text.replace(chr(10), ' | ')[:200]!r}")
    log.info(f"Chandra {label}: {len(text)} ký tự, {time.time() - t0:.2f}s")
    return text


# Sửa chính tả

# Sửa chính tả, lỗi thì trả nguyên bản
def spellfix(texts: list[str], job_id: str = "") -> list[str]:
    log = get_logger("spellfix", job_id or None)
    if not SPELLFIX_ENABLED:
        log.info("Spellfix đang tắt (SPELLFIX_ENABLED=0)")
        return texts
    if not any(t.strip() for t in texts):
        return texts
    t0 = time.time()
    try:
        resp = _post("Spellfix", SPELLFIX_API_URL, {"texts": texts}, 300, log, job_id)
        fixed = resp.json().get("texts", [])
        if len(fixed) != len(texts):
            log.warning(f"Spellfix trả {len(fixed)} đoạn, cần {len(texts)} → giữ nguyên bản")
            return texts
        for i, (before, after) in enumerate(zip(texts, fixed)):
            if before == after:
                continue
            if re.findall(r"\d+", before) != re.findall(r"\d+", after):
                log.warning(f"Bỏ kết quả sửa vì chữ số bị đổi: {before[:200]!r} → {after[:200]!r}")
                fixed[i] = before
            else:
                log.info(f"Sửa: {before[:200]!r} → {after[:200]!r}")
        log.info(f"Spellfix {len(texts)} đoạn, {sum(a != b for a, b in zip(texts, fixed))} đoạn thay đổi, "
                 f"{time.time() - t0:.2f}s")
        return fixed
    except Exception as e:
        log.warning(f"Spellfix lỗi, giữ nguyên bản: {e}")
        return texts


# Cấu hình trang cuối: Chandra OCR cả trang

CHANDRA_DATE_MAX_TOKENS = int(os.getenv("SORT_CHANDRA_DATE_MAX_TOKENS", "3000"))  # đủ cho cả trang, dòng In xong thường ở cuối
DATE_LAST_PAGES         = int(os.getenv("SORT_DATE_LAST_PAGES", "5"))  # tối đa 5 trang cuối, xét lần lượt từ trang cuối lên


# Số trang cần xét: trang cuối cùng đứng đầu
def last_page_numbers(page_count: int, n_pages: int) -> list[int]:
    return list(range(page_count, max(0, page_count - max(n_pages, 1)), -1))


# Render một trang PDF (chỉ render khi cần, dừng sớm thì không tốn render các trang còn lại)
def render_page(pdf_path: str, out_dir: str, no: int, job_id: str = "") -> str:
    os.makedirs(out_dir, exist_ok=True)
    t0 = time.time()
    path = os.path.join(out_dir, f"last_page_{no}.jpg")
    convert_from_path(pdf_path, dpi=PDF_DPI, first_page=no, last_page=no)[0].convert("RGB").save(path, "JPEG", quality=95)
    get_logger("last_page", job_id or None).info(f"Render trang {no}, {PDF_DPI} dpi, {time.time() - t0:.2f}s")
    return path


# Tách tháng, năm lưu chiểu

_LUU_CHIEU_KEY_RE = re.compile(r"in xong|l\w{0,2}u chi\w{0,3}u")  # cả lỗi font / OCR: 'lou chieau', 'luu chieu'
# Tháng / năm sau chữ In xong (text đã bỏ dấu): 'tháng 3-2021', 'tháng 3 – 2021', 'tháng 03/2021', 'tháng 3,2021',
# 'T3/2021', 'tháng 3-21', 'tháng ba năm 2021', 'tháng năm năm 2021' (= 5/2021), hoặc '3/2021' đứng ngay sau chữ In xong
_GACH = r"\s*(?:nam|[/.,\-\u2013\u2014])\s*"
_THANG_NAM_RE = re.compile(r"(?:\bthang|\bt)\s*\.?\s*(\d{1,2})" + _GACH + r"(\d{4}|\d{2})(?!\d)")
_THANG_CHU = {"mot": 1, "gieng": 1, "hai": 2, "ba": 3, "tu": 4, "bon": 4, "nam": 5, "sau": 6, "bay": 7, "tam": 8,
              "chin": 9, "muoi": 10, "muoi mot": 11, "muoi hai": 12, "chap": 12}
_THANG_CHU_RE = re.compile(r"\bthang\s+(muoi mot|muoi hai|muoi|gieng|chap|mot|hai|ba|tu|bon|sau|bay|tam|chin|nam(?=\s+nam\b))"
                           r"\s*(?:nam\b)?\s*(\d{4})(?!\d)")
_SO_DAU_RE = re.compile(r"\D*?(?<!\d)(\d{1,2})\s*[/.\-\u2013\u2014]\s*(\d{4})(?!\d)")  # số đầu tiên sau chữ In xong


# Năm 2 chữ số -> 4 chữ số ('21' -> 2021, '98' -> 1998)
def _nam4(y: str) -> int:
    y = int(y)
    if y >= 100:
        return y
    return 2000 + y if 2000 + y <= time.localtime().tm_year + 1 else 1900 + y


# Các cặp (tháng, năm) hợp lệ trong đoạn từ chữ In xong trở đi, theo thứ tự xuất hiện
def tim_thang_nam(doan: str) -> list:
    kq = [(m.start(), int(m.group(1)), _nam4(m.group(2))) for m in _THANG_NAM_RE.finditer(doan)]
    kq += [(m.start(), _THANG_CHU[m.group(1)], int(m.group(2))) for m in _THANG_CHU_RE.finditer(doan)]
    if not kq:
        m = _SO_DAU_RE.match(doan)
        if m:
            kq.append((m.start(1), int(m.group(1)), int(m.group(2))))
    return [(t, y) for _, t, y in sorted(kq) if 1 <= t <= 12 and 1900 <= y <= 2100]


# Dòng Quyết định xuất bản (text đã bỏ dấu): 'Quyết định xuất bản số 79-QĐ/NXBCTQG ngày 05/3/2021',
# 'Số QĐXB: 12/QĐ-NXBTP ngày 05 tháng 3 năm 2021', 'Số 79-QĐ/NXBCTQG ngày 05/3/2021' (nhãn ở dòng trên)
_QDXB_RE = re.compile(r"quyet dinh xuat ban|\bqdxb\b|\bqd\s*[-/]\s*nxb|\bnxb\w*\s*[-/]\s*qd\b")
_NGAY_QD_RE = re.compile(r"\bngay\s*:?\s*(\d{1,2})\s*(?:[/.\-\u2013]|thang)\s*(\d{1,2})\s*(?:[/.\-\u2013]|nam)\s*(\d{4}|\d{2})(?!\d)")


# Ngày Quyết định xuất bản -> (tháng, năm, câu) hoặc None; dòng QĐXB chưa có ngày thì nối dòng sau, lấy dòng cuối cùng
def ngay_qdxb(lines: list, norms: list):
    kq = None
    for i, n in enumerate(norms):
        m = _QDXB_RE.search(n)
        if not m:
            continue
        doan, cau = n[m.start():], lines[i]
        if not _NGAY_QD_RE.search(doan) and i + 1 < len(norms) and _NGAY_QD_RE.search(norms[i + 1]):
            doan, cau = f"{doan} {norms[i + 1]}", f"{lines[i]} {lines[i + 1]}"
        d = _NGAY_QD_RE.search(doan)
        if d and 1 <= int(d.group(1)) <= 31 and 1 <= int(d.group(2)) <= 12 and 1900 <= _nam4(d.group(3)) <= 2100:
            kq = (int(d.group(2)), _nam4(d.group(3)), cau)
    return kq


_NAM_RE = re.compile(r"(?<!\d)((?:19|20)\d{2})(?!\d)")
_COLOPHON_RE = re.compile(r"^(?:in\b|so dang ky|so xac nhan|quyet dinh|ma so|isbn|dia chi|kho\b|in xong|nop luu chieu)")


# Bỏ dấu tiếng Việt
def _bo_dau(s: str) -> str:
    s = s.replace("đ", "d").replace("Đ", "D")
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").lower()


# Dòng In xong chưa có ngày thì nối dòng sau (OCR / Word hay xuống dòng giữa câu)
def _noi_dong_sau(norms: list, i: int) -> bool:
    return (not tim_thang_nam(norms[i]) and not _NAM_RE.search(norms[i]) and i + 1 < len(norms)
            and not _LUU_CHIEU_KEY_RE.search(norms[i + 1]) and bool(_NAM_RE.search(norms[i + 1])))


# Tách tháng, năm: từ dòng có 'in xong / lưu chiểu' (chỉ xét phần từ chữ đó trở đi); dòng In xong không ra
# tháng + năm thì lấy ngày Quyết định xuất bản; vẫn không có thì năm của dòng In xong. theo: in_xong / qdxb
def extract_luu_chieu(text: str) -> dict:
    lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    norms = [_bo_dau(ln) for ln in lines]
    du, chi_nam = None, None
    for i, n in enumerate(norms):
        m = _LUU_CHIEU_KEY_RE.search(n)
        if not m:
            continue
        doan, cau = n[m.start():], lines[i]
        if _noi_dong_sau(norms, i):
            doan, cau = f"{doan} {norms[i + 1]}", f"{lines[i]} {lines[i + 1]}"
        hop_le = tim_thang_nam(doan)
        if hop_le:
            du = (*hop_le[-1], cau)
            continue
        years = [int(y) for y in _NAM_RE.findall(doan)]
        if years:
            chi_nam = (years[-1], cau)
    if du:
        return {"thang": du[0], "nam": du[1], "cau": du[2], "theo": "in_xong"}
    qd = ngay_qdxb(lines, norms)
    if qd:
        return {"thang": qd[0], "nam": qd[1], "cau": qd[2], "theo": "qdxb"}
    if chi_nam:
        return {"thang": None, "nam": chi_nam[0], "cau": chi_nam[1], "theo": "in_xong"}
    return {"thang": None, "nam": None, "cau": "", "theo": ""}


# Lấy khối thông tin xuất bản: quanh dòng In xong cuối cùng và dòng Quyết định xuất bản cuối cùng,
# kèm dòng sau nếu ngày bị xuống dòng
def colophon_block(text: str) -> str:
    lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    norms = [_bo_dau(ln) for ln in lines]
    moc = []
    idx = [i for i, n in enumerate(norms) if _LUU_CHIEU_KEY_RE.search(n)]
    if idx:
        moc += [idx[-1], idx[-1] + (1 if _noi_dong_sau(norms, idx[-1]) else 0)]
    qd = [i for i, n in enumerate(norms) if _QDXB_RE.search(n)]
    if qd:
        moc += [qd[-1], qd[-1] + (1 if qd[-1] + 1 < len(norms) and not _NGAY_QD_RE.search(norms[qd[-1]])
                                  and _NGAY_QD_RE.search(norms[qd[-1] + 1]) else 0)]
    if not moc:
        return ""
    start, end = min(moc), max(moc)
    while start > 0 and end - start < 12 and _COLOPHON_RE.search(norms[start - 1]):
        start -= 1
    return "\n".join(lines[start:end + 1])


# Có tháng hoặc năm chưa
def _co_ngay(luu: dict) -> bool:
    return luu["thang"] is not None or luu["nam"] is not None


# Các bước của luồng

EMPTY_RESULT = {"metadata": "", "metadata_raw": "", "metadata_score": None, "metadata_box": None,
                "metadata_page": None, "metadata_error": "", "luu_chieu_thang": None, "luu_chieu_nam": None,
                "luu_chieu_cau": "", "luu_chieu_source": "", "luu_chieu_theo": "", "date_word_reason": ""}


# Bước 1: đọc Word trang cuối
def date_prepare(context: dict) -> None:
    job_id = context.get("job_id", "")
    pdf_path, word_path = context.get("pdf_path", ""), context.get("docx_path", "")
    if pdf_path and not context.get("page_count"):
        context["page_count"] = pdf_page_count(pdf_path)
    log = get_logger("last_page", job_id or None)
    context.update(metadata={"text": "", "score": None, "box": None, "page": None}, last_page_paths=[], last_pages=[],
                   date_source="", date_word_reason="")
    word_text = ""
    if word_path:
        pages, problem = read_word_pages(word_path, DATE_LAST_PAGES, job_id, from_end=True)
        word_text = "" if problem else "\n".join(pages)
        if word_text:
            block = colophon_block(word_text)
            if block and _co_ngay(extract_luu_chieu(block)):
                context["metadata"].update(text=block)
                context["date_source"] = "word"
                log.info(f"Lấy ngày từ Word {os.path.basename(word_path)}: {block.splitlines()[-1]!r}")
                return
            problem = f"không thấy dòng 'In xong ...' / Quyết định xuất bản có ngày trong {DATE_LAST_PAGES} trang cuối Word"
        context["date_word_reason"] = problem
        log.warning(f"Word không cho ra ngày: {problem}" + (" → OCR trang cuối PDF" if pdf_path else ""))
    if not pdf_path:
        if word_path:
            context["date_source"] = "word"
            return
        raise ValueError("Thiếu file PDF")
    context["last_pages"] = last_page_numbers(context["page_count"], DATE_LAST_PAGES)
    context["date_source"] = "ocr"


# Bước 2: Chandra OCR cả trang, lần lượt từ trang cuối lên (tối đa DATE_LAST_PAGES trang),
# dừng ở trang có dòng 'In xong và nộp lưu chiểu ...' ra ngày; METADATA = khối thông tin xuất bản quanh dòng đó
def date_ocr(context: dict) -> None:
    if context.get("date_source") != "ocr":
        return
    job_id = context.get("job_id", "")
    log = get_logger("last_page", job_id or None)
    page_dir = os.path.join(context["job_folder"], "pages")
    metadata = {"text": "", "score": None, "box": None, "page": None}
    pages = context.get("last_pages", [])
    tam = None
    try:
        for no in pages:
            path = render_page(context["pdf_path"], page_dir, no, job_id)
            context.setdefault("last_page_paths", []).append((no, path))
            text = chandra_ocr(cv2.imread(path), max_tokens=CHANDRA_DATE_MAX_TOKENS,
                               label=f"trang {no}", category="last_page", job_id=job_id)
            if not metadata["text"]:
                metadata.update(text=text, page=no)  # chưa trang nào ra ngày: giữ text trang cuối cùng để xem lại
            r = extract_luu_chieu(text)
            if r["thang"] is not None:  # In xong ra tháng + năm, hoặc ngày Quyết định xuất bản
                metadata.update(text=colophon_block(text) or text, page=no)
                break
            if r["nam"] is not None and tam is None:
                tam = (colophon_block(text) or text, no)  # In xong chỉ có năm: giữ tạm, xem trang trước có tháng không
            log.info(f"Trang {no} không ra tháng / năm (In xong, Quyết định xuất bản) → xét trang trước")
        else:
            if tam:
                metadata.update(text=tam[0], page=tam[1])
            else:
                log.warning(f"Không thấy dòng 'In xong ...' / Quyết định xuất bản có ngày ở {len(pages)} trang cuối")
    except Exception as e:
        log.error(f"METADATA lỗi: {e}", exc_info=True)
        metadata["error"] = str(e)
    context["metadata"] = metadata


# Bước 3: trả kết quả tháng, năm
def date_result(context: dict) -> dict:
    job_id = context.get("job_id", "")
    metadata = context.get("metadata", {})
    metadata_raw = metadata.get("text", "")
    source = context.get("date_source", "")
    metadata_text = spellfix([metadata_raw], job_id)[0] if source == "ocr" else metadata_raw
    luu_chieu = extract_luu_chieu(metadata_raw)
    if not _co_ngay(luu_chieu):
        luu_chieu = extract_luu_chieu(metadata_text)
    if metadata_raw or source:
        get_logger("last_page", job_id or None).log(
            logging.INFO if _co_ngay(luu_chieu) else logging.WARNING,
            f"Lưu chiểu ({source or '-'}, theo {luu_chieu['theo'] or '-'}): tháng={luu_chieu['thang']} năm={luu_chieu['nam']} "
            f"← {luu_chieu['cau'] or 'không tìm thấy'!r}")
    return {
        "metadata": metadata_text,
        "metadata_raw": metadata_raw,
        "metadata_score": metadata.get("score"),
        "metadata_box": metadata.get("box"),
        "metadata_page": metadata.get("page"),
        "metadata_error": metadata.get("error", ""),
        "luu_chieu_thang": luu_chieu["thang"],
        "luu_chieu_nam": luu_chieu["nam"],
        "luu_chieu_cau": luu_chieu["cau"],
        "luu_chieu_source": source,
        "luu_chieu_theo": luu_chieu["theo"],  # in_xong / qdxb
        "date_word_reason": context.get("date_word_reason", ""),
    }
