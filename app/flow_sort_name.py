# Luồng lấy tên sách ở các trang đầu
import os
import re
import json
import time
import base64
import logging
import tempfile
import subprocess
import unicodedata
from concurrent.futures import ThreadPoolExecutor

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


# Cấu hình Qwen

FIRST_PAGES     = int(os.getenv("SORT_FIRST_PAGES", "8"))
QWEN_URL        = os.getenv("QWEN_LLM_URL",   "http://localhost:8113/v1/chat/completions")
QWEN_MODEL      = os.getenv("QWEN_LLM_MODEL", "Qwen/Qwen3-4B-Instruct-2507")
QWEN_MAX_CHARS  = int(os.getenv("SORT_QWEN_MAX_CHARS", "8000"))
QWEN_PAGE_CHARS = int(os.getenv("SORT_QWEN_PAGE_CHARS", "900"))
QWEN_MAX_TOKENS = int(os.getenv("SORT_QWEN_MAX_TOKENS", "300"))
QWEN_TIMEOUT    = int(os.getenv("SORT_QWEN_TIMEOUT", "300"))
BIA_MAX_PAGES   = int(os.getenv("SORT_BIA_MAX_PAGES", "2"))  # số trang đầu của PDF bìa đem OCR


# Đọc các trang đầu

# Render các trang đầu PDF
def render_first_pages(pdf_path: str, out_dir: str, page_count: int, job_id: str = "") -> list[str]:
    os.makedirs(out_dir, exist_ok=True)
    t0 = time.time()
    paths = []
    n = min(FIRST_PAGES, page_count)
    for i, pil in enumerate(convert_from_path(pdf_path, dpi=PDF_DPI, first_page=1, last_page=n)):
        path = os.path.join(out_dir, f"page_{i + 1}.jpg")
        pil.convert("RGB").save(path, "JPEG", quality=90)
        paths.append(path)
    get_logger("first_8_pages", job_id or None).info(
        f"Render {len(paths)}/{page_count} trang đầu, {PDF_DPI} dpi, {time.time() - t0:.2f}s")
    return paths


# OCR song song các trang đầu
def ocr_first_pages(image_paths: list[str], job_id: str = "") -> list[str]:
    log = get_logger("first_8_pages", job_id or None)

    def _one(idx_path):
        idx, path = idx_path
        try:
            return chandra_ocr(cv2.imread(path), label=f"trang {idx + 1}", job_id=job_id)
        except Exception as e:
            log.error(f"Chandra trang {idx + 1} lỗi: {e}")
            return ""

    t0 = time.time()
    with ThreadPoolExecutor(max_workers=max(1, min(CHANDRA_WORKERS, len(image_paths)))) as ex:
        pages = list(ex.map(_one, enumerate(image_paths)))
    failed = [i + 1 for i, p in enumerate(pages) if not p]
    log.info(f"OCR {len(pages)} trang xong {time.time() - t0:.2f}s, "
             f"{sum(len(p) for p in pages)} ký tự, trang rỗng/lỗi: {failed or '-'}")
    if image_paths and not any(pages):
        raise RuntimeError("Chandra OCR thất bại trên toàn bộ các trang đầu")
    return pages


# Ghép text các trang
def join_pages(pages: list[str], page_chars: int = 0) -> str:
    parts = []
    for i, p in enumerate(pages):
        if not p:
            continue
        if page_chars and len(p) > page_chars:
            p = p[:page_chars].rstrip() + " …"
        parts.append(f"[TRANG {i + 1}]\n{p}")
    return "\n\n".join(parts)


# Tạo text gửi Qwen
def build_qwen_input(context: dict) -> str:
    pages = context.get("source_pages")
    if pages:
        return join_pages(pages, 0 if context.get("text_source") == "word" else QWEN_PAGE_CHARS)
    return context.get("source_text", "")


# Qwen tìm tên sách

_QWEN_TITLE_PROMPT = """Bạn là hệ thống nhận diện tên sách từ văn bản của các trang đầu một cuốn sách
(bìa, trang lót, trang tên sách, lời nói đầu, mục lục...). Văn bản có thể có lỗi OCR.

Nhiệm vụ: tìm TÊN CUỐN SÁCH (tên chính của ấn phẩm).
- Ưu tiên tên in trên bìa / trang tên sách, thường là dòng chữ lớn, viết hoa, lặp lại nhiều lần.
- KHÔNG lấy tên tác giả, nhà xuất bản, tên bộ sách/tủ sách, tên chương, câu khẩu hiệu.
- Giữ nguyên câu chữ của tên sách, chỉ sửa lỗi OCR hiển nhiên (dấu, ký tự dính). Không tự bịa.
- Nếu tên sách viết hoa toàn bộ thì giữ nguyên dạng viết hoa.
- Không tìm thấy thì trả chuỗi rỗng.

Chỉ trả về JSON đúng dạng: {"ten_sach": "..."}"""


# Đọc JSON Qwen trả về
def _parse_json(raw: str, log) -> dict | None:
    raw = re.sub(r"<think>.*?</think>", "", raw, flags=re.DOTALL).strip()
    raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw).strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        from json_repair import repair_json
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        try:
            return json.loads(repair_json(m.group(0) if m else raw))
        except Exception:
            log.warning(f"Qwen JSON parse lỗi. raw={raw[:200]!r}")
            return None


_TOO_LONG_RE = re.compile(r"maximum context length|max_model_len|context length|too long|exceed", re.I)


# Cắt text theo giới hạn ký tự
def qwen_cut(text: str) -> str:
    return text if QWEN_MAX_CHARS <= 0 else text[:QWEN_MAX_CHARS]


# Gọi Qwen tìm tên sách
def qwen_find_title(text: str, job_id: str = "") -> str:
    log = get_logger("qwen", job_id or None)
    body = qwen_cut(text)
    t0 = time.time()
    for attempt in range(8):
        payload = {
            "model": QWEN_MODEL,
            "messages": [
                {"role": "system", "content": _QWEN_TITLE_PROMPT},
                {"role": "user",   "content": body + "\n/no_think"},
            ],
            "temperature": 0.0,
            "max_tokens": QWEN_MAX_TOKENS,
            "response_format": {"type": "json_object"},
            "chat_template_kwargs": {"enable_thinking": False},
        }
        log.info(f"Qwen → gửi {len(body)}/{len(text)} ký tự")
        try:
            resp = _post("Qwen", QWEN_URL, payload, QWEN_TIMEOUT, log, job_id)
            break
        except requests.HTTPError as e:
            r = e.response
            too_long = r is not None and r.status_code == 400 and _TOO_LONG_RE.search(r.text or "")
            if too_long and attempt < 7 and len(body) > 2000:
                body = body[:int(len(body) * 0.7)]
                log.warning(f"Qwen báo text quá dài so với max-model-len → cắt còn {len(body)} ký tự rồi gửi lại")
                continue
            raise
    raw = resp.json()["choices"][0]["message"]["content"].strip()
    log.info(f"Qwen ← {time.time() - t0:.2f}s raw={raw[:200]!r}")
    fields = _parse_json(raw, log) or {}
    title = fields.get("ten_sach", "")
    title = re.sub(r"\s+", " ", str(title)).strip() if title else ""
    if not title:
        log.warning("Qwen không tìm thấy tên sách")
    return title


# Các bước của luồng

EMPTY_RESULT = {"ten_sach": "", "ten_sach_raw": "", "text_source": "", "word_fallback_reason": ""}


# OCR các trang đầu PDF
def _ocr_pdf_pages(context: dict) -> None:
    job_id = context.get("job_id", "")
    paths = render_first_pages(context["pdf_path"], os.path.join(context["job_folder"], "pages"),
                               context["page_count"], job_id)
    pages = ocr_first_pages(paths, job_id)
    context.update(first_page_paths=paths, source_pages=pages, source_text=join_pages(pages), text_source="ocr")


# OCR file bìa (ảnh hoặc PDF bìa): dùng khi Word / PDF ruột không ra tên sách
def _ocr_bia(context: dict) -> None:
    job_id = context.get("job_id", "")
    log = get_logger("first_8_pages", job_id or None)
    out_dir = os.path.join(context["job_folder"], "bia")
    os.makedirs(out_dir, exist_ok=True)
    paths = []
    for p in context.get("bia_paths", []):
        if p.lower().endswith(".pdf"):
            goc = os.path.splitext(os.path.basename(p))[0]
            for i, pil in enumerate(convert_from_path(p, dpi=PDF_DPI, first_page=1, last_page=BIA_MAX_PAGES)):
                dst = os.path.join(out_dir, f"{goc}_{i + 1}.jpg")
                pil.convert("RGB").save(dst, "JPEG", quality=90)
                paths.append(dst)
        elif cv2.imread(p) is not None:
            paths.append(p)
        else:
            log.warning(f"Không đọc được ảnh bìa {os.path.basename(p)}")
    pages = ocr_first_pages(paths, job_id) if paths else []
    log.info(f"OCR bìa: {len(paths)} ảnh, {sum(len(x) for x in pages)} ký tự")
    context.update(first_page_paths=paths, source_pages=pages, source_text=join_pages(pages), text_source="bia")


# Bước 1: đọc Word hoặc render PDF
def name_prepare(context: dict) -> None:
    job_id = context.get("job_id", "")
    pdf_path, word_path = context.get("pdf_path", ""), context.get("docx_path", "")
    if pdf_path and not context.get("page_count"):
        context["page_count"] = pdf_page_count(pdf_path)
    word_pages, problem = [], ""
    if word_path:
        word_pages, problem = read_word_pages(word_path, FIRST_PAGES, job_id)
        # Lỗi bảng mã: text Word là rác, gửi Qwen vô ích kể cả khi không có PDF
        if problem and (pdf_path or "bảng mã" in problem):
            word_pages = []
        elif problem and word_pages:
            get_logger("input", job_id or None).warning("Không có PDF để OCR → dùng tạm text Word chất lượng kém")
    context["word_fallback_reason"] = problem
    if word_pages:
        context.update(source_pages=word_pages, source_text=join_pages(word_pages), text_source="word",
                       first_page_paths=[])
        return
    if not pdf_path:
        if context.get("bia_paths"):  # không có Word / PDF dùng được: để bước 3 OCR bìa
            context.update(source_pages=[], source_text="", text_source="", first_page_paths=[])
            return
        raise ValueError(f"Không dùng được file Word ({problem}) và không có PDF để OCR" if word_path
                         else "Thiếu file PDF")
    first_paths = render_first_pages(pdf_path, os.path.join(context["job_folder"], "pages"),
                                     context["page_count"], job_id)
    context.update(source_pages=[], source_text="", text_source="ocr", first_page_paths=first_paths)


# Bước 2: OCR các trang đầu
def name_ocr(context: dict) -> None:
    if context.get("text_source") == "ocr":
        pages = ocr_first_pages(context.get("first_page_paths", []), context.get("job_id", ""))
        context["source_pages"] = pages
        context["source_text"] = join_pages(pages)


# Bước 3: Qwen tìm tên sách
def name_result(context: dict) -> dict:
    job_id = context.get("job_id", "")
    qwen_input = build_qwen_input(context)
    title_raw = qwen_find_title(qwen_input, job_id) if qwen_input.strip() else ""
    if not title_raw and context.get("text_source") == "word" and context.get("pdf_path"):
        context["word_fallback_reason"] = "Qwen không tìm được tên sách từ text Word"
        get_logger("first_8_pages", job_id or None).warning(
            f"{context['word_fallback_reason']} → OCR {FIRST_PAGES} trang đầu PDF rồi tìm lại")
        _ocr_pdf_pages(context)
        qwen_input = build_qwen_input(context)
        title_raw = qwen_find_title(qwen_input, job_id) if qwen_input.strip() else ""
    if not title_raw and context.get("bia_paths"):
        log = get_logger("first_8_pages", job_id or None)
        log.warning("Word / PDF không ra tên sách → OCR file bìa rồi tìm lại")
        try:
            _ocr_bia(context)
            qwen_input = build_qwen_input(context)
            title_raw = qwen_find_title(qwen_input, job_id) if qwen_input.strip() else ""
        except Exception as e:
            log.error(f"OCR bìa lỗi: {e}")
    title = spellfix([title_raw], job_id)[0] if context.get("text_source") in ("ocr", "bia") else title_raw
    return {"ten_sach": title, "ten_sach_raw": title_raw, "text_source": context.get("text_source", ""),
            "word_fallback_reason": context.get("word_fallback_reason", "")}
