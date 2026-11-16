# Luồng lấy thông tin xuất bản ở trang cuối (Word trước, không có thì Chandra OCR cả trang PDF)
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


# Cấu hình Chandra, PDF

CHANDRA_URL        = os.getenv("CHANDRA_URL",   "http://localhost:8101/v1/chat/completions")
CHANDRA_MODEL      = os.getenv("CHANDRA_MODEL", "datalab-to/chandra-ocr-2")
CHANDRA_MAX_TOKENS = int(os.getenv("SORT_CHANDRA_MAX_TOKENS", "2400"))
CHANDRA_WORKERS    = int(os.getenv("SORT_CHANDRA_WORKERS", "4"))
CHANDRA_TIMEOUT    = int(os.getenv("SORT_CHANDRA_TIMEOUT", "600"))
MAX_LONG_SIDE      = int(os.getenv("SORT_CHANDRA_MAX_SIDE", "1800"))


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


# Cấu hình trang cuối

INFO_LAST_PAGES         = int(os.getenv("SORT_INFO_LAST_PAGES", "2"))
CHANDRA_INFO_MAX_TOKENS = int(os.getenv("SORT_CHANDRA_INFO_MAX_TOKENS", "3000"))
INFO_MIN_FIELDS         = int(os.getenv("SORT_INFO_MIN_FIELDS", "3"))


# Render các trang cuối PDF, trang cuối cùng đứng đầu
def render_last_pages(pdf_path: str, out_dir: str, page_count: int, n_pages: int, job_id: str = "") -> list[tuple]:
    os.makedirs(out_dir, exist_ok=True)
    t0 = time.time()
    first = max(1, page_count - max(n_pages, 1) + 1)
    pages = []
    for no, pil in zip(range(first, page_count + 1),
                       convert_from_path(pdf_path, dpi=PDF_DPI, first_page=first, last_page=page_count)):
        path = os.path.join(out_dir, f"info_page_{no}.jpg")
        pil.convert("RGB").save(path, "JPEG", quality=95)
        pages.append((no, path))
    get_logger("metadata", job_id or None).info(
        f"Render trang {first}-{page_count}/{page_count}, {PDF_DPI} dpi, {time.time() - t0:.2f}s")
    return pages[::-1]


# Tách thông tin bằng regex

FIELDS = ["nha_xuat_ban", "ban_bien_tap", "isbn", "luu_chieu", "luu_chieu_thang", "luu_chieu_nam",
          "giam_doc_tong_bien_tap", "che_ban", "trinh_bay_bia", "dia_chi_nxb", "dien_thoai_nxb",
          "sdt_quan_ly_xuat_ban", "sdt_tong_bien_tap", "fax_nxb", "dia_chi_co_so_in", "ngay_xac_nhan_dkxb"]

_NHAN_RE = re.compile(
    r"^(?:chiu trach nhiem|pho giam doc|pho tong bien tap|uy vien|sua ban in|doc sach mau|doc mau|quyet dinh|"
    r"email|e-mail|website|in xong|nop luu chieu|chi nhanh|so dang ky|so xac nhan|giay dang ky|xac nhan dang ky|"
    r"ma so|isbn|dia chi|tru so|dien thoai|dt\b|tel\b|fax\b|nha xuat ban|nxb\b|quan ly xuat ban|tong bien tap|"
    r"(?:ban )?bien tap|trinh bay|thiet ke|ve bia|bia\s*[:|]|che ban|giam doc|in\b)")
_TEN_NHAN = {
    "ban_bien_tap": re.compile(r"^(?:ban )?bien tap(?: noi dung| vien| ky thuat)?\s*[:|]\s*"),
    "trinh_bay_bia": re.compile(r"^(?:(?:trinh bay|thiet ke|ve) bia|bia)\s*[:|]\s*"),
    "che_ban": re.compile(r"^che ban(?: vi tinh| dien tu)?\s*[:|]\s*"),
}
_GIAM_DOC_RE = re.compile(r"^giam doc\s*(?:[-–—,]|kiem)?\s*tong bien tap\s*[:|]?\s*")
_NXB_RE = re.compile(r"^(?:nha xuat ban|nxb)\b\s*[:|]?\s*")
_DIA_CHI_RE = re.compile(r"(?:^|[^a-z])(?:dia chi|tru so)\s*[:|]\s*")
_CO_SO_IN_RE = re.compile(r"^in\b[^:]*\b(?:cuon|tai)\b|cong ty|xi nghiep|nha in|co so in")
_ISBN_RE = re.compile(r"(?:\bisbn|^ma so(?: sach)?)\b[^0-9\n]{0,15}([0-9][0-9xX\-–. ]{2,})")
_SDT_RE = re.compile(r"\(?\+?\d[\d .()\-]{5,}\d")
_SDT_HET = re.compile(r"fax|email|e-mail|website|quan ly|tong bien tap|dien thoai|\bdt\b")
_DKXB_RE = re.compile(r"dang ky(?: ke hoach)? xuat ban|dkxb|xac nhan dang ky")
_NGAY_RE = re.compile(r"ngay\s*:?\s*(\d{1,2})\s*(?:[/.\-]|thang)\s*(\d{1,2})\s*(?:[/.\-]|nam)\s*(\d{4})")
_LUU_CHIEU_RE = re.compile(r"in xong|l\w{0,2}u chi\w{0,3}u")  # cả lỗi font / OCR: 'lou chieau'
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
_DAU_TRANG_RE = re.compile(r"chiu trach nhiem xuat ban|in xong|nop luu chieu")
_KHOI_NXB_RE = re.compile(r"^(?:nha xuat ban|nxb\b|dia chi|tru so|dien thoai|dt\b|tel\b|fax\b|email|e-mail|website|"
                          r"quan ly xuat ban|tong bien tap|chi nhanh|van phong)")


# Bỏ dấu tiếng Việt, giữ nguyên độ dài
def _bo_dau(s: str) -> str:
    s = s.replace("đ", "d").replace("Đ", "D")
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").lower()


# Tách text thành dòng sạch
def _dong(text: str) -> tuple[list, list]:
    lines = []
    for ln in (text or "").splitlines():
        ln = unicodedata.normalize("NFC", ln).strip(" \t|")
        ln = re.sub(r"\s+", " ", ln)
        if ln:
            lines.append(ln)
    norms = [_bo_dau(ln) for ln in lines]
    return lines, norms


# Lấy phần text gốc sau nhãn
def _sau(line: str, norm: str, pos: int) -> str:
    if len(line) != len(norm):
        pos = line.find(":") + 1 if ":" in line else len(line)
    return line[pos:].strip(" :|.;,")


# Các dòng tên người nối tiếp sau nhãn
def _ten_tiep(lines: list, norms: list, i: int, toi_da: int) -> list:
    out = []
    for j in range(i + 1, min(len(lines), i + 1 + toi_da)):
        if _NHAN_RE.match(norms[j]) or re.search(r"\d", norms[j]) or len(lines[j]) > 60:
            break
        out.append(lines[j].strip(" .;,"))
    return out


# Nối dòng bị xuống dòng giữa chừng
def _noi_dong(lines: list, i: int, value: str) -> str:
    j = i + 1
    while value and j < len(lines) and lines[j][:1].islower():
        value = f"{value} {lines[j]}".strip(" .;,")
        j += 1
    return value


# Lấy số điện thoại sau nhãn
def _sdt(norms: list, nhan: str, bo_qua: str = "") -> str:
    rx = re.compile(rf"(?:^|[^a-z])(?:{nhan})\s*[:.]?\s*")
    for n in norms:
        if bo_qua and re.search(bo_qua, n):
            continue
        for m in rx.finditer(n):
            doan = n[m.end():]
            het = _SDT_HET.search(doan)
            doan = doan[:het.start()] if het else doan
            so = [s.strip() for s in _SDT_RE.findall(doan)]
            if so:
                return "; ".join(so)
    return ""


# Tách tháng, năm lưu chiểu: từ dòng có 'in xong / lưu chiểu', chỉ xét phần từ chữ đó trở đi (không ra tháng + năm thì
# lấy ngày Quyết định xuất bản);
# dòng đó chưa có năm thì nối dòng sau (bị xuống dòng giữa câu)
def _luu_chieu(lines: list, norms: list) -> tuple:
    cau, thang, nam = "", None, None
    for i, (line, n) in enumerate(zip(lines, norms)):
        k = _LUU_CHIEU_RE.search(n)
        if not k:
            continue
        doan = n[k.start():]
        if (not _NAM_RE.search(doan) and i + 1 < len(norms) and not _LUU_CHIEU_RE.search(norms[i + 1])
                and _NAM_RE.search(norms[i + 1])):
            doan, line = f"{doan} {norms[i + 1]}", f"{line} {lines[i + 1]}"
        hop_le = tim_thang_nam(doan)
        if hop_le:
            return line, *hop_le[0]
        years = _NAM_RE.findall(doan)
        if years:
            cau, nam = line, int(years[-1])
        elif not cau:
            cau = line
    qd = ngay_qdxb(lines, norms)
    if qd:
        return qd[2], qd[0], qd[1]
    return cau, thang, nam


# Tách các trường thông tin xuất bản từ text trang cuối
def extract_thong_tin(text: str) -> dict:
    lines, norms = _dong(text)
    kq = {k: "" for k in FIELDS}
    kq["luu_chieu_thang"] = kq["luu_chieu_nam"] = None
    ten = {k: [] for k in _TEN_NHAN}

    for i, (line, n) in enumerate(zip(lines, norms)):
        for key, rx in _TEN_NHAN.items():
            m = rx.match(n)
            if m:
                v = _sau(line, n, m.end())
                ten[key] += ([v] if v else []) + _ten_tiep(lines, norms, i, 5)
                break

        m = _GIAM_DOC_RE.match(n)
        if m and not kq["giam_doc_tong_bien_tap"]:
            v = _sau(line, n, m.end())
            if not v or re.search(r"\d", v):
                v = (_ten_tiep(lines, norms, i, 1) or [""])[0]
            kq["giam_doc_tong_bien_tap"] = v

        m = _NXB_RE.match(n)
        if m and not kq["nha_xuat_ban"] and len(n) > m.end():
            v = _sau(line, n, m.end()) if ":" in line[:m.end() + 1] else line.strip(" .;,")
            kq["nha_xuat_ban"] = _noi_dong(lines, i, v)

        m = _DIA_CHI_RE.search(n)
        if m:
            v = _noi_dong(lines, i, _sau(line, n, m.end()))
            co_so_in = any(_CO_SO_IN_RE.search(norms[j]) for j in range(max(0, i - 3), i + 1))
            nxb = any(_NXB_RE.match(norms[j]) for j in range(max(0, i - 3), i + 1))
            key = "dia_chi_co_so_in" if co_so_in and not nxb else "dia_chi_nxb"
            if v and not kq[key]:
                kq[key] = v

        m = _ISBN_RE.search(n)
        if m and not kq["isbn"]:
            kq["isbn"] = re.sub(r"\s+", "", m.group(1)).strip(".")

        if _DKXB_RE.search(n) and not kq["ngay_xac_nhan_dkxb"]:
            sau = n + (" " + norms[i + 1] if i + 1 < len(norms) and norms[i + 1].startswith("ngay") else "")
            m = _NGAY_RE.search(sau)
            if m and 1 <= int(m.group(2)) <= 12:
                kq["ngay_xac_nhan_dkxb"] = f"{int(m.group(1)):02d}/{int(m.group(2)):02d}/{m.group(3)}"

    for key, ds in ten.items():
        kq[key] = "; ".join(dict.fromkeys(x for x in ds if x))
    kq["dien_thoai_nxb"] = _sdt(norms, r"dien thoai|dt|tel", r"quan ly xuat ban|tong bien tap|bien tap")
    kq["fax_nxb"] = _sdt(norms, r"fax")
    kq["sdt_quan_ly_xuat_ban"] = _sdt(norms, r"quan ly xuat ban(?:\s*[:.]?\s*(?:dt|dien thoai))?")
    kq["sdt_tong_bien_tap"] = _sdt(norms, r"tong bien tap(?:\s*[:.]?\s*(?:dt|dien thoai))?")
    kq["luu_chieu"], kq["luu_chieu_thang"], kq["luu_chieu_nam"] = _luu_chieu(lines, norms)
    return kq


# Đếm số trường đã có
def so_truong(kq: dict) -> int:
    return sum(bool(kq.get(k)) for k in FIELDS if k not in ("luu_chieu_thang", "luu_chieu_nam"))


# Cắt từ dòng Chịu trách nhiệm xuất bản trở đi
def khoi_trang_cuoi(text: str) -> str:
    lines = [ln for ln in (text or "").splitlines() if ln.strip()]
    idx = [i for i, ln in enumerate(lines) if "chiu trach nhiem xuat ban" in _bo_dau(ln)]
    if not idx:
        return "\n".join(lines)
    start = idx[-1]
    while start > 0 and _KHOI_NXB_RE.search(_bo_dau(lines[start - 1].strip())):
        start -= 1
    return "\n".join(lines[start:])


# Trang có phải trang thông tin xuất bản không
def _la_trang_cuoi(text: str) -> bool:
    return bool(_DAU_TRANG_RE.search(_bo_dau(text or "")))


# Các bước của luồng

EMPTY_RESULT = {"thong_tin": {**{k: "" for k in FIELDS}, "luu_chieu_thang": None, "luu_chieu_nam": None},
                "thong_tin_text": "", "thong_tin_source": "", "thong_tin_page": None,
                "thong_tin_error": "", "thong_tin_word_reason": ""}


# Bước 1: đọc Word trang cuối, không được thì render PDF
def info_prepare(context: dict) -> None:
    job_id = context.get("job_id", "")
    pdf_path, word_path = context.get("pdf_path", ""), context.get("docx_path", "")
    if pdf_path and not context.get("page_count"):
        context["page_count"] = pdf_page_count(pdf_path)
    log = get_logger("metadata", job_id or None)
    context.update(info={"text": "", "page": None, "pages": {}}, info_page_paths=[],
                   info_source="", info_word_reason="")
    if word_path:
        pages, problem = read_word_pages(word_path, INFO_LAST_PAGES, job_id, from_end=True)
        if not problem:
            block = khoi_trang_cuoi("\n".join(pages))
            n = so_truong(extract_thong_tin(block))
            if _la_trang_cuoi(block) or n >= INFO_MIN_FIELDS:
                context["info"]["text"] = block
                context["info_source"] = "word"
                log.info(f"Lấy thông tin trang cuối từ Word {os.path.basename(word_path)}: {n} trường")
                return
            problem = f"không thấy trang thông tin xuất bản trong {INFO_LAST_PAGES} trang cuối Word"
        context["info_word_reason"] = problem
        log.warning(f"Word không dùng được: {problem}" + (" → OCR trang cuối PDF" if pdf_path else ""))
    if not pdf_path:
        if word_path:
            context["info_source"] = "word"
            return
        raise ValueError("Thiếu file PDF")
    context["info_page_paths"] = render_last_pages(pdf_path, os.path.join(context["job_folder"], "info_pages"),
                                                   context["page_count"], INFO_LAST_PAGES, job_id)
    context["info_source"] = "ocr"


# Bước 2: OCR cả trang từ trang cuối lên, dừng ở trang thông tin xuất bản
def info_ocr(context: dict) -> None:
    if context.get("info_source") != "ocr":
        return
    job_id = context.get("job_id", "")
    log = get_logger("metadata", job_id or None)
    info = {"text": "", "page": None, "pages": {}}
    try:
        for no, path in context.get("info_page_paths", []):
            text = chandra_ocr(cv2.imread(path), max_tokens=CHANDRA_INFO_MAX_TOKENS,
                               label=f"trang {no}", category="metadata", job_id=job_id)
            info["pages"][str(no)] = text
            if _la_trang_cuoi(text):
                info.update(text=text, page=no)
                break
        else:
            if info["pages"]:
                no, text = max(info["pages"].items(), key=lambda x: so_truong(extract_thong_tin(x[1])))
                info.update(text=text, page=int(no))
                log.warning(f"Không thấy trang thông tin xuất bản → dùng trang {no}")
    except Exception as e:
        log.error(f"OCR trang cuối lỗi: {e}", exc_info=True)
        info["error"] = str(e)
    context["info"] = info


# Bước 3: tách các trường bằng regex
def info_result(context: dict) -> dict:
    job_id = context.get("job_id", "")
    info = context.get("info", {})
    text = info.get("text", "")
    kq = extract_thong_tin(khoi_trang_cuoi(text))
    for no, other in info.get("pages", {}).items():
        if str(no) != str(info.get("page")):
            for k, v in extract_thong_tin(other).items():
                if v and not kq.get(k):
                    kq[k] = v
    source = context.get("info_source", "")
    if text or source:
        get_logger("metadata", job_id or None).log(
            logging.INFO if so_truong(kq) else logging.WARNING,
            f"Thông tin trang cuối ({source or '-'}): {so_truong(kq)}/{len(FIELDS) - 2} trường "
            f"{ {k: v for k, v in kq.items() if v} }")
    return {
        "thong_tin": kq,
        "thong_tin_text": text,
        "thong_tin_source": source,
        "thong_tin_page": info.get("page"),
        "thong_tin_error": info.get("error", ""),
        "thong_tin_word_reason": context.get("info_word_reason", ""),
    }
