# Luồng lấy metadata âm thanh (phần không dùng AI: header bằng ffprobe, thẻ bằng mutagen)
import os
import re
import json
import base64
import shutil
import mimetypes
import subprocess
import unicodedata

import mutagen
from mutagen.flac import Picture
from dotenv import load_dotenv

from .logger import get_logger

load_dotenv()


# Cấu hình

FFPROBE_BIN     = os.getenv("FFPROBE_BIN", "ffprobe")
FFPROBE_TIMEOUT = int(os.getenv("SORT_FFPROBE_TIMEOUT", "60"))

AUDIO_EXT = (".mp3", ".m4a", ".m4b", ".aac", ".wav", ".flac", ".ogg", ".oga", ".opus", ".wma", ".amr", ".aiff", ".aif")

MIME = {".mp3": "audio/mpeg", ".m4a": "audio/mp4", ".m4b": "audio/mp4", ".aac": "audio/aac", ".wav": "audio/wav",
        ".flac": "audio/flac", ".ogg": "audio/ogg", ".oga": "audio/ogg", ".opus": "audio/ogg",
        ".wma": "audio/x-ms-wma", ".amr": "audio/amr", ".aiff": "audio/aiff", ".aif": "audio/aiff"}

NGON_NGU = {"vie": "Tiếng Việt", "vi": "Tiếng Việt", "eng": "Tiếng Anh", "en": "Tiếng Anh", "fra": "Tiếng Pháp",
            "fre": "Tiếng Pháp", "fr": "Tiếng Pháp", "zho": "Tiếng Trung", "chi": "Tiếng Trung", "zh": "Tiếng Trung",
            "rus": "Tiếng Nga", "ru": "Tiếng Nga", "jpn": "Tiếng Nhật", "ja": "Tiếng Nhật", "kor": "Tiếng Hàn",
            "ko": "Tiếng Hàn", "deu": "Tiếng Đức", "ger": "Tiếng Đức", "de": "Tiếng Đức"}

# Trường xuất ra, theo thứ tự file Metadata_03_AmThanh_Video.xlsx
FIELDS = ["ten_tep_goc", "dinh_dang", "dung_luong", "thoi_gian_ghi", "thiet_bi", "gps", "tac_gia_trong_tep",
          "thoi_luong", "codec", "tan_so_kenh_bitrate", "tieu_de", "album", "track",
          "loai_noi_dung", "ngon_ngu", "anh_bia", "ban_quyen_nguon", "tac_gia", "nguoi_doc"]
TEN_TRUONG = {"ten_tep_goc": "Tên tệp gốc", "dinh_dang": "Định dạng (MIME / container)", "dung_luong": "Dung lượng",
              "thoi_gian_ghi": "Thời gian ghi", "thiet_bi": "Thiết bị ghi", "gps": "Tọa độ GPS",
              "tac_gia_trong_tep": "Tác giả ghi trong tệp", "thoi_luong": "Thời lượng", "codec": "Codec / định dạng",
              "tan_so_kenh_bitrate": "Tần số lấy mẫu · kênh · bitrate", "tieu_de": "Tiêu đề (thẻ)",
              "album": "Album (thẻ)", "track": "Track (thẻ)", "loai_noi_dung": "Loại nội dung (gợi ý từ thẻ)",
              "ngon_ngu": "Ngôn ngữ (gợi ý từ thẻ)", "anh_bia": "Ảnh bìa nhúng",
              "ban_quyen_nguon": "Bản quyền / nguồn (gợi ý)", "tac_gia": "Tác giả (gợi ý)",
              "nguoi_doc": "Người đọc (gợi ý)"}

# Tên thẻ theo từng chuẩn (ID3, MP4, Vorbis, RIFF, ASF), so sánh không phân biệt hoa thường
THE = {
    "tieu_de":   ["TIT2", "©nam", "title", "INAM", "Title"],
    "album":     ["TALB", "©alb", "album", "IPRD", "WM/AlbumTitle"],
    "track":     ["TRCK", "trkn", "tracknumber", "track", "ITRK", "WM/TrackNumber"],
    "the_loai":  ["TCON", "©gen", "genre", "IGNR", "WM/Genre"],
    "ngay":      ["TDRC", "TDOR", "©day", "date", "year", "ICRD", "WM/Year"],
    "artist":    ["TPE1", "©ART", "artist", "IART", "Author"],
    "album_artist": ["TPE2", "aART", "albumartist", "album_artist", "WM/AlbumArtist"],
    "composer":  ["TCOM", "©wrt", "composer", "WM/Composer"],
    "nguoi_doc": ["TMCL", "performer", "narrator", "reader", "TXXX:narrator", "TXXX:reader",
                  "----:com.apple.iTunes:narrator", "WM/Narrator"],
    "ghi_chu":   ["COMM", "©cmt", "comment", "description", "ICMT", "Description"],
    "ban_quyen": ["TCOP", "cprt", "copyright", "ICOP", "Copyright"],
    "nxb":       ["TPUB", "publisher", "organization", "label", "WM/Publisher"],
    "ngon_ngu":  ["TLAN", "language", "ILNG", "WM/Language"],
    "phan_mem":  ["TSSE", "©too", "encoder", "encoded_by", "ISFT", "WM/EncodingSettings"],
    "gps":       ["©xyz", "location", "com.apple.quicktime.location.ISO6709", "location-eng"],
    "hang_may":  ["com.apple.quicktime.make", "com.android.manufacturer", "©mak", "make"],
    "model_may": ["com.apple.quicktime.model", "com.android.model", "©mod", "model"],
    "may_ghi":   ["originator"],
}

_NGUOI_DOC_RE = re.compile(r"(?:read by|narrated by|narrator|reader|người đọc|giọng đọc|đọc bởi)\s*[:\-]?\s*([^\n;|]+)",
                           re.IGNORECASE)
_URL_RE = re.compile(r"https?://\S+")
_GPS_RE = re.compile(r"([+-]\d+(?:\.\d+)?)([+-]\d+(?:\.\d+)?)")
_VN_RIENG = re.compile("[ĂăĐđĨĩŨũƠơƯưẠ-ỹ]")


# Đọc header bằng ffprobe, chưa cài thì trả None
def ffprobe(path: str, job_id: str = "") -> dict | None:
    log = get_logger("sound", job_id or None)
    if not shutil.which(FFPROBE_BIN):
        log.warning(f"Không thấy {FFPROBE_BIN} → dùng mutagen đọc thông số kỹ thuật")
        return None
    try:
        out = subprocess.run([FFPROBE_BIN, "-v", "error", "-print_format", "json", "-show_format", "-show_streams",
                              path], capture_output=True, timeout=FFPROBE_TIMEOUT, check=True)
        return json.loads(out.stdout.decode("utf-8", "replace"))
    except Exception as e:
        log.error(f"ffprobe lỗi {os.path.basename(path)}: {e}")
        return None


# Sửa chữ tiếng Việt bị ghi sai mã (thẻ ID3 Latin-1)
def sua_ma(s: str) -> str:
    if not s or _VN_RIENG.search(s):
        return s
    try:
        raw = s.encode("latin-1")
    except UnicodeEncodeError:
        try:
            raw = s.encode("cp1252")
        except UnicodeEncodeError:
            return s
    try:
        return unicodedata.normalize("NFC", raw.decode("utf-8"))
    except UnicodeDecodeError:
        pass
    moi = unicodedata.normalize("NFC", raw.decode("cp1258", "replace"))
    return moi if _VN_RIENG.search(moi) else s


# Đổi giá trị một thẻ mutagen ra chữ
def _chu(v) -> str:
    if hasattr(v, "people"):
        return "; ".join(f"{a}: {b}" if a else b for a, b in v.people)
    if hasattr(v, "text"):
        text = "; ".join(str(x) for x in v.text)
        return sua_ma(text) if getattr(v, "encoding", None) == 0 else text
    if isinstance(v, (list, tuple)):
        return "; ".join(f"{x[0]}/{x[1]}" if isinstance(x, tuple) and len(x) == 2 and x[1] else
                         str(x[0]) if isinstance(x, tuple) else
                         x.decode("utf-8", "replace") if isinstance(x, bytes) else str(x) for x in v)
    return str(v)


# Lấy toàn bộ thẻ chữ (bỏ ảnh, dữ liệu nhị phân)
def doc_the(f) -> dict:
    the = {}
    if not f or not f.tags:
        return the
    for k, v in f.tags.items():
        if k.split(":")[0] in ("APIC", "PRIV", "GEOB", "MCDI", "covr", "metadata_block_picture", "WM/Picture"):
            continue
        if k.startswith(("COMM:", "TXXX:", "USLT:")):
            ten = k.split(":")
            k = ten[0] if ten[0] == "COMM" else f"{ten[0]}:{ten[1]}"
        try:
            chu = _chu(v).strip().strip("\x00")
        except Exception:
            continue
        if chu:
            the[k] = f"{the[k]}; {chu}" if k in the and chu not in the[k] else chu
    return the


# Tìm giá trị theo danh sách tên thẻ
def _lay(the: dict, ff: dict, ten: str) -> str:
    thuong = {k.lower(): v for k, v in the.items()}
    for k in THE[ten]:
        if thuong.get(k.lower()):
            return thuong[k.lower()]
    for k in THE[ten]:
        if ff.get(k.lower()):
            return ff[k.lower()]
    return ""


# Lưu ảnh bìa nhúng ra file
def luu_anh_bia(f, out_dir: str) -> str:
    if not f or not f.tags or not out_dir:
        return ""
    data, mime = None, "image/jpeg"
    try:
        tags = f.tags
        if hasattr(tags, "getall"):
            anh = sorted(tags.getall("APIC"), key=lambda a: a.type != 3)
            if anh:
                data, mime = anh[0].data, anh[0].mime
        if data is None and "covr" in tags and tags["covr"]:
            c = tags["covr"][0]
            data, mime = bytes(c), "image/png" if c.imageformat == 14 else "image/jpeg"
        if data is None and getattr(f, "pictures", None):
            p = sorted(f.pictures, key=lambda a: a.type != 3)[0]
            data, mime = p.data, p.mime
        if data is None and "metadata_block_picture" in tags:
            p = Picture(base64.b64decode(tags["metadata_block_picture"][0]))
            data, mime = p.data, p.mime
    except Exception:
        return ""
    if not data:
        return ""
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, "anh_bia" + (mimetypes.guess_extension(mime or "") or ".jpg"))
    with open(path, "wb") as fo:
        fo.write(data)
    return path


# Định dạng số

def _dung_luong(n: int) -> str:
    for dv in ("B", "KB", "MB", "GB"):
        if n < 1024 or dv == "GB":
            return f"{n:.0f} {dv}" if dv == "B" else f"{n:.2f} {dv}"
        n /= 1024


def _thoi_luong(giay: float) -> str:
    giay = int(round(giay or 0))
    return f"{giay // 3600:02d}:{giay % 3600 // 60:02d}:{giay % 60:02d}"


def _kenh(n, layout: str = "") -> str:
    return {1: "mono", 2: "stereo"}.get(n) or layout or (f"{n} kênh" if n else "")


def _ngay(s: str) -> str:
    m = re.match(r"(\d{4}-\d{2}-\d{2})[T ](\d{2}:\d{2}:\d{2})(?:\.\d+)?(Z)?", s or "")
    return f"{m[1]} {m[2]}" + (" UTC" if m[3] else "") if m else s


def _gps(s: str) -> str:
    m = _GPS_RE.search(s or "")
    return f"{float(m[1]):.6f}, {float(m[2]):.6f}" if m else s


# Thông số kỹ thuật: ưu tiên ffprobe, không có thì mutagen
def ky_thuat(f, probe: dict | None) -> dict:
    kt = {"nguon": "", "container": "", "codec": "", "thoi_luong_giay": None, "bitrate": None,
          "che_do_bitrate": "", "tan_so_mau": None, "so_kenh": None, "kenh": ""}
    if probe:
        fmt = probe.get("format", {})
        st = next((s for s in probe.get("streams", []) if s.get("codec_type") == "audio"), {})
        kt.update(nguon="ffprobe", container=fmt.get("format_long_name") or fmt.get("format_name", ""),
                  codec=st.get("codec_long_name") or st.get("codec_name", ""),
                  thoi_luong_giay=float(st.get("duration") or fmt.get("duration") or 0) or None,
                  bitrate=int(st.get("bit_rate") or fmt.get("bit_rate") or 0) or None,
                  tan_so_mau=int(st.get("sample_rate") or 0) or None, so_kenh=st.get("channels"),
                  kenh=_kenh(st.get("channels"), st.get("channel_layout", "")))
    if f and f.info:
        info = f.info
        ten = type(info).__name__
        if ten == "MPEGInfo":
            codec = f"MP3 (MPEG-{info.version:g} Layer {info.layer})"
            kt["che_do_bitrate"] = str(info.bitrate_mode).split(".")[-1].replace("UNKNOWN", "")
            if info.encoder_info:
                kt["phan_mem_ma_hoa"] = info.encoder_info
        elif ten == "WaveStreamInfo":
            codec = f"PCM {getattr(info, 'bits_per_sample', '')}-bit"
        else:
            codec = getattr(info, "codec_description", "") or getattr(info, "codec", "") or ten.replace("Info", "")
        so_kenh = getattr(info, "channels", None)
        if not kt["nguon"]:
            kt.update(nguon="mutagen", container=type(f).__name__, codec=codec,
                      thoi_luong_giay=getattr(info, "length", None), bitrate=getattr(info, "bitrate", None) or None,
                      tan_so_mau=getattr(info, "sample_rate", None), so_kenh=so_kenh, kenh=_kenh(so_kenh))
    return kt


# Gợi ý người đọc: thẻ người biểu diễn, không có thì tìm trong ghi chú
def _nguoi_doc(the: dict, ff: dict) -> str:
    v = _lay(the, ff, "nguoi_doc")
    if v:
        return "; ".join(x.split(": ", 1)[-1] for x in v.split("; "))
    m = _NGUOI_DOC_RE.search(_lay(the, ff, "ghi_chu"))
    return m[1].strip() if m else ""


# Trích toàn bộ trường không cần AI cho một file
def extract_sound(path: str, out_dir: str = "", job_id: str = "") -> dict:
    log = get_logger("sound", job_id or None)
    ext = os.path.splitext(path)[1].lower()
    try:
        f = mutagen.File(path)
    except Exception as e:
        log.warning(f"mutagen không đọc được {os.path.basename(path)}: {e}")
        f = None
    probe = ffprobe(path, job_id)
    if f is None and probe is None:
        raise ValueError("Không đọc được file âm thanh (mutagen và ffprobe đều lỗi)")
    the = doc_the(f)
    ff = {k.lower(): str(v) for k, v in ((probe or {}).get("format", {}).get("tags") or {}).items()}
    kt = ky_thuat(f, probe)
    kt["dung_luong_byte"] = os.path.getsize(path)
    kt["phan_mem_ma_hoa"] = _lay(the, ff, "phan_mem") or kt.get("phan_mem_ma_hoa", "")

    ngay = _lay(the, ff, "ngay")
    if ff.get("origination_date"):
        ngay = f"{ff['origination_date']} {ff.get('origination_time', '')}".strip()
    elif not ngay and ff.get("creation_time"):
        ngay = ff["creation_time"]
    may = " ".join(x for x in (_lay(the, ff, "hang_may"), _lay(the, ff, "model_may")) if x) or _lay(the, ff, "may_ghi")

    ghi_chu = _lay(the, ff, "ghi_chu")
    url = _URL_RE.search(ghi_chu)
    ban_quyen = _lay(the, ff, "ban_quyen") or _lay(the, ff, "nxb") or (f"Nguồn: {url[0]}" if url else "")
    ngon_ngu = _lay(the, ff, "ngon_ngu")

    tan_so = [f"{kt['tan_so_mau'] / 1000:g} kHz" if kt["tan_so_mau"] else "", kt["kenh"],
              " ".join(x for x in (f"{kt['bitrate'] // 1000} kbps" if kt["bitrate"] else "", kt["che_do_bitrate"]) if x)]
    codec = kt["codec"] + (f" ({kt['phan_mem_ma_hoa']})" if kt["phan_mem_ma_hoa"] else "")

    kq = {
        "ten_tep_goc": os.path.basename(path),
        "dinh_dang": " · ".join(x for x in (MIME.get(ext) or mimetypes.guess_type(path)[0] or "", kt["container"]) if x),
        "dung_luong": _dung_luong(kt["dung_luong_byte"]),
        "thoi_gian_ghi": _ngay(ngay),
        "thiet_bi": may,
        "gps": _gps(_lay(the, ff, "gps")),
        "tac_gia_trong_tep": _lay(the, ff, "artist"),
        "thoi_luong": _thoi_luong(kt["thoi_luong_giay"]) if kt["thoi_luong_giay"] else "",
        "codec": codec,
        "tan_so_kenh_bitrate": " · ".join(x for x in tan_so if x),
        "tieu_de": _lay(the, ff, "tieu_de"),
        "album": _lay(the, ff, "album"),
        "track": _lay(the, ff, "track"),
        "loai_noi_dung": _lay(the, ff, "the_loai"),
        "ngon_ngu": NGON_NGU.get(ngon_ngu.strip().lower(), ngon_ngu),
        "anh_bia": luu_anh_bia(f, out_dir),
        "ban_quyen_nguon": ban_quyen,
        "tac_gia": _lay(the, ff, "composer") or _lay(the, ff, "album_artist") or _lay(the, ff, "artist"),
        "nguoi_doc": _nguoi_doc(the, ff),
    }
    log.info(f"Âm thanh {os.path.basename(path)} ({kt['nguon']}): {so_truong(kq)}/{len(FIELDS)} trường "
             f"{ {k: v for k, v in kq.items() if v and k != 'ten_tep_goc'} }")
    return {"truong": kq, "ky_thuat": kt, "the": the, "the_ffprobe": ff}


# Đếm số trường có giá trị
def so_truong(kq: dict) -> int:
    return sum(bool(kq.get(k)) for k in FIELDS)


# Các bước của luồng

EMPTY_RESULT = {"am_thanh": {k: "" for k in FIELDS}, "am_thanh_ky_thuat": {}, "am_thanh_the": {},
                "am_thanh_error": ""}


# Bước 1: đọc header và thẻ
def sound_prepare(context: dict) -> None:
    path = context.get("audio_path", "")
    if not path:
        raise ValueError("Thiếu file âm thanh")
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    out_dir = os.path.join(context["job_folder"], "sound") if context.get("job_folder") else ""
    try:
        context["sound"] = extract_sound(path, out_dir, context.get("job_id", ""))
    except Exception as e:
        get_logger("sound", context.get("job_id") or None).error(f"Đọc âm thanh lỗi: {e}", exc_info=True)
        context["sound"] = {"error": str(e)}


# Bước 2: trả kết quả
def sound_result(context: dict) -> dict:
    s = context.get("sound", {})
    return {
        "am_thanh": s.get("truong") or dict(EMPTY_RESULT["am_thanh"]),
        "am_thanh_ky_thuat": s.get("ky_thuat", {}),
        "am_thanh_the": {**s.get("the_ffprobe", {}), **s.get("the", {})},
        "am_thanh_error": s.get("error", ""),
    }
