"""
📖 Trans-Tool Web Interface
AI Novel Translation Tool with Diff View
"""

import streamlit as st
import difflib
import os
import sys
import time
import html as html_lib
import threading
import random
import json
from datetime import datetime, timezone, timedelta

def now_gmt7():
    return datetime.now(timezone(timedelta(hours=7)))
from dotenv import load_dotenv
import re
import shutil

# ============================================================
# CONFIG & PATHS
# ============================================================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
load_dotenv(os.path.join(BASE_DIR, '.env'))

def get_env(key: str, default=None):
    """Read from st.secrets (Streamlit Cloud) first, then fall back to os.environ (local)."""
    try:
        return st.secrets[key]
    except Exception:
        return os.environ.get(key, default)

def get_windows_sort_key(filename):
    """Hỗ trợ sort Natural và format copy của Windows: '60.jpg' vs '60 (1).jpg'"""
    m = re.match(r'^(.*?)(?: \(([0-9]+)\))?(\.[a-zA-Z0-9_]+)?$', filename)
    if m:
        base, dup_num, ext = m.groups()
        base_parts = [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', base or "")]
        return base_parts + [int(dup_num) if dup_num else 0, ext or ""]
    return [filename]

# ============================================================
# LOGGING SYSTEM  
# ============================================================
LOGS_DIR = os.path.join(BASE_DIR, 'logs')
os.makedirs(LOGS_DIR, exist_ok=True)

# 50 con vật dễ thương — mỗi session được gán 1 tên ngẫu nhiên
ANIMAL_TOKENS = [
    "🦫 Hải Ly",     "🦊 Cáo",        "🐼 Gấu Trúc",   "🐧 Chim Cánh Cụt",
    "🦉 Cú Mèo",    "🦁 Sư Tử",      "🐯 Hổ",         "🐨 Koala",
    "🦒 Hươu Cao Cổ","🦓 Ngựa Vằn",  "🐘 Voi",         "🦏 Tê Giác",
    "🦛 Hà Mã",     "🐆 Báo",        "🐺 Sói",         "🦌 Nai",
    "🦔 Nhím",      "🐿️ Sóc",        "🐰 Thỏ",         "🦘 Kangaroo",
    "🐸 Ếch",       "🦎 Kỳ Nhông",   "🐢 Rùa",         "🦑 Mực",
    "🐙 Bạch Tuộc", "🦞 Tôm Hùm",   "🦀 Cua",         "🐡 Cá Nóc",
    "🐬 Cá Heo",    "🦈 Cá Mập",     "🦢 Thiên Nga",   "🦩 Hồng Hạc",
    "🦜 Vẹt",       "🦚 Công",        "🦃 Gà Tây",      "🦤 Chim Dodo",
    "🐦 Chim Sẻ",   "🦆 Vịt",         "🦅 Đại Bàng",    "🦋 Bướm",
    "🐝 Ong",       "🪲 Bọ Cánh Cứng","🦗 Dế",          "🕷️ Nhện",
    "🦂 Bọ Cạp",   "🐊 Cá Sấu",     "🦭 Hải Cẩu",    "🐻 Gấu",
    "🐮 Bò",        "🐷 Lợn",
]

def _get_cookie_manager():
    """CookieManager must be initialized in every run to handle browser communication."""
    import extra_streamlit_components as stx
    if 'cookie_manager' not in st.session_state:
        st.session_state['cookie_manager'] = stx.CookieManager(key="trans_tool_cookies")
    return st.session_state['cookie_manager']

def assign_animal_token() -> str:
    """
    Assign a random animal name that persists in the browser cookie across F5 reloads.
    Resets only when the user clears browser cookies/cache, or the app is redeployed.
    """
    # 1. Fast path: already in session_state this run
    if 'animal_token' in st.session_state:
        return st.session_state['animal_token']

    # 2. Try reading synchronously from browser headers (fixes F5 reset issue)
    try:
        import urllib.parse
        cookies_str = ""
        if hasattr(st, "context") and hasattr(st.context, "headers"):
            cookies_str = st.context.headers.get("Cookie", "") or st.context.headers.get("cookie", "")
        for item in cookies_str.split(";"):
            item = item.strip()
            if item.startswith("trans_animal="):
                val = urllib.parse.unquote(item.split("=", 1)[1])
                if val.startswith('"') and val.endswith('"'): val = val[1:-1]
                if val in ANIMAL_TOKENS:
                    st.session_state['animal_token'] = val
                    return val
    except Exception:
        pass

    # 2.5 Fallback to CookieManager
    try:
        cm = _get_cookie_manager()
        existing = cm.get("trans_animal")
        if existing and existing in ANIMAL_TOKENS:
            st.session_state['animal_token'] = existing
            return existing
    except Exception:
        pass

    # 3. First visit / cookie not set — pick a new animal and persist it
    token = random.choice(ANIMAL_TOKENS)
    st.session_state['animal_token'] = token
    try:
        from datetime import timedelta
        cm = _get_cookie_manager()
        cm.set("trans_animal", token, expires_at=now_gmt7() + timedelta(days=365))
    except Exception:
        pass
    return token


def get_device_type() -> str:
    """Get a simple device type label from the User-Agent header."""
    try:
        ua = st.context.headers.get("User-Agent", "")
        if "Mobile" in ua or "Android" in ua or "iPhone" in ua:
            return "Mobile"
        elif "Tablet" in ua or "iPad" in ua:
            return "Tablet"
        elif ua:
            return "Desktop"
    except Exception:
        pass
    return "Unknown"

def log_action(feature: str, details: str = ""):
    """Append one line to the daily log file. Never crashes the app."""
    try:
        today_str = now_gmt7().strftime("%Y-%m-%d")
        log_file = os.path.join(LOGS_DIR, f"{today_str}.log")
        token = assign_animal_token()
        device = get_device_type()
        ts = now_gmt7().strftime("%H:%M:%S")
        entry = f"[{today_str} {ts}] | {token:<18} | {device:<8} | {feature:<22} | {details}\n"
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(entry)
    except Exception:
        pass

st.set_page_config(
    page_title="Trans-Tool | AI Novel Translator",
    page_icon="📖",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Static configuration for Chinese translation - devs can fill/customize here
ZH_PROMPT_CONFIG = """
Dưới đây là quy định xưng hô và tên nhân vật cho bản dịch tiếng Trung -> tiếng Việt:
- Cách xưng hô:
  + Truyện đam mỹ nên cặp đôi chính là nam nam.
  + Nhân vật chính là Ôn Mộ Ngôn.
  + Ngôi kể là ngôi thứ ba (vẫn gọi tên hoặc dùng các đại từ là "anh").
- Chuyển đổi tên riêng (Hán Việt):
  + 温慕言 → Ôn Mộ Ngôn
  + 兰漾 → Lan Dạng
- BẢNG XƯNG HÔ HỘI THOẠI (QUY ĐỊNH CẶP ĐÔI):
  + [Ôn Mộ Ngôn] nói chuyện với [Lan Dạng]: [Ôn Mộ Ngôn] xưng là "tôi" - gọi [Lan Dạng] là "cậu".
  + [Lan Dạng] nói chuyện với [Ôn Mộ Ngôn]: [Lan Dạng] xưng là "tôi" - gọi [Ôn Mộ Ngôn] là "anh".
- GÓC NHÌN DẪN TRUYỆN (NARRATION POV - Ngôi thứ 3 từ góc nhìn của Ôn Mộ Ngôn):
  + Mô tả suy nghĩ, nội tâm của Ôn Mộ Ngôn: Dùng ngôi thứ 3, xưng "anh" khi nhắc đến Ôn Mộ Ngôn, xưng "hắn" khi nhắc đến Lan Dạng.
"""

STATE_FILE = os.path.join(BASE_DIR, 'output', 'cn_to_vi_state.json')

PATHS = {
    'eng_trans': os.path.join(BASE_DIR, 'input', 'trans', 'eng.txt'),
    'kor_trans': os.path.join(BASE_DIR, 'input', 'trans', 'kor.txt'),
    'zh_trans': os.path.join(BASE_DIR, 'input', 'trans', 'zh.txt'),
    'vi_qc': os.path.join(BASE_DIR, 'input', 'qc', 'vi_to_qc.txt'),
    'kor_qc': os.path.join(BASE_DIR, 'input', 'qc', 'kor.txt'),
    'eng_qc': os.path.join(BASE_DIR, 'input', 'qc', 'eng.txt'),
    'glossary': os.path.join(BASE_DIR, 'glossary', 'glossary.md'),
    'notes': os.path.join(BASE_DIR, 'glossary', 'personal_notes.md'),
    'output': os.path.join(BASE_DIR, 'output', 'vi_final.txt'),
    'output_prev': os.path.join(BASE_DIR, 'output', 'vi_previous.txt'),
    'cn_to_vi': os.path.join(BASE_DIR, 'output', 'cn_to_vi.txt'),
    'qc_report': os.path.join(BASE_DIR, 'output', 'qc_report.txt'),
    'new_terms': os.path.join(BASE_DIR, 'output', 'new_glossary_terms.txt'),
}

HIDE_LOCAL_FILE_OPTION = str(get_env("HIDE_LOCAL_FILE_OPTION")).strip().lower() in ("true", "1", "yes")

# ============================================================
# CUSTOM CSS
# ============================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    .stApp { font-family: 'Inter', sans-serif; }
    .app-header {
        background: linear-gradient(135deg, #4f5b93 0%, #685b8c 100%);
        padding: 1.5rem 2rem; border-radius: 12px; margin-bottom: 1.5rem; color: #e5e9f0;
    }
    .app-header h1 { margin: 0; font-size: 1.8rem; font-weight: 700; }
    .app-header p { margin: 0.3rem 0 0; opacity: 0.85; font-size: 0.95rem; }
    .diff-container {
        font-family: 'Consolas', 'Courier New', monospace; font-size: 13px;
        line-height: 1.6; border-radius: 10px; overflow: hidden;
        border: 1px solid #414559; max-height: 600px; overflow-y: auto;
    }
    .diff-add { background: rgba(81,207,102,0.12); color: #51cf66; padding: 3px 12px; border-left: 3px solid #51cf66; }
    .diff-del { background: rgba(255,107,107,0.12); color: #ff6b6b; padding: 3px 12px; border-left: 3px solid #ff6b6b; }
    .diff-info { background: rgba(140,170,238,0.1); color: #8caaee; padding: 3px 12px; font-weight: 600; }
    .diff-ctx { color: #a5adce; padding: 3px 12px; }
    .glossary-box {
        background: #292c3c; border: 1px solid #414559; border-radius: 10px;
        padding: 1rem; max-height: 500px; overflow-y: auto;
    }
    /* Side-by-side comparison */
    .sbs-table { width: 100%; border-collapse: collapse; font-size: 14px; line-height: 1.7; }
    .sbs-table th {
        background: linear-gradient(135deg, #51576d, #414559); color: #c6d0f5;
        padding: 10px 14px; text-align: left; position: sticky; top: 0; z-index: 1;
    }
    .sbs-table td {
        padding: 8px 14px; border-bottom: 1px solid #414559;
        vertical-align: top; word-wrap: break-word;
    }
    .sbs-table tr:hover td { background: rgba(140,170,238,0.08); }
    .sbs-num { color: #737994; font-size: 12px; text-align: center; min-width: 35px; user-select: none; }
    .sbs-src { color: #a5adce; max-width: 45%; }
    .sbs-vi { color: #e5c890; max-width: 45%; }
    .sbs-empty { color: #51576d; font-style: italic; }
    .sbs-wrap {
        max-height: 650px; overflow-y: auto; border-radius: 10px;
        border: 1px solid #414559; background: #232634;
    }
    .term-hl {
        background-color: rgba(229, 200, 144, 0.15);
        color: #e5c890;
        border-bottom: 1px dashed #e5c890;
        border-radius: 2px;
        padding: 0 2px;
        cursor: help;
        transition: background-color 0.2s;
    }
    /* Sticky Manhwa Logic */
    /* 1. The Column MUST be tall (matching the image) to act as a 'runway' */
    div[data-testid="stColumn"]:has(.sticky-anchor) {
        height: inherit !important;
        min-height: 100% !important;
    }
    
    /* 2. Target the inner block to be sticky and small enough to slide */
    div[data-testid="stColumn"]:has(.sticky-anchor) [data-testid="stVerticalBlock"] {
        position: -webkit-sticky !important;
        position: sticky !important;
        top: 80px !important;
        z-index: 999;
        background: #1a1b26; /* Dark theme background */
        padding: 15px;
        border-radius: 12px;
        box-shadow: 0 8px 32px rgba(0,0,0,0.3);
        border: 1px solid #414559;
        height: auto !important;
        width: 100% !important; /* Ensure it fills the column width */
    }

    /* 3. Broadly allow overflow and prevent clipping */
    [data-testid="stHorizontalBlock"],
    [data-testid="stColumn"],
    [data-testid="stVerticalBlock"],
    [data-baseweb="tab-panel"],
    .main .block-container {
        overflow: visible !important;
    }
</style>
</style>
""", unsafe_allow_html=True)

# ============================================================
# API KEY ROTATOR + RPD TRACKER
# ============================================================
RPD_COUNTER_FILE = os.path.join(LOGS_DIR, "rpd_counter.json")
@st.cache_resource
def get_rpd_lock():
    return threading.Lock()

_rpd_lock = get_rpd_lock()

# RPD limits for each model per API key
RPD_LIMITS = {
    "gemini-3.1-flash-lite-preview": 500,
    "gemini-2.5-flash": 20,
    "gemini-3-flash-preview": 20,
    "gemini-2.5-flash-lite": 20,
}

def _load_rpd_counter() -> dict:
    """Load today's request counts from JSON file."""
    today = now_gmt7().strftime("%Y-%m-%d")
    try:
        if os.path.exists(RPD_COUNTER_FILE):
            with open(RPD_COUNTER_FILE, 'r') as f:
                data = json.load(f)
            if data.get('date') == today:
                return data
    except Exception:
        pass
    return {'date': today, 'counts': {}}

def _save_rpd_counter(data: dict):
    try:
        with open(RPD_COUNTER_FILE, 'w') as f:
            json.dump(data, f)
    except Exception:
        pass

def increment_rpd(key_idx: int, model: str):
    """Increment request count for key_idx and model. Thread-safe."""
    with _rpd_lock:
        data = _load_rpd_counter()
        k = f"{key_idx}_{model}"
        data['counts'][k] = data['counts'].get(k, 0) + 1
        _save_rpd_counter(data)

def get_rpd_counts() -> dict:
    """Return counts dict for today."""
    with _rpd_lock:
        return _load_rpd_counter().get('counts', {})


class GeminiKeyRotator:
    """Thread-safe multi-key rotator with per-model RPD awareness."""
    def __init__(self, clients: list):
        self._clients = clients
        self._idx = 0
        self._lock = threading.Lock()

    @property
    def current(self):
        return self._clients[self._idx]

    @property
    def current_idx(self):
        return self._idx

    @property
    def total(self):
        return len(self._clients)

    def is_near_limit(self, idx: int, model: str, threshold: float = 0.95) -> bool:
        """True if key has used >= threshold of its RPD for the target model."""
        lim = RPD_LIMITS.get(model, 20)
        used = get_rpd_counts().get(f"{idx}_{model}", 0)
        return used >= lim * threshold

    def rotate(self, model: str, reason: str = ""):
        """Rotate to next key. Skips keys near RPD limit for this model if alternatives exist."""
        with self._lock:
            original = self._idx
            for _ in range(self.total):
                self._idx = (self._idx + 1) % self.total
                if not self.is_near_limit(self._idx, model):
                    break
                if self._idx == original:
                    break  # all exhausted, stay
        return self._idx

    def is_exhausted(self, model: str, threshold: float = 0.95) -> bool:
        """True if ALL keys have reached their RPD limit for this model."""
        with self._lock:
            for idx in range(self.total):
                if not self.is_near_limit(idx, model, threshold):
                    return False
        return True

    def ensure_best_key(self, model: str):
        """Before a call, proactively switch if current key is near limit."""
        with self._lock:
            if self.is_near_limit(self._idx, model) and self.total > 1:
                original = self._idx
                for _ in range(self.total):
                    self._idx = (self._idx + 1) % self.total
                    if not self.is_near_limit(self._idx, model):
                        break
                    if self._idx == original:
                        break

@st.cache_resource
def init_rotator():
    import json
    from google import genai
    keys = []
    i = 1
    while True:
        key = get_env(f"GEMINI_API_KEY_{i}") or (get_env("GEMINI_API_KEY") if i == 1 else None)
        if not key or key.strip() == "":
            break
        keys.append(key.strip())
        i += 1
        if i > 20:
            break
    if not keys:
        return None
    clients = [genai.Client(api_key=k) for k in keys]
    return GeminiKeyRotator(clients)

rotator = init_rotator()
client = rotator.current if rotator else None  # kept for backward compat

# ============================================================
# UTILITY FUNCTIONS
# ============================================================
def load_file(path, default=""):
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            return f.read()
    return default

def save_file(path, content):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)

def generate_with_retry(model, contents, system_instruction, status_w=None, retries=8):
    from google.genai import types
    config = types.GenerateContentConfig(system_instruction=system_instruction, temperature=0.3)
    
    fallback_model = "gemini-3.1-flash-lite-preview"
    if rotator and rotator.is_exhausted(model) and model != fallback_model:
        if status_w:
            status_w.warning(f"⚠️ `{model}` đã hết RPD trên toàn bộ Key! Tự động fallback về `{fallback_model}`.")
        model = fallback_model

    for i in range(retries):
        # Proactively rotate if current key is near RPD limit for this specific model
        if rotator:
            rotator.ensure_best_key(model)
            active_client = rotator.current
            key_idx = rotator.current_idx
        else:
            return ""

        key_label = f"Key {key_idx + 1}/{rotator.total}"
        try:
            resp = active_client.models.generate_content(model=model, contents=contents, config=config)
            if resp and resp.text:
                increment_rpd(key_idx, model)  # count only successful calls mapped to this model
                return resp.text
            return ""
        except Exception as e:
            err_str = str(e)
            if "429" in err_str or "503" in err_str or "quota" in err_str.lower() or "resource_exhausted" in err_str.lower() or "unavailable" in err_str.lower() or "permission_denied" in err_str.lower() or "403" in err_str:
                if rotator and rotator.total > 1:
                    new_idx = rotator.rotate(model, reason="429_503_or_403")
                    if status_w:
                        status_w.warning(f"⚠️ [{key_label}] API quá tải/Lỗi Key (503/429)! Chuyển sang Key {new_idx + 1}... (Lần {i+1}/{retries})")
                    time.sleep(5)
                else:
                    if status_w:
                        status_w.warning(f"⚠️ Model đang quá tải (503/429). Chờ 60s... (Lần {i+1}/{retries})")
                    time.sleep(60)
                continue
            elif "payload" in err_str.lower() or "too large" in err_str.lower() or "400" in err_str:
                if status_w: status_w.warning(f"⚠️ Ảnh/Dữ liệu quá nặng. Đang thử lại... (Lần {i+1})")
                time.sleep(5)
                continue
            if status_w: status_w.error(f"❌ Lỗi API [{key_label}]: {e}")
            time.sleep(5)
    return ""

def optimize_image_for_api(img, max_dimension=2048):
    """
    Giảm kích thước ảnh và convert sang định dạng tối ưu để tránh lỗi payload/rate limit
    nhưng vẫn giữ độ nét tương đối cho OCR.
    """
    import PIL.Image
    import io
    
    # Chỉ xử lý nếu ảnh tồn tại và là loại hình ảnh
    if not isinstance(img, PIL.Image.Image):
        return img
        
    # Chuyển đổi sang RGB nếu đang ở định dạng có alpha (RGBA/P)
    if img.mode in ('RGBA', 'P'):
        img = img.convert('RGB')
        
    width, height = img.size
    
    # Resize nếu kích thước vượt ngưỡng (Manhwa dài thì height thường rất lớn)
    if width > max_dimension or height > max_dimension:
        # Tính tỷ lệ thu nhỏ
        ratio = min(max_dimension / width, max_dimension / height)
        new_width = int(width * ratio)
        new_height = int(height * ratio)
        # Sử dụng LANCZOS để giữ nét chữ tốt nhất có thể
        img = img.resize((new_width, new_height), PIL.Image.Resampling.LANCZOS)
    
    # Save qua bộ nhớ đệm dạng thư viện tối ưu BytesIO thay vì truyền thẳng object nặng nề
    img_byte_arr = io.BytesIO()
    # Lưu dưới chuẩn chất lượng JPEG vừa phải để qua cửa Rate Limit (payload limit)
    img.save(img_byte_arr, format='JPEG', quality=85)
    
    # Load lại ảnh nhẹ từ bytes
    img_byte_arr.seek(0)
    optimized_img = PIL.Image.open(img_byte_arr)
    return optimized_img

def trigger_scroll_to_bottom():
    import streamlit.components.v1 as components
    components.html(
        "<script>window.parent.scrollTo({ top: window.parent.document.body.scrollHeight, behavior: 'smooth' });</script>",
        height=0
    )

def render_diff_html(text1, text2):
    """Render unified diff as colored HTML."""
    lines1 = text1.splitlines()
    lines2 = text2.splitlines()
    diff_lines = list(difflib.unified_diff(lines1, lines2, fromfile="Bản cũ", tofile="Bản mới", lineterm=""))
    if not diff_lines:
        return '<div style="text-align:center;color:#51cf66;padding:2rem;font-size:1.1rem;">✅ Hai bản giống nhau hoàn toàn!</div>'
    html = ['<div class="diff-container">']
    for line in diff_lines:
        esc = html_lib.escape(line)
        if line.startswith('+++') or line.startswith('---'):
            html.append(f'<div class="diff-info">{esc}</div>')
        elif line.startswith('+'):
            html.append(f'<div class="diff-add">{esc}</div>')
        elif line.startswith('-'):
            html.append(f'<div class="diff-del">{esc}</div>')
        elif line.startswith('@@'):
            html.append(f'<div class="diff-info">{esc}</div>')
        else:
            html.append(f'<div class="diff-ctx">{esc}</div>')
    html.append('</div>')
    return '\n'.join(html)

def compute_diff_stats(text1, text2):
    sm = difflib.SequenceMatcher(None, text1.splitlines(), text2.splitlines())
    added = deleted = changed = 0
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == 'insert': added += j2 - j1
        elif op == 'delete': deleted += i2 - i1
        elif op == 'replace': changed += max(i2 - i1, j2 - j1)
    return added, deleted, changed

@st.cache_data
def build_highlight_pattern(gl_text, notes_text):
    import re
    terms = set()
    for text in [gl_text, notes_text]:
        if not text: continue
        for line in text.splitlines():
            line = line.strip()
            if not line.startswith('-'): continue
            if '->' in line:
                parts = line.split('->')
                term_part = parts[-1].split('/')[0].strip() # in case of multiple with /
                if term_part: terms.add(term_part)
            else:
                m = re.match(r'- ([A-Za-z0-9\s\w]+)(?:[\(\[]|$)', line)
                if m:
                    name = m.group(1).strip()
                    if name and len(name) > 2: terms.add(name)
                    
    # Only keep terms with length >= 3 to avoid matching common short words
    valid_terms = [re.escape(t) for t in terms if len(t) >= 3]
    valid_terms.sort(key=len, reverse=True) # Sort longest first to prioritize exact full names
    if not valid_terms: return None
    
    # Word boundary doesn't always work perfectly with unicode if not re.UNICODE
    # But Python 3 re defaults to unicode. Using `\b` is generally fine.
    pattern_str = r'\b(' + '|'.join(valid_terms) + r')\b'
    try:
        return re.compile(pattern_str, re.IGNORECASE)
    except:
        return None

# ============================================================
# HEADER
# ============================================================
st.markdown("""<div class="app-header">
    <h1>📖 Trans-Tool</h1>
    <p>AI Novel Translation Tool — Dịch thuật tiểu thuyết thông minh với Diff View</p>
</div>""", unsafe_allow_html=True)

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("### ⚙️ Cấu hình")
    if rotator:
        if rotator.total == 1:
            st.success("🟢 API Key OK (1 key)")
        else:
            st.success(f"🟢 {rotator.total} API Keys — Key {rotator.current_idx + 1} đang dùng")
    else:
        st.error("🔴 Thiếu API Key!")

    # User-facing Model Selection & RPD guide
    model_guide = {
        "gemini-3-flash-preview": "📝 Dịch Thuật",
        "gemini-2.5-flash": "🔍 QC Review",
        "gemini-2.5-flash-lite": "🎨 Truyện Tranh",
        "gemini-3.1-flash-lite-preview": "🛡️ Trợ thủ Fallback (500 RPD)"
    }
    
    st.markdown("**🤖 AI Models / Tự động điều phối**")
    st.caption("Ứng dụng tự động chọn model phù hợp nhất cho từng tác vụ và tự động chạy sang Fallback khi hết Rate Limit.")
    
    # RPD Usage tracker for ALL models — real-time via fragment
    if rotator and rotator.total > 0:
        @st.fragment(run_every=10)
        def _rpd_tracker():
            counts = get_rpd_counts()
            for mod, desc in model_guide.items():
                with st.expander(f"{desc} ({mod})", expanded=True):
                    lim = RPD_LIMITS.get(mod, 20)
                    for idx in range(rotator.total):
                        used = counts.get(f"{idx}_{mod}", 0)
                        label = f"Key {idx+1}"
                        pct = min(used / lim, 1.0) if lim > 0 else 0
                        if pct >= 0.95:
                            color = "#ff6b6b"   # đỏ
                        elif pct >= 0.75:
                            color = "#f0a500"   # cam
                        else:
                            color = "#51cf66"   # xanh
                        st.markdown(
                            f"""
                            <div style='margin-bottom:6px'>
                            <div style='font-size:11px;color:#a5adce;display:flex;justify-content:space-between'>
                                <span>{label}</span><span style='color:{color}'>{used:,} / {lim:,}</span></div>
                            <div style='background:#292c3c;border-radius:4px;height:4px;overflow:hidden'>
                                <div style='width:{pct*100:.1f}%;background:{color};height:100%;border-radius:4px;
                                transition:width 0.3s'></div></div></div>
                            """,
                            unsafe_allow_html=True
                        )
        _rpd_tracker()

    chunk_size = st.slider("Đoạn/chunk (dịch)", 5, 30, 15, 5)


    st.divider()
    # Log viewer
    st.markdown("### 📋 Activity Logs")
    log_dates = sorted(
        [f.replace('.log', '') for f in os.listdir(LOGS_DIR) if f.endswith('.log')],
        reverse=True
    )
    if not log_dates:
        st.caption("Chưa có log nào.")
    else:
        selected_date = st.selectbox("Chọn ngày:", log_dates, key="log_date_sel")
        log_path = os.path.join(LOGS_DIR, f"{selected_date}.log")
        if os.path.exists(log_path):
            with open(log_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            st.caption(f"{len(lines)} sự kiện")
            # Show last 50 entries, newest first
            log_text = "".join(reversed(lines[-50:]))
            st.code(log_text.strip(), language=None)
            st.download_button(
                "⬇️ Tải log", ''.join(lines),
                f"log_{selected_date}.txt",
                key="log_dl"
            )
    st.divider()
    st.caption(f"📅 {now_gmt7().strftime('%d/%m/%Y %H:%M')}")

# ============================================================
# MAIN NAVIGATION (Persistent on F5)
# ============================================================
MENU_ITEMS = ["🏠 Hướng dẫn", "📝 Dịch Thuật", "🇨🇳 Dịch Trung-Việt", "📊 So Sánh", "📖 Đối Chiếu", "✨ Edit QT", "📚 Glossary"]

tabs = st.tabs(MENU_ITEMS)
current_menu = None # Not used


# Log page visit (once per session)
if 'session_logged' not in st.session_state:
    st.session_state['session_logged'] = True
    log_action("Truy cập", "Mở ứng dụng")

# =================== TAB 0: HOME / HƯỚNG DẪN ===================
with tabs[0]:
    st.markdown("""
    ## Chào mừng bạn đến với **Trans-Tool** 👋  
    *Công cụ hỗ trợ Dịch thuật tiểu thuyết thông minh tích hợp AI cực mạnh do Team xây dựng.*

    Dưới đây là cẩm nang nhanh để bạn nắm rõ các chức năng và cách vận hành:

    ---

    ### 🧩 Các Tính Năng Cốt Lõi

    #### **1. 📝 Dịch Thuật EN/KR (Bám sát Source Eng & Hàn)**  
    - AI nhận song song **Tiếng Anh (EN)** và **Tiếng Hàn (KR)** để cho ra bản dịch Tiếng Việt chuẩn xác nhất.
    - AI tuân thủ nghiêm ngặt từ điển thuật ngữ (`Glossary`) và các ghi chú dịch thuật riêng (`Notes`).
    - Tính năng **Re-Refine** hữu ích khi bạn đã có 1 bản dịch từ trước mà chỉ muốn AI sửa lại cấu trúc/văn phong, giúp tiết kiệm chi phí! 

    #### **2. 🇨🇳 Dịch Trung-Việt (Tiếng Trung Giản Thể)**  
    - Dịch văn bản tiếng Trung Giản Thể sang Tiếng Việt.
    - Hỗ trợ lưu trữ tiến trình thông minh: dịch theo từng phần, F5 không mất dữ liệu và hỗ trợ dịch lại chỉ các phần bị lỗi.
    - Cấu hình tĩnh `ZH_PROMPT_CONFIG` giúp quy định cách xưng hô và dịch tên nhân vật nhất quán và dễ dàng tùy chỉnh bởi nhà phát triển.

    #### **3. 📊 So Sánh (Diff View)**
    - So sánh hai phiên bản dịch cũ và mới để thấy rõ sự khác biệt (thêm, xóa, sửa), rất hữu ích sau khi Re-Refine hoặc cập nhật bản dịch mới.

    #### **4. 📖 Đối Chiếu (Side-by-Side Review)**
    - Hỗ trợ dàn trang **Bản Dịch Tiếng Việt** nằm CẠNH **Bản Dịch Gốc** để dễ dàng đối chiếu, rà soát. Tích hợp bôi sáng thuật ngữ (Highlight) từ Glossary.

    #### **5. ✨ Edit QT (Convert to VI)**
    - Chuyển đổi văn bản Convert/QT (VietPhrase) thô cứng thành văn phong Việt mượt mà, tự nhiên và đúng ngữ pháp.

    ---

    ### 📋 Cách Sử Dụng
    * Hệ thống đã được cấu hình chung từ điển AI (Glossary) siêu xịn, nên các bạn dịch cứ thỏa sức nhé.
    * **Tuyệt đối ưu tiên tính năng Mặc định: 📋 PASTE (Dán Văn Bản)** ở tất cả các tab vì việc Copy/Paste trực tiếp nhanh hơn rất nhiều trong một quy trình làm việc Team.
    
    ### 🛡️ Cơ chế điều phối API Keys (Tự động)
    - App được tích hợp siêu luân phiên tới **20 API Keys** để tự động nhảy sang key khác khi một key hết hạn ngạch.
    - Cột `Cấu Hình` bên tay trái biểu thị sức khỏe (máu báo hiệu xanh/đỏ) của các Models thông minh, bạn có thể tự tin sử dụng mà chẳng âu lo.
    """)

# =================== TAB 1: DỊCH THUẬT ===================
with tabs[1]:
    if not client:
        st.warning("⚠️ Cấu hình API Key trong `.env` trước.")
        st.stop()

    st.markdown("#### 📥 Dữ liệu đầu vào")
    src_opts = ["📋 Paste"] if HIDE_LOCAL_FILE_OPTION else ["📂 File có sẵn (input/trans/)", "📋 Paste"]
    src_idx = 0 if len(src_opts) == 1 else 1
    src = st.radio("Nguồn:", src_opts, index=src_idx, horizontal=True, key="t_src")

    if not src.startswith("📋"):
        eng_text = load_file(PATHS['eng_trans'])
        kor_text = load_file(PATHS['kor_trans'])
        c1, c2 = st.columns(2)
        with c1: st.info(f"EN: {len(eng_text.splitlines())} dòng" if eng_text else "⚠️ Chưa có eng.txt")
        with c2: st.info(f"KR: {len(kor_text.splitlines())} dòng" if kor_text else "⚠️ Chưa có kor.txt")
    else:
        c1, c2 = st.columns(2)
        with c1: eng_text = st.text_area("Tiếng Anh", height=220, key="t_en")
        with c2: kor_text = st.text_area("Tiếng Hàn", height=220, key="t_kr")

    mode = st.radio("Chế độ:", ["🔄 Dịch mới (Draft+Refine)", "✨ Re-Refine (chỉnh vi_final)"], horizontal=True, key="t_mode")

    if st.button("🚀 Bắt đầu dịch", type="primary"):
        target_model = "gemini-3-flash-preview"
        log_action("Dịch Thuật", f"Chế độ: {'Re-Refine' if mode.startswith('✨') else 'Dịch mới'} | EN: {len((eng_text or '').splitlines())} dòng | Model: AUTO")

        if not eng_text.strip() and not kor_text.strip():
            st.error("❌ Thiếu dữ liệu EN hoặc KR! Vui lòng nhập ít nhất một ngôn ngữ nguồn.")
            st.stop()

        glossary = load_file(PATHS['glossary'])
        notes = load_file(PATHS['notes'])

        # Backup for diff
        prev = load_file(PATHS['output'])
        if prev: save_file(PATHS['output_prev'], prev)

        eng_p = [p.strip() for p in eng_text.split('\n') if p.strip()]
        kor_p = [p.strip() for p in kor_text.split('\n') if p.strip()]
        is_refine = mode.startswith("✨")

        draft_p = []
        if is_refine:
            dt = load_file(PATHS['output'])
            if not dt:
                st.error("❌ Không có vi_final.txt để re-refine!")
                st.stop()
            draft_p = [p.strip() for p in dt.split('\n') if p.strip()]

        n_chunks = (max(len(eng_p), len(kor_p)) + chunk_size - 1) // chunk_size
        final = [None] * n_chunks
        bar = st.progress(0, "Chuẩn bị...")
        status = st.status(f"🚀 {'Re-Refine' if is_refine else 'Dịch'} — {n_chunks} phần (Smart Chunking & Parallel)", expanded=True)
        t0 = time.time()

        def process_chunk(idx):
            s, e = idx * chunk_size, (idx+1) * chunk_size
            ec = "\n\n".join(eng_p[s:e])
            kc = "\n\n".join(kor_p[s:e])
            
            # Context Aware: get last 2 paragraphs from previous chunk
            prev_ec = ""
            prev_kc = ""
            if idx > 0:
                prev_s = max(0, (idx-1) * chunk_size)
                prev_ec = "\n".join(eng_p[prev_s:s][-2:])
                prev_kc = "\n".join(kor_p[prev_s:s][-2:])

            draft = ""
            if not is_refine:
                sys_d = (
                    "You are a professional novel translator. Translate English into natural Vietnamese. "
                    "STRICT RULE: ONLY include suffixes (-ie,-ah,-ya) IF present in source. "
                    "Output ONLY the translation."
                )
                prompt_d = f"--- PREVIOUS CONTEXT (For reference only) ---\n{prev_ec}\n\n--- TRANSLATE THIS ---\n{ec}" if prev_ec else ec
                draft = generate_with_retry(target_model, prompt_d, sys_d, None)
            else:
                draft = "\n\n".join(draft_p[s:e])

            sys_r = (
                "You are a strict novel editor. Refine Vietnamese translation comparing EN and KR sources. "
                "RULES: 1.Output ONLY final Vietnamese. 2.Follow source dialogue structure. "
                "3.Keep suffixes from EN source. 4.Keep ahjussi,-ssi,-nim,-gun. "
                "5.Follow Glossary. 6.No creative rewriting."
            )
            
            context_str = f"--- PREVIOUS CONTEXT (DO NOT TRANSLATE) ---\nEN: {prev_ec}\nKR: {prev_kc}\n\n" if prev_ec else ""
            pr = f"{context_str}--- GLOSSARY ---\n{glossary}\n\n--- NOTES ---\n{notes}\n\n--- EN ---\n{ec}\n\n--- KR ---\n{kc}\n\n--- DRAFT ---\n{draft}"
            refined = generate_with_retry(target_model, pr, sys_r, None)

            lines = refined.strip().split('\n')
            clean = [l for l in lines if not l.startswith(('*', 'Đây là', 'Bản dịch', 'Tuyệt vời', 'Đã sửa'))]
            return idx, "\n".join(clean)

        import concurrent.futures
        completed = 0
        with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
            futures = [executor.submit(process_chunk, i) for i in range(n_chunks)]
            for future in concurrent.futures.as_completed(futures):
                try:
                    idx, text = future.result()
                    final[idx] = text
                    completed += 1
                    bar.progress(completed / n_chunks, f"Hoàn thành {completed}/{n_chunks} phần...")
                    status.write(f"  ✅ Phần {idx+1} đã xong!")
                except Exception as e:
                    status.write(f"  ❌ Lỗi ở một phần: {e}")
                trigger_scroll_to_bottom()

        bar.progress(1.0, "✅ Hoàn tất!")
        status.update(label=f"✅ Xong trong {time.time()-t0:.0f}s!", state="complete")

        result = "\n\n".join(final)
        # Apply smart quotes
        # Mở rộng để thay thế cặp "" và ''
        import re
        result = re.sub(r'"([^"]*)"', r'“\1”', result)
        result = re.sub(r"'([^']*)'", r'‘\1’', result)
        save_file(PATHS['output'], result)
        st.session_state['trans_result'] = result
        st.session_state['_t_out_ver'] = st.session_state.get('_t_out_ver', 0) + 1
        st.balloons()

    if 'trans_result' in st.session_state:
        st.divider()
        st.markdown("#### 📤 Kết quả")
        st.text_area("Bản dịch", st.session_state['trans_result'], height=350, key=f"t_out_{st.session_state.get('_t_out_ver', 0)}")
        c1, c2 = st.columns([1, 3])
        with c1:
            st.download_button("⬇️ Tải file", st.session_state['trans_result'],
                               f"vi_final_{now_gmt7().strftime('%Y%m%d_%H%M')}.txt", use_container_width=True)
        with c2:
            st.info("💾 Đã lưu `output/vi_final.txt` | Bản cũ lưu tại `vi_previous.txt`")

# =================== TAB 2: QC REVIEW ===================
# =================== TAB 2: DỊCH TRUNG-VIỆT ===================
with tabs[2]:
    if not client:
        st.warning("⚠️ Cấu hình API Key trong `.env` trước.")
        st.stop()

    def load_zh_state():
        if os.path.exists(STATE_FILE):
            try:
                with open(STATE_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        return None

    def save_zh_state(state):
        try:
            os.makedirs(os.path.dirname(STATE_FILE), exist_ok=True)
            with open(STATE_FILE, 'w', encoding='utf-8') as f:
                json.dump(state, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def delete_zh_state():
        try:
            if os.path.exists(STATE_FILE):
                os.remove(STATE_FILE)
        except Exception:
            pass

    st.markdown("#### 🇨🇳 Dịch thuật Tiếng Trung Giản Thể -> Tiếng Việt")
    
    zh_state = load_zh_state()

    if zh_state is not None:
        source_text = zh_state.get("source_text", "")
        chunk_size_val = zh_state.get("chunk_size", chunk_size)
        paragraphs = zh_state.get("paragraphs", [])
        translated_chunks = zh_state.get("translated_chunks", {})
        n_chunks = len(translated_chunks)
        
        st.info("🔄 Đang có tiến trình dịch dở dang trên hệ thống.")
        
        st.text_area("Văn bản tiếng Trung Giản Thể (Đang dịch)", value=source_text, height=220, disabled=True)
        
        completed = sum(1 for v in translated_chunks.values() if v is not None)
        pct = completed / n_chunks if n_chunks > 0 else 0
        st.progress(pct, f"Đã hoàn thành {completed}/{n_chunks} phần...")
        
        failed_indices = [int(k) for k, v in translated_chunks.items() if v is None]
        
        if failed_indices:
            st.warning(f"⚠️ Có {len(failed_indices)}/{n_chunks} phần chưa hoàn thành hoặc bị lỗi.")
        else:
            st.success("🎉 Tất cả các phần đã được dịch thành công! Đang tự động lưu kết quả...")
            
        c1, c2 = st.columns(2)
        with c1:
            run_lbl = "🚀 Tiếp tục dịch (Dịch các phần chưa hoàn thành / lỗi)" if failed_indices else "🚀 Biên dịch lại toàn bộ"
            if st.button(run_lbl, type="primary", use_container_width=True):
                st.session_state['zh_translating'] = True
                st.rerun()
        with c2:
            if st.button("🆕 Xóa tiến trình & Dịch mới", use_container_width=True):
                delete_zh_state()
                if 'zh_translating' in st.session_state:
                    del st.session_state['zh_translating']
                if 'zh_trans_result' in st.session_state:
                    del st.session_state['zh_trans_result']
                st.rerun()
                
        # Run translation loop if active
        if st.session_state.get('zh_translating', False):
            # If not failed_indices, they clicked "Biên dịch lại toàn bộ", reset
            if not failed_indices:
                translated_chunks = {str(i): None for i in range(n_chunks)}
                zh_state['translated_chunks'] = translated_chunks
                save_zh_state(zh_state)
                failed_indices = list(range(n_chunks))
                
            glossary = load_file(PATHS['glossary'])
            notes = load_file(PATHS['notes'])
            
            bar = st.progress(completed / n_chunks if n_chunks > 0 else 0, "Đang khởi chạy dịch...")
            status = st.status(f"🚀 Tiến trình dịch Trung-Việt — {n_chunks} phần...", expanded=True)
            t0 = time.time()
            
            def process_zh_chunk(idx):
                try:
                    s, e = idx * chunk_size_val, (idx + 1) * chunk_size_val
                    chunk_text = "\n\n".join(paragraphs[s:e])
                    
                    # Context Aware: get last 2 paragraphs from previous chunk
                    prev_zh = ""
                    if idx > 0:
                        prev_s = max(0, (idx - 1) * chunk_size_val)
                        prev_zh = "\n".join(paragraphs[prev_s:s][-2:])
                    
                    sys_instruction = (
                        "You are a professional Chinese-to-Vietnamese novel translator. Translate Simplified Chinese (tiếng Trung giản thể) into natural, fluent Vietnamese. "
                        "STRICT RULES:\n"
                        "1. Output ONLY the translated Vietnamese text. Do NOT include any explanations, annotations, or original Chinese text.\n"
                        "2. Keep dialogue format and punctuation consistent with the source.\n"
                        "3. Follow the names, terminology, and forms of address (xưng hô) specified in the config and glossary."
                    )
                    
                    context_str = f"--- PREVIOUS CONTEXT (DO NOT TRANSLATE, for reference only) ---\n{prev_zh}\n\n" if prev_zh else ""
                    prompt = (
                        f"--- CONFIG & RULES (Forms of address & names) ---\n{ZH_PROMPT_CONFIG}\n\n"
                        f"--- GLOSSARY ---\n{glossary}\n\n"
                        f"--- NOTES ---\n{notes}\n\n"
                        f"{context_str}"
                        f"--- CHINESE SOURCE TO TRANSLATE ---\n{chunk_text}"
                    )
                    
                    target_model = "gemini-3-flash-preview"
                    res = generate_with_retry(target_model, prompt, sys_instruction, None)
                    if not res:
                        return idx, None, "API returned empty (maybe blocked/quota)"
                    lines = res.strip().split('\n')
                    clean = [l for l in lines if not l.startswith(('*', 'Đây là', 'Bản dịch', 'Tuyệt vời', 'Đã sửa'))]
                    final_v = "\n".join(clean).strip()
                    if not final_v:
                        return idx, None, "Empty text after filtering commentary"
                    return idx, final_v, None
                except Exception as e:
                    return idx, None, str(e)
            
            import concurrent.futures
            # Only run for failed/pending chunks
            with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
                futures = {executor.submit(process_zh_chunk, idx): idx for idx in failed_indices}
                for future in concurrent.futures.as_completed(futures):
                    idx = futures[future]
                    try:
                        _, text, err = future.result()
                        if text:
                            translated_chunks[str(idx)] = text
                            zh_state['translated_chunks'] = translated_chunks
                            save_zh_state(zh_state)
                            completed = sum(1 for v in translated_chunks.values() if v is not None)
                            bar.progress(completed / n_chunks, f"Hoàn thành {completed}/{n_chunks} phần...")
                            status.write(f"  ✅ Phần {idx+1} đã dịch xong!")
                        else:
                            status.write(f"  ❌ Phần {idx+1} thất bại: {err or 'Lỗi không rõ'}")
                    except Exception as e:
                        status.write(f"  ❌ Lỗi ở phần {idx+1}: {e}")
                    trigger_scroll_to_bottom()
            
            bar.progress(1.0, "✅ Đã dịch xong lượt này!")
            
            # Check if all completed now
            all_done = all(v is not None for v in translated_chunks.values())
            if all_done:
                status.update(label=f"✅ Hoàn tất toàn bộ {n_chunks} phần trong {time.time()-t0:.0f}s!", state="complete")
                # Compile final output
                final_result = "\n\n".join(translated_chunks[str(i)] for i in range(n_chunks))
                import re
                final_result = re.sub(r'"([^"]*)"', r'“\1”', final_result)
                final_result = re.sub(r"'([^']*)'", r'‘\1’', final_result)
                
                # Backup output
                prev = load_file(PATHS['output'])
                if prev:
                    save_file(PATHS['output_prev'], prev)
                    
                # Save to both paths
                save_file(PATHS['cn_to_vi'], final_result)
                save_file(PATHS['output'], final_result)
                
                # Save to session_state so other tabs see it immediately
                st.session_state['zh_trans_result'] = final_result
                st.session_state['trans_result'] = final_result
                st.session_state['_zh_out_ver'] = st.session_state.get('_zh_out_ver', 0) + 1
                st.session_state['_t_out_ver'] = st.session_state.get('_t_out_ver', 0) + 1
                
                delete_zh_state()
                if 'zh_translating' in st.session_state:
                    del st.session_state['zh_translating']
                st.balloons()
                st.rerun()
            else:
                st.session_state['zh_translating'] = False
                status.update(label="⚠️ Quá trình dịch bị gián đoạn hoặc có phần bị lỗi. Vui lòng bấm 'Tiếp tục dịch' để dịch lại các phần lỗi.", state="error")
                st.rerun()

    else:
        # Standard input interface for new translation
        st.markdown("##### 📥 Dữ liệu đầu vào (Tiếng Trung Giản Thể)")
        zh_src_opts = ["📋 Paste"] if HIDE_LOCAL_FILE_OPTION else ["📂 File có sẵn (input/trans/zh.txt)", "📋 Paste"]
        zh_src_idx = 0 if len(zh_src_opts) == 1 else 1
        zh_src_choice = st.radio("Nguồn:", zh_src_opts, index=zh_src_idx, horizontal=True, key="zh_src_choice")
        
        if not zh_src_choice.startswith("📋"):
            zh_text = load_file(PATHS['zh_trans'])
            st.info(f"ZH (zh.txt): {len(zh_text.splitlines())} dòng" if zh_text else "⚠️ Chưa có file `input/trans/zh.txt`")
        else:
            zh_text = st.text_area("Dán văn bản tiếng Trung Giản Thể", height=300, key="zh_text_area", placeholder="Dán văn bản tiếng Trung cần dịch vào đây (Hỗ trợ tới 5000 dòng)...")
            
        if st.button("🚀 Bắt đầu dịch Trung-Việt", type="primary", use_container_width=True):
            if not zh_text.strip():
                st.error("❌ Vui lòng nhập hoặc cấu hình file văn bản tiếng Trung cần dịch!")
            else:
                paragraphs = [p.strip() for p in zh_text.split('\n') if p.strip()]
                n_chunks = (len(paragraphs) + chunk_size - 1) // chunk_size
                
                # Save new state
                new_state = {
                    "source_text": zh_text,
                    "chunk_size": chunk_size,
                    "paragraphs": paragraphs,
                    "translated_chunks": {str(i): None for i in range(n_chunks)}
                }
                save_zh_state(new_state)
                st.session_state['zh_translating'] = True
                st.rerun()

    # Display final result if exists
    # If the file exists, we can show it so they can always view/download the last translated novel chapter
    last_compiled = load_file(PATHS['cn_to_vi'])
    if last_compiled:
        st.divider()
        st.markdown("#### 📤 Kết quả dịch gần nhất (cn_to_vi.txt)")
        st.text_area("Bản dịch tiếng Việt hoàn chỉnh", last_compiled, height=350, key=f"zh_out_view_{st.session_state.get('_zh_out_ver', 0)}")
        c_dl1, c_dl2 = st.columns([1, 3])
        with c_dl1:
            st.download_button("⬇️ Tải file", last_compiled,
                               f"vi_chinese_{now_gmt7().strftime('%Y%m%d_%H%M')}.txt", use_container_width=True)
        with c_dl2:
            st.success("💾 Đã lưu tại `output/cn_to_vi.txt` và `output/vi_final.txt` | Sẵn sàng để So Sánh/Đối Chiếu!")

# =================== TAB 3: SO SÁNH (DIFF) ===================
with tabs[3]:
    st.markdown("#### 📊 So sánh bản dịch")
    st.caption("So sánh hai phiên bản dịch để thấy sự khác biệt — hữu ích sau khi Re-Refine.")

    diff_font = st.slider("🔤 Cỡ chữ (px)", 13, 22, 15, 1, key="diff_font")

    diff_src_opts = ["📋 Paste thủ công"] if HIDE_LOCAL_FILE_OPTION else ["📂 vi_previous.txt ↔ vi_final.txt (tự động)", "📋 Paste thủ công"]
        
    diff_src_idx = 0 if len(diff_src_opts) == 1 else 1
    diff_src = st.radio("Nguồn dữ liệu:", diff_src_opts, index=diff_src_idx, horizontal=True, key="d_src")

    if not diff_src.startswith("📋"):
        old_text = load_file(PATHS['output_prev'])
        new_text = load_file(PATHS['output'])
        if not old_text:
            st.warning("⚠️ Chưa có `vi_previous.txt`. Hãy dịch hoặc Re-Refine 1 lần để tạo bản backup.")
        elif not new_text:
            st.warning("⚠️ Chưa có `vi_final.txt`.")
        else:
            st.success(f"✅ Bản cũ: {len(old_text.splitlines())} dòng | Bản mới: {len(new_text.splitlines())} dòng")
    else:
        c1, c2 = st.columns(2)
        with c1:
            old_text = st.text_area("📄 Bản cũ", height=250, key="d_old", placeholder="Paste bản cũ...")
        with c2:
            new_text = st.text_area("📄 Bản mới", height=250, key="d_new", placeholder="Paste bản mới...")

    if st.button("🔍 So sánh", type="primary", key="d_btn"):
        log_action("So Sánh (Diff)", f"Bản cũ: {len((old_text or '').splitlines())} dòng | Bản mới: {len((new_text or '').splitlines())} dòng")
        if not old_text or not new_text:
            st.error("❌ Cần cả hai bản để so sánh!")
        else:
            added, deleted, changed = compute_diff_stats(old_text, new_text)

            st.divider()
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("➕ Thêm mới", f"{added} dòng")
            c2.metric("➖ Xóa bỏ", f"{deleted} dòng")
            c3.metric("✏️ Thay đổi", f"{changed} dòng")
            total_changes = added + deleted + changed
            total_lines = max(len(old_text.splitlines()), len(new_text.splitlines()))
            pct = (total_changes / total_lines * 100) if total_lines else 0
            c4.metric("📊 Tỷ lệ thay đổi", f"{pct:.1f}%")

            diff_html = render_diff_html(old_text, new_text)
            st.markdown(f'<div style="font-size:{diff_font}px;line-height:1.7;">{diff_html}</div>', unsafe_allow_html=True)

            st.session_state['last_diff'] = diff_html

# =================== TAB 4: ĐỐI CHIẾU SIDE-BY-SIDE ===================
with tabs[4]:
    st.markdown("#### 📖 Đối chiếu bản dịch với bản gốc")
    st.caption("Hiển thị song song từng dòng bản dịch và bản gốc để bạn tự đối chiếu, rà soát.")

    sbs_font = st.slider("🔤 Cỡ chữ (px)", 13, 22, 15, 1, key="sbs_font")
    c_s1, c_s2 = st.columns(2)
    with c_s1:
        edit_mode = st.toggle("✏️ Chế độ chỉnh sửa tay", value=False, key="sbs_edit",
                              help="Bật để chỉnh sửa trực tiếp bản dịch VI theo từng dòng")
    with c_s2:
        hl_mode = st.toggle("✨ Highlight Thuật ngữ", value=True, key="sbs_hl",
                            help="Bôi sáng các thuật ngữ có trong Glossary")

    # --- Chọn nguồn ---
    sbs_src_opts = ["📋 Paste thủ công"] if HIDE_LOCAL_FILE_OPTION else ["📂 Từ file có sẵn", "📋 Paste thủ công"]
    sbs_src_idx = 0 if len(sbs_src_opts) == 1 else 1
    sbs_src = st.radio("Nguồn bản gốc:", sbs_src_opts, index=sbs_src_idx, horizontal=True, key="sbs_src")
    sbs_lang = st.radio("Ngôn ngữ gốc hiển thị:", ["🇺🇸 Tiếng Anh (EN)", "🇰🇷 Tiếng Hàn (KR)", "🇺🇸🇰🇷 Cả hai"], horizontal=True, key="sbs_lang")

    if not sbs_src.startswith("📋"):
        sbs_vi = load_file(PATHS['output'])
        sbs_en = load_file(PATHS['eng_trans'])
        sbs_kr = load_file(PATHS['kor_trans'])
        info_parts = []
        if sbs_vi: info_parts.append(f"VI: {len(sbs_vi.splitlines())} dòng")
        if sbs_en: info_parts.append(f"EN: {len(sbs_en.splitlines())} dòng")
        if sbs_kr: info_parts.append(f"KR: {len(sbs_kr.splitlines())} dòng")
        if info_parts:
            st.info(" | ".join(info_parts))
        if not sbs_vi:
            st.warning("⚠️ Chưa có `output/vi_final.txt`. Hãy dịch trước.")
    else:
        # Key widget thay đổi mỗi lần Lưu → Streamlit tạo widget mới, nhận value mới
        _vi_ver = st.session_state.get('_sbs_vi_ver', 0)
        _vi_default = ""
        if '_sbs_vi_pending' in st.session_state:
            _vi_default = st.session_state.pop('_sbs_vi_pending')
        sbs_vi = st.text_area("Bản dịch Tiếng Việt", value=_vi_default, height=180,
                               key=f"sbs_vi_in_{_vi_ver}", placeholder="Paste bản dịch VI...")
        c1, c2 = st.columns(2)
        with c1:
            sbs_en = st.text_area("Bản gốc EN", height=180, key="sbs_en_in", placeholder="Paste bản EN...")
        with c2:
            sbs_kr = st.text_area("Bản gốc KR", height=180, key="sbs_kr_in", placeholder="Paste bản KR...")

    if st.button("📖 Hiển thị đối chiếu", type="primary", key="sbs_btn"):
        log_action("Đối Chiếu (SBS)", f"VI: {len((sbs_vi or '').splitlines())} dòng | Ngôn ngữ: {sbs_lang}")
        if not sbs_vi:
            st.error("❌ Thiếu bản dịch VI!")
        else:
            st.session_state['sbs_current_page'] = 1
            st.session_state['sbs_data'] = {
                'vi': sbs_vi.splitlines(),
                'en': (sbs_en.splitlines() if sbs_en else []),
                'kr': (sbs_kr.splitlines() if sbs_kr else []),
            }

    # --- Hiển thị kết quả (dùng session_state để persist qua reruns) ---
    if 'sbs_data' in st.session_state:
        vi_lines = list(st.session_state['sbs_data']['vi'])
        en_lines = st.session_state['sbs_data']['en']
        kr_lines = st.session_state['sbs_data']['kr']

        show_en = "EN" in sbs_lang or "Cả hai" in sbs_lang
        show_kr = "KR" in sbs_lang or "Cả hai" in sbs_lang
        show_both = show_en and show_kr
        max_lines = max(len(vi_lines), len(en_lines), len(kr_lines))

        # Phân trang
        lines_per_page = 50
        total_pages = max(1, (max_lines + lines_per_page - 1) // lines_per_page)

        # Khởi tạo trạng thái trang nếu chưa có
        if 'sbs_current_page' not in st.session_state:
            st.session_state['sbs_current_page'] = 1
        
        def update_sbs_page(key):
            st.session_state['sbs_current_page'] = st.session_state[key]
            st.session_state['sbs_scroll_top'] = True

        page_options = list(range(1, total_pages + 1))
        page_format = lambda x: f"Trang {x} (dòng {(x-1)*lines_per_page+1}~{min(x*lines_per_page, max_lines)})"

        st.markdown('<div id="sbs-top-anchor"></div>', unsafe_allow_html=True)
        if st.session_state.pop('sbs_scroll_top', False):
            import streamlit.components.v1 as components
            components.html(
                '''<script>
                    const anchor = window.parent.document.getElementById('sbs-top-anchor');
                    if(anchor) anchor.scrollIntoView({behavior: 'smooth'});
                </script>''', height=0
            )

        st.divider()
        c1, c2, c3 = st.columns([1, 2, 1])
        with c2:
            # Dropdown phía TRÊN
            st.selectbox(f"Trang (tổng {total_pages})",
                         page_options,
                         index=st.session_state['sbs_current_page'] - 1,
                         key="sbs_top",
                         on_change=update_sbs_page,
                         args=("sbs_top",),
                         format_func=page_format)
        
        # Lấy giá trị trang hiện tại để tính toán start/end
        page = st.session_state['sbs_current_page']

        st.divider()
        c1, c2, c3 = st.columns([1, 2, 1])
        start = (page - 1) * lines_per_page
        end = min(start + lines_per_page, max_lines)

        if not edit_mode:
            # ===== CHẾ ĐỘ XEM (READ-ONLY) =====
            # Build highlight pattern
            hl_pattern = None
            if hl_mode:
                gl_text_hl = load_file(PATHS['glossary'])
                notes_text_hl = load_file(PATHS['notes'])
                hl_pattern = build_highlight_pattern(gl_text_hl, notes_text_hl)

            html_parts = [f'<div class="sbs-wrap"><table class="sbs-table" style="font-size:{sbs_font}px;line-height:1.8;">']
            if show_both:
                html_parts.append('<tr><th class="sbs-num">#</th><th>🇺🇸 English</th><th>🇰🇷 Korean</th><th>🇻🇳 Tiếng Việt</th></tr>')
            elif show_en:
                html_parts.append('<tr><th class="sbs-num">#</th><th>🇺🇸 English</th><th>🇻🇳 Tiếng Việt</th></tr>')
            else:
                html_parts.append('<tr><th class="sbs-num">#</th><th>🇰🇷 Korean</th><th>🇻🇳 Tiếng Việt</th></tr>')

            for idx in range(start, end):
                num = idx + 1
                vi_l = html_lib.escape(vi_lines[idx]) if idx < len(vi_lines) else '<span class="sbs-empty">—</span>'
                if hl_pattern and idx < len(vi_lines) and vi_lines[idx].strip():
                    vi_l = hl_pattern.sub(r'<span class="term-hl" title="Thuật ngữ Glossary">\1</span>', vi_l)
                
                en_l = html_lib.escape(en_lines[idx]) if idx < len(en_lines) else '<span class="sbs-empty">—</span>'
                kr_l = html_lib.escape(kr_lines[idx]) if idx < len(kr_lines) else '<span class="sbs-empty">—</span>'
                if show_both:
                    html_parts.append(f'<tr><td class="sbs-num">{num}</td><td class="sbs-src">{en_l}</td><td class="sbs-src">{kr_l}</td><td class="sbs-vi">{vi_l}</td></tr>')
                elif show_en:
                    html_parts.append(f'<tr><td class="sbs-num">{num}</td><td class="sbs-src">{en_l}</td><td class="sbs-vi">{vi_l}</td></tr>')
                else:
                    html_parts.append(f'<tr><td class="sbs-num">{num}</td><td class="sbs-src">{kr_l}</td><td class="sbs-vi">{vi_l}</td></tr>')

            html_parts.append('</table></div>')
            st.markdown('\n'.join(html_parts), unsafe_allow_html=True)
            st.caption(f"Hiển thị dòng {start+1} → {end} / {max_lines}")

            # Pagination ở dưới cho chế độ xem
            st.divider()
            cb1, cb2, cb3 = st.columns([1, 2, 1])
            with cb2:
                st.selectbox("Trang dưới",
                             page_options,
                             index=st.session_state['sbs_current_page'] - 1,
                             key="sbs_bottom_view",
                             on_change=update_sbs_page,
                             args=("sbs_bottom_view",),
                             format_func=page_format,
                             label_visibility="collapsed")
        else:
            # ===== CHẾ ĐỘ CHỈNH SỬA TAY (custom rows - tự giãn chiều cao) =====
            st.info("✏️ Chỉnh sửa trực tiếp ô **Tiếng Việt** bên phải. Bấm **💾 Lưu** khi xong.")

            # Xác định bố cục cột dựa vào nguồn hiển thị
            if show_en and show_kr:
                ratios = [1, 5, 5, 7]
                hdr_labels = ["#", "🇺🇸 EN", "🇰🇷 KR", "🇻🇳 Tiếng Việt"]
            elif show_en:
                ratios = [1, 6, 7]
                hdr_labels = ["#", "🇺🇸 EN", "🇻🇳 Tiếng Việt"]
            else:
                ratios = [1, 6, 7]
                hdr_labels = ["#", "🇰🇷 KR", "🇻🇳 Tiếng Việt"]

            # Header
            hdr_cols = st.columns(ratios)
            for i, lbl in enumerate(hdr_labels):
                hdr_cols[i].markdown(f"<small><b>{lbl}</b></small>", unsafe_allow_html=True)
            st.divider()

            # Từng dòng
            for idx in range(start, end):
                vi_val = vi_lines[idx] if idx < len(vi_lines) else ""
                # Tính chiều cao text_area dựa trên số dòng thực tế (tối thiểu 3 dòng)
                n_lines = max(3, len(vi_val.splitlines()) + 1) if vi_val else 3
                ta_height = min(400, max(100, n_lines * 26 + 20))

                cols = st.columns(ratios)
                with cols[0]:
                    st.markdown(
                        f"<div style='padding-top:8px;color:#888;font-size:12px;text-align:center'>{idx+1}</div>",
                        unsafe_allow_html=True
                    )
                src_col_i = 1
                if show_en:
                    with cols[src_col_i]:
                        en_val = html_lib.escape(en_lines[idx]) if idx < len(en_lines) else ""
                        st.markdown(
                            f"<div style='padding:6px 4px;font-size:13px;white-space:pre-wrap;word-break:break-word;line-height:1.5'>{en_val}</div>",
                            unsafe_allow_html=True
                        )
                    src_col_i += 1
                if show_kr:
                    with cols[src_col_i]:
                        kr_val = html_lib.escape(kr_lines[idx]) if idx < len(kr_lines) else ""
                        st.markdown(
                            f"<div style='padding:6px 4px;font-size:13px;white-space:pre-wrap;word-break:break-word;line-height:1.5'>{kr_val}</div>",
                            unsafe_allow_html=True
                        )
                with cols[-1]:
                    # Dùng key duy nhất cho mỗi dòng
                    st.text_area(
                        "VI", value=vi_val, height=ta_height,
                        key=f"vi_edit_p{page}_{idx}",
                        label_visibility="collapsed"
                    )
                st.divider()

            # Dropdown phía DƯỚI (Chế độ Sửa) - đưa lên trước nút Lưu
            st.divider()
            cb1, cb2, cb3 = st.columns([1, 2, 1])
            with cb2:
                st.selectbox("Trang dưới edit",
                             page_options,
                             index=st.session_state['sbs_current_page'] - 1,
                             key="sbs_bottom_edit",
                             on_change=update_sbs_page,
                             args=("sbs_bottom_edit",),
                             format_func=page_format,
                             label_visibility="collapsed")

            col_save, col_info = st.columns([1, 3])
            with col_save:
                if st.button("💾 Lưu thay đổi", type="primary", key="sbs_save"):
                    for idx in range(start, end):
                        new_val = st.session_state.get(
                            f"vi_edit_p{page}_{idx}",
                            vi_lines[idx] if idx < len(vi_lines) else ""
                        )
                        if idx < len(vi_lines):
                            vi_lines[idx] = new_val
                    st.session_state['sbs_data']['vi'] = vi_lines
                    full_text = "\n".join(vi_lines)
                    save_file(PATHS['output'], full_text)
                    # Cập nhật state để phản hồi ngay lập tức
                    st.session_state['_sbs_vi_pending'] = full_text
                    st.session_state['trans_result'] = full_text
                    st.session_state['_sbs_vi_ver'] = st.session_state.get('_sbs_vi_ver', 0) + 1
                    st.rerun()
            with col_info:
                st.caption(f"Đang sửa dòng {start+1} → {end} / {max_lines}")

# =================== TAB 5: EDIT QT (CONVERT -> VI) ===================
with tabs[5]:
    if not client:
        st.warning("⚠️ Cấu hình API Key trước.")
    else:
        st.markdown("#### ✨ Chuyển đổi QT/Convert -> Tiếng Việt chuẩn")
        st.caption("Biến bản Convert/QT (VietPhrase) thô cứng thành văn phong Tiếng Việt mượt mà, đúng ngữ pháp.")

        qt_input = st.text_area("Văn bản QT/Convert (VietPhrase):", height=300, key="edit_qt_in", placeholder="Dán văn bản QT vào đây...")
        
        use_glossary_edit = st.checkbox("Sử dụng Glossary để giữ tên riêng", value=False, key="edit_qt_use_gl")
        
        col_btn1, col_btn2 = st.columns([1, 1])
        with col_btn1:
            btn_main = st.button("🚀 Bắt đầu biên tập", type="primary", key="edit_qt_btn", use_container_width=True)
        with col_btn2:
            if st.session_state.get('edit_qt_active', False):
                if st.button("🛑 Hủy bỏ", key="edit_qt_stop", use_container_width=True):
                    st.session_state['edit_qt_active'] = False
                    st.rerun()

        if btn_main:
            # Reset dữ liệu để bắt đầu mới
            if 'edit_qt_temp_parts' in st.session_state:
                del st.session_state['edit_qt_temp_parts']
            if 'edit_qt_result' in st.session_state:
                del st.session_state['edit_qt_result']
            st.session_state['edit_qt_active'] = True
            st.rerun()

        # Biến điều khiển retry
        if st.session_state.get('edit_qt_retry_trigger', False):
            st.session_state['edit_qt_retry_trigger'] = False
            st.session_state['edit_qt_active'] = True

        # --- LOGIC XỬ LÝ CHÍNH ---
        if st.session_state.get('edit_qt_active', False):
            if not qt_input.strip():
                st.error("❌ Vui lòng nhập văn bản QT.")
                st.session_state['edit_qt_active'] = False
            else:
                # Sử dụng lite model cho batch lớn để tránh Rate Limit (RPD 500)
                target_model = "gemini-3.1-flash-lite-preview"
                log_action("Edit QT", f"QT: {len(qt_input.splitlines())} dòng | Model: {target_model}")
                
                glossary = load_file(PATHS['glossary']) if use_glossary_edit else ""
                lines = [l.strip() for l in qt_input.split('\n') if l.strip()]
                
                q_chunk_size = 25 
                n_chunks = (len(lines) + q_chunk_size - 1) // q_chunk_size
                
                # Khởi tạo hoặc kiểm tra tính nhất quán của temp_parts
                if 'edit_qt_temp_parts' not in st.session_state or len(st.session_state['edit_qt_temp_parts']) != n_chunks:
                    st.session_state['edit_qt_temp_parts'] = [None] * n_chunks
                
                bar = st.progress(0, "Đang xử lý...")
                status = st.status(f"🚀 Tiến trình biên tập — tổng {n_chunks} phần...", expanded=True)
                
                # Hiển thị các phần đã xong trước đó để người dùng an tâm
                done_indices = [i for i, p in enumerate(st.session_state['edit_qt_temp_parts']) if p is not None]
                if done_indices:
                    status.write(f"  ℹ️ Đã có {len(done_indices)}/{n_chunks} phần hoàn thành. Đang bỏ qua và chỉ chạy các phần lỗi...")
                
                sys_edit_qt = (
                    "You are a professional Vietnamese novel editor. "
                    "Your task is to convert 'QT/Convert' (VietPhrase) text into natural, grammatical, and polished Vietnamese. "
                    "RULES:\n1. Fix word order and grammar.\n2. Keep terminology from Glossary.\n3. Keep tone/meaning.\n4. Output ONLY Vietnamese."
                )

                def process_q_chunk(idx):
                    try:
                        s, e = idx * q_chunk_size, (idx + 1) * q_chunk_size
                        chunk_text = "\n\n".join(lines[s:e])
                        prompt = f"--- GLOSSARY ---\n{glossary}\n\n--- QT SOURCE ---\n{chunk_text}"
                        res = generate_with_retry(target_model, prompt, sys_edit_qt, None)
                        if not res:
                            return idx, None, "API returned empty (maybe blocked or quota)"
                        clines = res.strip().split('\n')
                        clean = [cl for cl in clines if not cl.startswith(('*', 'Đây là', 'Bản dịch', 'Đã sửa'))]
                        final_v = "\n".join(clean).strip()
                        return idx, (final_v if final_v else None), ("Filtered to empty" if not final_v else None)
                    except Exception as e:
                        return idx, None, str(e)

                import concurrent.futures
                import time
                t0 = time.time()
                
                pending_indices = [i for i, p in enumerate(st.session_state['edit_qt_temp_parts']) if p is None]
                
                if not pending_indices:
                    st.success("Tất cả các phần đã hoàn thành!")
                else:
                    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
                        # Map index to future
                        future_to_idx = {executor.submit(process_q_chunk, i): i for i in pending_indices}
                        
                        for future in concurrent.futures.as_completed(future_to_idx):
                            idx = future_to_idx[future]
                            try:
                                _, text, err = future.result()
                                if text:
                                    st.session_state['edit_qt_temp_parts'][idx] = text
                                    completed = sum(1 for p in st.session_state['edit_qt_temp_parts'] if p is not None)
                                    bar.progress(completed / n_chunks, f"Xong {completed}/{n_chunks}...")
                                    status.write(f"  ✅ Phần {idx+1}: Thành công")
                                else:
                                    status.write(f"  ❌ Phần {idx+1}: {err or 'API Error'}")
                            except Exception as e:
                                status.write(f"  ❌ Phần {idx+1}: Lỗi - {str(e)[:100]}")

                # Kiểm tra lại tổng thể
                all_done = all(p is not None for p in st.session_state['edit_qt_temp_parts'])
                
                if all_done:
                    result_vi = "\n\n".join(st.session_state['edit_qt_temp_parts'])
                    st.session_state['edit_qt_result'] = result_vi
                    st.session_state['_eqt_ver'] = st.session_state.get('_eqt_ver', 0) + 1
                    # Tắt các cờ để thoát khỏi Loop
                    st.session_state['edit_qt_active'] = False
                    st.session_state['edit_qt_retry_trigger'] = False
                    del st.session_state['edit_qt_temp_parts']
                    status.update(label=f"✅ Hoàn tất toàn bộ {n_chunks} phần!", state="complete")
                    st.balloons()
                    st.rerun()
                else:
                    failed_count = sum(1 for p in st.session_state['edit_qt_temp_parts'] if p is None)
                    status.update(label=f"⚠️ Còn {failed_count} phần chưa hoàn thành.", state="error")
                    st.warning(f"Quá trình biên tập chưa hoàn tất 100%. Bạn có thể nhấn 'Thử lại' để chạy tiếp các phần lỗi.")
                    if st.button("🔄 Thử lại các phần lỗi", key="eqt_retry_btn"):
                        st.rerun()

                result_vi = "\n\n".join([p for p in st.session_state['edit_qt_temp_parts'] if p])
                
                if result_vi:
                    st.session_state['edit_qt_result'] = result_vi
                    st.session_state['_eqt_ver'] = st.session_state.get('_eqt_ver', 0) + 1
                    status.update(label=f"✅ Đã biên tập xong!", state="complete")
                else:
                    st.error("❌ Biên tập thất bại hoặc không có kết quả.")
        
        # --- HIỂN THỊ KẾT QUẢ TẠM THỜI (BACKUP) NẾU LỖI ---
        if 'edit_qt_temp_parts' in st.session_state and any(st.session_state['edit_qt_temp_parts']):
            with st.expander("⚠️ Dữ liệu Backup (Phòng trường hợp lỗi/gián đoạn)"):
                backup_text = "\n\n".join([p for p in st.session_state['edit_qt_temp_parts'] if p])
                st.text_area("Bản dịch đã hoàn thành một phần:", backup_text, height=300, key="eqt_backup_view")
                st.download_button("⬇️ Tải bản backup này", backup_text, f"backup_edit_qt_{int(time.time())}.txt")

        if 'edit_qt_result' in st.session_state:
            st.divider()
            st.markdown("#### 📤 Kết quả Tiếng Việt chuẩn")
            # Sử dụng versioned key để force refresh nội dung text_area
            eqt_ver = st.session_state.get('_eqt_ver', 0)
            st.text_area("Bản dịch mượt", st.session_state['edit_qt_result'], height=400, key=f"edit_qt_out_{eqt_ver}")
            st.download_button("⬇️ Tải bản dịch", st.session_state['edit_qt_result'], f"vi_edit_{int(time.time())}.txt")

# =================== TAB 6: GLOSSARY ===================
with tabs[6]:
    st.markdown("#### 📚 Quản lý Glossary")

    g_tab1, g_tab2, g_tab3 = st.tabs(["📖 Glossary", "📝 Personal Notes", "🔄 Đồng bộ"])

    with g_tab1:
        gl = load_file(PATHS['glossary'])
        if gl:
            st.markdown(f'<div class="glossary-box">{gl[:8000]}{"..." if len(gl)>8000 else ""}</div>', unsafe_allow_html=True)
            st.caption(f"📏 {len(gl)} ký tự | {len(gl.splitlines())} dòng")
        else:
            st.info("Chưa có glossary. Chạy đồng bộ từ Google Sheets.")

    with g_tab2:
        notes = load_file(PATHS['notes'])
        edited = st.text_area("Chỉnh sửa Personal Notes:", notes, height=300, key="g_notes")
        if st.button("💾 Lưu Notes", key="g_save"):
            save_file(PATHS['notes'], edited)
            st.success("✅ Đã lưu!")
            st.rerun()

    with g_tab3:
        st.markdown("Đồng bộ glossary từ Google Sheets bằng script `update_glossary.py`.")
        if st.button("🔄 Chạy đồng bộ", key="g_sync"):
            with st.spinner("Đang đồng bộ..."):
                import subprocess
                result = subprocess.run(
                    [sys.executable, os.path.join(BASE_DIR, 'scripts', 'update_glossary.py')],
                    capture_output=True, text=True, cwd=BASE_DIR, encoding='utf-8'
                )
                if result.returncode == 0:
                    st.success("✅ Đồng bộ thành công!")
                    st.code(result.stdout)
                else:
                    st.error("❌ Lỗi đồng bộ!")
                    st.code(result.stderr)
