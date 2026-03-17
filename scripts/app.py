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
from datetime import datetime
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

@st.cache_resource
def _get_cookie_manager():
    """Singleton CookieManager — must be cached to avoid multiple component instances."""
    import extra_streamlit_components as stx
    return stx.CookieManager(key="trans_tool_cookies")

def assign_animal_token() -> str:
    """
    Assign a random animal name that persists in the browser cookie across F5 reloads.
    Resets only when the user clears browser cookies/cache, or the app is redeployed.
    """
    # 1. Fast path: already in session_state this run
    if 'animal_token' in st.session_state:
        return st.session_state['animal_token']

    # 2. Try reading from browser cookie
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
        cm.set("trans_animal", token, expires_at=datetime.now() + timedelta(days=365))
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
        today_str = datetime.now().strftime("%Y-%m-%d")
        log_file = os.path.join(LOGS_DIR, f"{today_str}.log")
        token = assign_animal_token()
        device = get_device_type()
        ts = datetime.now().strftime("%H:%M:%S")
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

PATHS = {
    'eng_trans': os.path.join(BASE_DIR, 'input', 'trans', 'eng.txt'),
    'kor_trans': os.path.join(BASE_DIR, 'input', 'trans', 'kor.txt'),
    'vi_qc': os.path.join(BASE_DIR, 'input', 'qc', 'vi_to_qc.txt'),
    'kor_qc': os.path.join(BASE_DIR, 'input', 'qc', 'kor.txt'),
    'eng_qc': os.path.join(BASE_DIR, 'input', 'qc', 'eng.txt'),
    'glossary': os.path.join(BASE_DIR, 'glossary', 'glossary.md'),
    'notes': os.path.join(BASE_DIR, 'glossary', 'personal_notes.md'),
    'output': os.path.join(BASE_DIR, 'output', 'vi_final.txt'),
    'output_prev': os.path.join(BASE_DIR, 'output', 'vi_previous.txt'),
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
    today = datetime.now().strftime("%Y-%m-%d")
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
            if "429" in err_str or "quota" in err_str.lower() or "resource_exhausted" in err_str.lower() or "permission_denied" in err_str.lower() or "403" in err_str:
                if rotator.total > 1:
                    new_idx = rotator.rotate(model, reason="429_or_403")
                    if status_w:
                        status_w.warning(f"⚠️ [{key_label}] Rate limit/Lỗi Key! Chuyển sang Key {new_idx + 1}... (Lần {i+1}/{retries})")
                    time.sleep(3)
                else:
                    if status_w:
                        status_w.warning(f"⚠️ Quá tải API/Lỗi Key. Chờ 65s... (Lần {i+1}/{retries})")
                    time.sleep(65)
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
    
    # RPD Usage tracker for ALL models
    if rotator and rotator.total > 0:
        counts = get_rpd_counts()
        for mod, desc in model_guide.items():
            with st.expander(f"{desc} ({mod})", expanded=True):
                lim = RPD_LIMITS.get(mod, 20)
                for idx in range(rotator.total):
                    used = counts.get(f"{idx}_{mod}", 0)
                    is_active = (idx == rotator.current_idx)
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

    chunk_size = st.slider("Đoạn/chunk (dịch)", 5, 30, 15, 5)

    st.divider()
    st.markdown("### 📂 File status")
    for name, key in [("EN (dịch)", 'eng_trans'), ("KR (dịch)", 'kor_trans'), ("Glossary", 'glossary'), ("Output", 'output')]:
        p = PATHS[key]
        ok = os.path.exists(p) and os.path.getsize(p) > 0
        sz = f"({os.path.getsize(p)/1024:.1f}KB)" if ok else "(chưa có)"
        st.markdown(f"{'✅' if ok else '⬜'} **{name}** {sz}")
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
            for line in reversed(lines[-50:]):
                st.code(line.strip(), language=None)
            st.download_button(
                "⬇️ Tải log", ''.join(lines),
                f"log_{selected_date}.txt",
                key="log_dl"
            )
    st.divider()
    st.caption(f"📅 {datetime.now().strftime('%d/%m/%Y %H:%M')}")

# ============================================================
# MAIN TABS
# ============================================================
tab_home, tab_trans, tab_qc, tab_diff, tab_sbs, tab_manhwa, tab_dl, tab_glossary = st.tabs(["🏠 Hướng dẫn", "📝 Dịch Thuật", "🔍 QC Review", "📊 So Sánh", "📖 Đối Chiếu", "🎨 Truyện Tranh", "📥 Tải Truyện", "📚 Glossary"])

# Log page visit (once per session)
if 'session_logged' not in st.session_state:
    st.session_state['session_logged'] = True
    log_action("Truy cập", "Mở ứng dụng")

# =================== TAB 0: HOME / HƯỚNG DẪN ===================
with tab_home:
    st.markdown("""
    ## Chào mừng bạn đến với **Trans-Tool** 👋  
    *Công cụ hỗ trợ Dịch thuật, Kiểm duyệt QC và Quét Truyện Tranh thông minh tích hợp AI cực mạnh do Team xây dựng.*

    Dưới đây là cẩm nang nhanh để bạn nắm rõ các chức năng và cách vận hành:

    ---

    ### 🧩 Các Tính Năng Cốt Lõi

    #### **1. 📝 Dịch Thuật (Bám sát Source Eng & Hàn)**  
    - AI nhận song song **Tiếng Anh (EN)** và **Tiếng Hàn (KR)** để cho ra bản dịch Tiếng Việt chuẩn xác nhất.
    - AI tuân thủ nghiêm ngặt từ điển thuật ngữ (`Glossary`) và các ghi chú dịch thuật riêng (`Notes`).
    - Tính năng **Re-Refine** hữu ích khi bạn đã có 1 bản dịch từ trước mà chỉ muốn AI sửa lại cấu trúc/văn phong, giúp tiết kiệm chi phí! 

    #### **2. 🔍 QC Review (Trợ lý bắt lỗi tinh vi)**
    - Ném cho AI bản dịch Tiếng Việt của bạn cùng với bản gốc. AI sẽ chỉ ra cụ thể: *Dòng nào sai sót, sai ý nghĩa, sai tên nhân vật hay bỏ sót hậu tố?*
    - **QC** cực kỳ khắt khe: Sẽ liên tục check chéo với Glossary của nhóm.
    - Đồng thời nó cũng tự động nhặt ra các **"Thuật ngữ mới"** chưa có trong từ điển để báo cáo cho Lead cập nhật!

    #### **3. 🎨 Dịch Truyện Tranh (Manhwa Translator)**
    - Quét và trích xuất chữ Hàn từ bong bóng thoại trong hình ảnh, sau đó dịch trực tiếp sang Tiếng Việt.
    - Chữ thoại sẽ được sắp xếp và dịch gọn gàng, liền mạch theo chuẩn khung tranh! Rất tiết kiệm công sức gõ text.

    #### **4. 📖 Đối Chiếu (Side-by-Side Review)**
    - Hỗ trợ dàn trang **Bản Dịch Tiếng Việt** nằm CẠNH **Bản Dịch Tiếng Anh/Hàn** để Edit/QC tự soát.
    - *Có tích hợp Highlight màu* cho những từ ngữ thuộc phạm vi Glossary, giúp bạn soi lỗi thuật ngữ siêu dễ.

    ---

    ### 📋 Cách Sử Dụng
    * Hệ thống đã được cấu hình chung từ điển AI (Glossary) siêu xịn, nên các bạn dịch cứ thỏa sức nhé.
    * **Tuyệt đối ưu tiên tính năng Mặc định: 📋 PASTE (Dán Văn Bản)** ở tất cả các tab vì việc Copy/Paste trực tiếp nhanh hơn rất nhiều trong một quy trình làm việc Team.
    
    ### 🛡️ Cơ chế điều phối API Keys (Tự động)
    - App được tích hợp siêu luân phiên tới **20 API Keys** để tự động nhảy sang key khác khi một key hết hạn ngạch.
    - Cột `Cấu Hình` bên tay trái biểu thị sức khỏe (máu báo hiệu xanh/đỏ) của các Models thông minh, bạn có thể tự tin sử dụng mà chẳng âu lo.
    """)

# =================== TAB 1: DỊCH THUẬT ===================
with tab_trans:
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

        if not eng_text or not kor_text:
            st.error("❌ Thiếu dữ liệu EN hoặc KR!")
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

        bar.progress(1.0, "✅ Hoàn tất!")
        status.update(label=f"✅ Xong trong {time.time()-t0:.0f}s!", state="complete")

        result = "\n\n".join(final)
        save_file(PATHS['output'], result)
        st.session_state['trans_result'] = result
        st.balloons()

    if 'trans_result' in st.session_state:
        st.divider()
        st.markdown("#### 📤 Kết quả")
        st.text_area("Bản dịch", st.session_state['trans_result'], height=350, key="t_out")
        c1, c2 = st.columns([1, 3])
        with c1:
            st.download_button("⬇️ Tải file", st.session_state['trans_result'],
                               f"vi_final_{datetime.now().strftime('%Y%m%d_%H%M')}.txt", use_container_width=True)
        with c2:
            st.info("💾 Đã lưu `output/vi_final.txt` | Bản cũ lưu tại `vi_previous.txt`")

# =================== TAB 2: QC REVIEW ===================
with tab_qc:
    if not client:
        st.warning("⚠️ Cấu hình API Key trước.")
    else:
        st.markdown("#### 📥 Dữ liệu QC")
        qsrc_opts = ["📋 Paste"] if HIDE_LOCAL_FILE_OPTION else ["📂 File (input/qc/)", "📋 Paste"]
        qsrc_idx = 0 if len(qsrc_opts) == 1 else 1
        qsrc = st.radio("Nguồn:", qsrc_opts, index=qsrc_idx, horizontal=True, key="q_src")

        if not qsrc.startswith("📋"):
            vi_t = load_file(PATHS['vi_qc'])
            kr_t = load_file(PATHS['kor_qc'])
            en_t = load_file(PATHS['eng_qc'])
            st.info(f"VI: {len(vi_t.splitlines())} dòng | KR: {len(kr_t.splitlines())} | EN: {len(en_t.splitlines())}")
        else:
            vi_t = st.text_area("Bản dịch VI", height=200, key="q_vi")
            c1, c2 = st.columns(2)
            with c1: kr_t = st.text_area("Tiếng Hàn", height=200, key="q_kr")
            with c2: en_t = st.text_area("Tiếng Anh (tùy chọn)", height=200, key="q_en")

        if st.button("🔍 Chạy QC", type="primary"):
            target_model = "gemini-2.5-flash"
            log_action("QC Review", f"VI: {len((vi_t or '').splitlines())} dòng | KR: {len((kr_t or '').splitlines())} dòng | Model: AUTO")
            if not vi_t or not kr_t:
                st.error("❌ Thiếu VI hoặc KR!")
                st.stop()

            glossary = load_file(PATHS['glossary'])
            notes = load_file(PATHS['notes'])
            vi_lines = vi_t.split('\n')
            kr_lines = kr_t.split('\n')
            en_lines = en_t.split('\n') if en_t else []

            lpc = 50
            nc = (max(len(vi_lines), len(kr_lines)) + lpc - 1) // lpc
            report = [f"# BÁO CÁO QC — {datetime.now().strftime('%d/%m/%Y %H:%M')}\n"]
            new_terms = []

            bar = st.progress(0)
            status = st.status(f"🔍 QC — {nc} phần", expanded=True)

            for i in range(nc):
                si, ei = i*lpc, (i+1)*lpc
                bar.progress(i/nc, f"Phần {i+1}/{nc}...")
                status.write(f"📋 Phần {i+1}/{nc} — dòng {si+1}~{min(ei, len(vi_lines))}")

                vi_chunk = ""
                for idx, ln in enumerate(vi_lines[si:ei], start=si+1):
                    vi_chunk += f"{idx}: {ln}\n"
                kr_chunk = "\n".join(kr_lines[si:ei])
                en_ref = ""
                if en_lines:
                    en_ref = f"==== EN ====\n{chr(10).join(en_lines[si:ei])}\n\n"

                pqc = (
                    f"==== GLOSSARY ====\n{glossary}\n\n==== NOTES ====\n{notes}\n\n"
                    "==== TASK ====\nCompare VI translation with sources. Report by line number.\n"
                    "Format: ### Dòng [N]:\n- [Lỗi]: \"quote\" - reason\n- [Gợi ý]: fix\n\n"
                    f"==== KR ====\n{kr_chunk}\n\n{en_ref}==== VI ====\n{vi_chunk}"
                )
                rc = generate_with_retry(target_model, pqc, "Professional QC Editor.", status)
                if rc and rc.strip():
                    report.append(rc)
                    status.write(f"  ⚠️ Lỗi ở phần {i+1}")
                else:
                    status.write(f"  ✅ Phần {i+1} OK")

                # Extract new terms
                pe = (
                    f"Find terms in SOURCE not in GLOSSARY.\n==== GLOSSARY ====\n{glossary}\n{notes}\n\n"
                    f"==== SOURCE ====\n{kr_chunk}\n\nFormat: - [Term]: [Translation]"
                )
                ec = generate_with_retry(target_model, pe, "Glossary Extractor.", status)
                if ec and ec.strip(): new_terms.append(ec.strip())
                time.sleep(2)

            bar.progress(1.0, "✅ QC xong!")
            status.update(label="✅ QC Hoàn tất!", state="complete")

            rpt = "\n".join([str(r) for r in report if r])
            save_file(PATHS['qc_report'], rpt)
            st.session_state['qc_report'] = rpt

            if new_terms:
                nt = "\n".join(list(set(new_terms)))
                save_file(PATHS['new_terms'], nt)
                st.session_state['new_terms'] = nt

        if 'qc_report' in st.session_state:
            st.divider()
            st.markdown("#### 📋 Báo cáo QC")
            st.markdown(st.session_state['qc_report'])
            st.download_button("⬇️ Tải báo cáo", st.session_state['qc_report'], "qc_report.txt")
            if 'new_terms' in st.session_state:
                with st.expander("✨ Thuật ngữ mới gợi ý"):
                    st.markdown(st.session_state['new_terms'])

# =================== TAB 3: SO SÁNH (DIFF) ===================
with tab_diff:
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
with tab_sbs:
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
        sbs_vi = st.text_area("Bản dịch Tiếng Việt", height=180, key="sbs_vi_in", placeholder="Paste bản dịch VI...")
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

        st.divider()
        c1, c2, c3 = st.columns([1, 2, 1])
        with c2:
            page = st.selectbox(f"Trang (tổng {total_pages} trang, {max_lines} dòng)",
                                range(1, total_pages + 1), key="sbs_page",
                                format_func=lambda x: f"Trang {x} (dòng {(x-1)*lines_per_page+1}~{min(x*lines_per_page, max_lines)})")

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
        else:
            # ===== CHẾ ĐỘ CHỈNH SỬA TAY =====
            import pandas as pd
            st.info("✏️ **Nhấp đúp** vào ô cột **Tiếng Việt** để sửa. Nhấn **Enter** xác nhận, rồi bấm **💾 Lưu**.")

            data = []
            for idx in range(start, end):
                row = {"#": idx + 1}
                if show_en:
                    row["🇺🇸 EN"] = en_lines[idx] if idx < len(en_lines) else ""
                if show_kr:
                    row["🇰🇷 KR"] = kr_lines[idx] if idx < len(kr_lines) else ""
                row["🇻🇳 Tiếng Việt"] = vi_lines[idx] if idx < len(vi_lines) else ""
                data.append(row)

            df = pd.DataFrame(data)

            disabled_cols = ["#"]
            if show_en: disabled_cols.append("🇺🇸 EN")
            if show_kr: disabled_cols.append("🇰🇷 KR")

            col_config = {
                "#": st.column_config.NumberColumn("#", width="small"),
                "🇻🇳 Tiếng Việt": st.column_config.TextColumn("🇻🇳 Tiếng Việt", width="large"),
            }

            edited = st.data_editor(
                df, disabled=disabled_cols, column_config=col_config,
                use_container_width=True, hide_index=True,
                num_rows="fixed", key=f"sbs_editor_p{page}"
            )

            col_save, col_info = st.columns([1, 3])
            with col_save:
                if st.button("💾 Lưu thay đổi", type="primary", key="sbs_save"):
                    for _, row in edited.iterrows():
                        line_idx = int(row["#"]) - 1
                        new_val = str(row["🇻🇳 Tiếng Việt"]) if row["🇻🇳 Tiếng Việt"] else ""
                        if line_idx < len(vi_lines):
                            vi_lines[line_idx] = new_val
                    st.session_state['sbs_data']['vi'] = vi_lines
                    save_file(PATHS['output'], "\n".join(vi_lines))
                    st.success("✅ Đã lưu thay đổi vào `output/vi_final.txt`!")
                    st.balloons()
            with col_info:
                st.caption(f"Đang sửa dòng {start+1} → {end} / {max_lines}")

# =================== TAB 5: TRUYỆN TRANH ===================
with tab_manhwa:
    if not client:
        st.warning("⚠️ Cấu hình API Key trước.")
    else:
        st.markdown("#### 🎨 Dịch Truyện Tranh (Manhwa Translator)")
        st.caption("Tải lên các ảnh của chapter. AI sẽ nhìn ảnh, nhận diện bong bóng thoại (OCR Hàn) và dịch sang tiếng Việt.")
        
        mh_hist_dir = os.path.join(BASE_DIR, 'output', 'manhwa_history')
        os.makedirs(mh_hist_dir, exist_ok=True)
        
        sessions = sorted([d for d in os.listdir(mh_hist_dir) if os.path.isdir(os.path.join(mh_hist_dir, d))], reverse=True)
        
        st.divider()
        c_sess1, c_sess2 = st.columns([3, 1])
        with c_sess1:
            opts = ["+ TẠO PHIÊN BẢN MỚI"] + sessions
            # Retrieve default index from session_state if it exists
            default_ix = 0
            if 'current_mh_sess' in st.session_state and st.session_state['current_mh_sess'] in opts:
                default_ix = opts.index(st.session_state['current_mh_sess'])
                
            sess_choice = st.selectbox("📂 Chọn Chapter / Phiên bản đã lưu:", opts, index=default_ix)
            # Update state with current choice immediately so we remember it
            st.session_state['current_mh_sess'] = sess_choice
            
        with c_sess2:
            st.write(" ")
            if sess_choice != "+ TẠO PHIÊN BẢN MỚI":
                if st.button("🗑️ Xóa Lịch sử này", use_container_width=True):
                    import shutil
                    shutil.rmtree(os.path.join(mh_hist_dir, sess_choice))
                    st.success("✅ Đã xóa!")
                    st.rerun()

        st.divider()

        if sess_choice == "+ TẠO PHIÊN BẢN MỚI":
            # Initialize a stable default name if not exists
            if 'mh_new_sess_def' not in st.session_state:
                st.session_state['mh_new_sess_def'] = f"Chapter_{datetime.now().strftime('%Y%m%d_%H%M')}"
                
            new_sess_name = st.text_input("Tên Chapter mới (Tạo thư mục):", 
                                         value=st.session_state['mh_new_sess_def'],
                                         key="mh_new_sess_input")
            
            # Use the user's input from the widget key to be safe
            final_sess_name = st.session_state["mh_new_sess_input"].strip().replace('/', '-').replace('\\', '-')
            if not final_sess_name:
                final_sess_name = st.session_state['mh_new_sess_def']
            
            uploaded_files = st.file_uploader("🖼️ Chọn ảnh truyện tranh (JPG, PNG, WEBP)", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True)
            stitch_mh = st.checkbox("🧩 Tự động nối dải ảnh trước khi dịch", value=True, help="Nếu ảnh bị cắt ngắn, ghép chúng lại thành dải dài (Stitching) sẽ giúp AI đọc chuẩn xác không bị đứt câu.")
            
            c1, c2 = st.columns([1, 1])
            with c1:
                st.info("Trình thông dịch AI: AUTO 🤖")
            with c2:
                process_btn = st.button("🚀 Bắt đầu Quét & Dịch", type="primary", use_container_width=True)

            if process_btn and uploaded_files:
                target_model = "gemini-2.5-flash-lite"
                log_action("Truyện Tranh", f"Ảnh: {len(uploaded_files)} | Session: {final_sess_name} | Model: AUTO")
                import PIL.Image
                
                # --- STITCHING PRE-PROCESSING ---
                files_to_process = []
                if stitch_mh and len(uploaded_files) > 1:
                    status_stitch = st.status("🧩 Đang phân tích và ghép nối ảnh...", expanded=True)
                    # Sắp xếp thứ tự tự nhiên (Natural Sort) + xử lý Windows Duplicate (VD: ảnh (1) xếp sau ảnh gốc)
                    uploaded_files = sorted(uploaded_files, key=lambda x: get_windows_sort_key(x.name))
                    images = []
                    for uf in uploaded_files:
                        try:
                            im = PIL.Image.open(uf).convert('RGB')
                            images.append((uf.name, im))
                        except Exception:
                            pass
                    
                    if images:
                        stitched_chunks = []
                        current_chunk = []
                        current_h = 0
                        current_w = 0
                        MAX_HEIGHT = 15000
                        
                        for f, img in images:
                            if current_h + img.height > MAX_HEIGHT and current_chunk:
                                stitched_chunks.append((current_chunk, current_w, current_h))
                                current_chunk = [(f, img)]
                                current_h = img.height
                                current_w = img.width
                            else:
                                current_chunk.append((f, img))
                                current_h += img.height
                                current_w = max(current_w, img.width)
                                
                        if current_chunk:
                            stitched_chunks.append((current_chunk, current_w, current_h))
                            
                        class StitchedImg:
                            def __init__(self, name, img):
                                self.name = name
                                self.img = img
                                
                        for i, (chunk, w, h) in enumerate(stitched_chunks):
                            canvas = PIL.Image.new('RGB', (w, h), (255, 255, 255))
                            y_offset = 0
                            for f, img in chunk:
                                x_offset = (w - img.width) // 2
                                canvas.paste(img, (x_offset, y_offset))
                                y_offset += img.height
                            
                            fname = f"stitched_{i+1:03d}.jpg"
                            files_to_process.append(StitchedImg(fname, canvas))
                        
                        status_stitch.update(label=f"✅ Nối ảnh xong! (Ghép thành {len(files_to_process)} dải dài)", state="complete")
                
                if not files_to_process:
                    class RawImg:
                        def __init__(self, name, uf):
                            self.name = name
                            self.img = PIL.Image.open(uf).convert('RGB')
                    files_to_process = [RawImg(uf.name, uf) for uf in uploaded_files]
                # --- TẠO THƯ MỤC VÀ LƯU ẢNH TRƯỚC ĐỂ BACKUP ---
                sess_dir = os.path.join(mh_hist_dir, final_sess_name)
                sess_img_dir = os.path.join(sess_dir, 'images')
                os.makedirs(sess_img_dir, exist_ok=True)
                
                for fo in files_to_process:
                    save_path = os.path.join(sess_img_dir, fo.name)
                    if not os.path.exists(save_path):
                        fo.img.save(save_path, quality=90)
                
                glossary = load_file(PATHS['glossary'])
                notes = load_file(PATHS['notes'])
                
                status = st.status(f"🚀 Đang xử lý {len(files_to_process)} ảnh/dải...", expanded=True)
                bar = st.progress(0)
                
                all_results = []
                t0 = time.time()
                consecutive_errors = 0
                
                for i, file_obj in enumerate(files_to_process):
                    fname = file_obj.name
                    bar.progress(i / len(files_to_process), f"Đang quét ảnh {i+1}/{len(files_to_process)}...")
                    status.write(f"🖼️ Đang quét ảnh: `{fname}`")
                    
                    try:
                        img = file_obj.img
                        # Tối ưu kích thước và dung lượng ảnh trước khi gửi
                        optimized_img = optimize_image_for_api(img)
                        
                        sys_m = (
                            "You are an expert Manhwa/Webtoon translator and typesetter assistant. "
                            "You extract Korean text strictly from speech bubbles or important narrative boxes and translate it into natural, flowing Vietnamese. "
                            "Ignore small background SFX (Sound Effects) unless they are crucial to the plot. "
                            "CRITICAL RULE: If a single speech bubble contains multiple lines of text, you MUST join them into a SINGLE line separated by a space in both the KR and VI output. Do NOT preserve line breaks within the same dialogue box.\n"
                            "Format your output cleanly and exactly like this:\n"
                            "**[Khung thoại]**\n"
                            "KR: <Korean text in a SINGLE line>\n"
                            "VI: <Vietnamese translation in a SINGLE line>\n\n"
                            "Rules: Follow the provided glossary. Ensure pronouns match the Korean nuances and glossary rules."
                        )
                        
                        prompt = f"--- GLOSSARY ---\n{glossary}\n\n--- NOTES ---\n{notes}\n\n--- TASK ---\nExtract dialogues from this image and translate them to Vietnamese. Keep them in reading order (top to bottom, right to left generally)."
                        
                        contents = [optimized_img, prompt]
                        res = generate_with_retry(target_model, contents, sys_m, status)
                        
                        if res and res.strip():
                            all_results.append(f"### 📄 ẢNH: {fname}\n\n{res}\n")
                            consecutive_errors = 0
                            status.write(f"   ✅ Xong `{fname}`")
                            # Incremental progress save
                            save_file(os.path.join(sess_dir, "script.txt"), "\n\n".join(all_results))
                            time.sleep(15) 
                        else:
                            consecutive_errors += 1
                            all_results.append(f"### 📄 ẢNH: {fname}\n\n[LỖI HOẶC HẾT TOKEN]\n")
                            status.write(f"   ⚠️ Thất bại `{fname}`")
                            save_file(os.path.join(sess_dir, "script.txt"), "\n\n".join(all_results))
                            
                            if consecutive_errors >= 2:
                                status.error("🚨 Quá trình dịch liên tục thất bại do Rate Limit hoặc lỗi API! Dừng sớm để bảo toàn dữ liệu.")
                                break
                    except Exception as e:
                        consecutive_errors += 1
                        all_results.append(f"### 📄 ẢNH: {fname}\n\n[LỖI HỆ THỐNG: {e}]\n")
                        status.write(f"   ❌ Lỗi `{fname}`: {e}")
                        save_file(os.path.join(sess_dir, "script.txt"), "\n\n".join(all_results))
                        
                        if consecutive_errors >= 2:
                            status.error("🚨 Quá trình dịch liên tục thất bại do lỗi phần mềm! Dừng sớm để bảo toàn dữ liệu.")
                            break
                        
                bar.progress(1.0, "✅ Hoàn tất tiến trình hiện tại!")
                status.update(label=f"✅ Kết thúc quá trình quét (Save Backup) trong {time.time()-t0:.0f}s", state="complete")
                
                st.balloons()
                st.session_state['current_mh_sess'] = final_sess_name
                if 'mh_new_sess_def' in st.session_state: del st.session_state['mh_new_sess_def']
                st.rerun()
            
        elif sess_choice != "+ TẠO PHIÊN BẢN MỚI":
            # Mode: View and Edit existing history
            sess_dir = os.path.join(mh_hist_dir, sess_choice)
            sess_script_path = os.path.join(sess_dir, "script.txt")
            sess_img_dir = os.path.join(sess_dir, "images")
            
            mh_content = load_file(sess_script_path)
            
            # Helper: Parse master script into a dict of {filename: content}
            import re
            parts = re.split(r'### 📄 ẢNH: (.*?)\n', mh_content)
            parsed_data = {}
            if len(parts) > 1:
                for i in range(1, len(parts), 2):
                    fname = parts[i].strip()
                    text = parts[i+1].strip() if i+1 < len(parts) else ""
                    parsed_data[fname] = text
            
            st.markdown(f"#### 📑 Biên tập Chapter: {sess_choice}")
            
            if os.path.exists(sess_img_dir):
                saved_imgs = sorted(os.listdir(sess_img_dir))
                if saved_imgs:
                    st.divider()
                    
                    # Individual image-text rows
                    new_parsed_data = {}
                    for img_name in saved_imgs:
                        st.markdown(f"🖼️ **{img_name}**")
                        c1, c2 = st.columns([1, 1])
                        with c1:
                            st.image(os.path.join(sess_img_dir, img_name), use_container_width=True)
                        with c2:
                            st.markdown('<div class="sticky-anchor"></div>', unsafe_allow_html=True)
                            current_val = parsed_data.get(img_name, "")
                            input_val = st.text_area("Bản dịch:", current_val, height=500, key=f"mh_edit_{sess_choice}_{img_name}")
                            new_parsed_data[img_name] = input_val
                        st.divider()

                    # Reconstruct script for saving/downloading
                    reconstructed_script = "\n\n".join([f"### 📄 ẢNH: {n}\n\n{new_parsed_data.get(n, '')}\n" for n in saved_imgs])

                    # Bottom Actions
                    col_save, col_dl = st.columns([1, 1])
                    with col_dl:
                         st.download_button("⬇️ Tải Kịch bản (.txt)", reconstructed_script, 
                                           f"{sess_choice}.txt", use_container_width=True)
                    with col_save:
                        if st.button("💾 Lưu tất cả thay đổi", type="primary", use_container_width=True, key="mh_save_all"):
                            save_file(sess_script_path, reconstructed_script)
                            st.success("✅ Đã lưu toàn bộ bản dịch!")
                            st.rerun()
                else:
                    st.info("Chưa có ảnh gốc nào được lưu lại cho Lịch sử này.")
            else:
                st.error("Không tìm thấy thư mục ảnh cho phiên bản này.")

# =================== TAB 6: TẢI TRUYỆN ===================
with tab_dl:
    st.markdown("#### 📥 Tải ảnh truyện tranh hàng loạt")
    st.caption("Công cụ này sử dụng `gallery-dl` để tự động cào ảnh gốc từ các trang truyện (Naver, Kakao, Webtoons...) về rồi nén thành file ZIP cho bạn.")
    
    dl_url = st.text_input("🔗 Nhập Link Truyện (URL):", placeholder="https://comic.naver.com/webtoon/detail?titleId=...")
    stitch_images = st.checkbox("🧩 Tự động nối các dải ảnh bị cắt đứt (Stitching)", value=True, help="Ghép nối liền mạch các ảnh bị cắt ngắn (VD: Webtoon) thành 1 dải ảnh dài. Giới hạn 15,000px cấu hình mỗi file ảnh để tránh quá khổ cho AI và trình xem.")
    
    c1, c2 = st.columns([1, 1])
    with c1:
        submit_dl = st.button("🚀 Bắt đầu Tải & Nén", type="primary", use_container_width=True)
        
    if submit_dl:
        if not dl_url.strip():
            st.error("❌ Link không được bỏ trống!")
        else:
            log_action("Tải Truyện", f"URL: {dl_url[:40]}...")
            import tempfile, subprocess
            
            with st.spinner("⏳ Đang cào ảnh từ Web... Xin vui lòng đợi, quá trình này có thể mất vài phút! (Không F5 hay đóng trang)"):
                tmp_dir = tempfile.mkdtemp()
                zip_path = os.path.join(tempfile.gettempdir(), f"manhwa_{int(time.time())}")
                try:
                    # Execute gallery-dl module
                    res = subprocess.run([sys.executable, "-m", "gallery_dl", "--directory", tmp_dir, dl_url], capture_output=True, text=True)
                    
                    # Optional stitching logic
                    if stitch_images:
                        import PIL.Image
                        for root, dirs, files in os.walk(tmp_dir):
                            img_files = [f for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg', '.webp', '.gif'))]
                            if not img_files:
                                continue
                            
                            # Sort properly using natural sorting + Windows Duplicate handling
                            img_files.sort(key=get_windows_sort_key)
                            
                            images = []
                            for f in img_files:
                                try:
                                    img = PIL.Image.open(os.path.join(root, f))
                                    if img.mode != 'RGB':
                                        img = img.convert('RGB')
                                    images.append((f, img))
                                except Exception:
                                    pass
                            
                            if not images:
                                continue
                            
                            stitched_chunks = []
                            current_chunk = []
                            current_h = 0
                            current_w = 0
                            MAX_HEIGHT = 15000
                            
                            for f, img in images:
                                if current_h + img.height > MAX_HEIGHT and current_chunk:
                                    stitched_chunks.append((current_chunk, current_w, current_h))
                                    current_chunk = [(f, img)]
                                    current_h = img.height
                                    current_w = img.width
                                else:
                                    current_chunk.append((f, img))
                                    current_h += img.height
                                    current_w = max(current_w, img.width)
                                    
                            if current_chunk:
                                stitched_chunks.append((current_chunk, current_w, current_h))
                                
                            # Create new images and delete old ones
                            for i, (chunk, w, h) in enumerate(stitched_chunks):
                                canvas = PIL.Image.new('RGB', (w, h), (255, 255, 255))
                                y_offset = 0
                                for f, img in chunk:
                                    x_offset = (w - img.width) // 2  # Center align
                                    canvas.paste(img, (x_offset, y_offset))
                                    y_offset += img.height
                                
                                save_path = os.path.join(root, f"stitched_{i+1:03d}.jpg")
                                canvas.save(save_path, 'JPEG', quality=90)
                                
                            for f, img in images:
                                img.close()
                                try:
                                    os.remove(os.path.join(root, f))
                                except Exception:
                                    pass

                    # Count downloaded/stitched images
                    total_files = 0
                    for root, dirs, files in os.walk(tmp_dir):
                        total_files += len([f for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg', '.webp', '.gif'))])
                        
                    if total_files > 0:
                        st.success(f"✅ Đã tải về thành công **{total_files}** ảnh. Đang nén thành Zip...")
                        
                        # Zip the directory
                        shutil.make_archive(zip_path, 'zip', tmp_dir)
                        
                        with open(f"{zip_path}.zip", "rb") as f:
                            zip_data = f.read()
                            
                        mb_size = f"{len(zip_data) / (1024 * 1024):.2f}"
                        st.download_button(
                            label=f"⬇️ TẢI FILE ZIP ({mb_size} MB)",
                            data=zip_data,
                            file_name=f"truyen_{int(time.time())}.zip",
                            mime="application/zip",
                            type="primary",
                            use_container_width=True
                        )
                    else:
                        st.error("❌ Không tải được ảnh nào. Có thể trang web này không được hỗ trợ hoặc cần đăng nhập (Trả phí).")
                        if res.stderr or res.stdout:
                            with st.expander("Xem bảng log hệ thống"):
                                st.code(res.stderr + "\n" + res.stdout)
                                
                except Exception as e:
                    st.error(f"❌ Có lỗi hệ thống bất ngờ xảy ra: {e}")
                finally:
                    # Cleanup generated files from server disk
                    shutil.rmtree(tmp_dir, ignore_errors=True)
                    try:
                        os.remove(f"{zip_path}.zip")
                    except Exception:
                        pass

# =================== TAB 7: GLOSSARY ===================
with tab_glossary:
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
