# ============================================================
# STM Image to Excel Parser — Streamlit Edition
# KBANK + KRUNGSRI + BBL + SCB
# รัน: streamlit run app.py
# ============================================================

import os
import re
import json
import math
import tempfile
import io
import streamlit as st
import pandas as pd
from datetime import datetime
from PIL import Image, ImageEnhance, ImageFilter, ImageOps

# ─── Page config ────────────────────────────────────────────
st.set_page_config(
    page_title="STM Parser",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Global ── */
[data-testid="stAppViewContainer"] { background: #0F1117; }
[data-testid="stSidebar"] { background: #1A1E2A; border-right: 1px solid #2D3347; }
[data-testid="stSidebar"] * { color: #E8ECF4 !important; }

/* ── Header bar ── */
.stApp > header { background: transparent !important; }

/* ── Metric cards ── */
[data-testid="metric-container"] {
    background: #1A1E2A;
    border: 1px solid #2D3347;
    border-radius: 10px;
    padding: 14px 18px !important;
}
[data-testid="stMetricValue"] { font-size: 1.4rem !important; }

/* ── Dataframe ── */
[data-testid="stDataFrame"] { border: 1px solid #2D3347; border-radius: 8px; }

/* ── Upload box ── */
[data-testid="stFileUploader"] {
    border: 2px dashed #2D3347;
    border-radius: 10px;
    padding: 8px;
}

/* ── Buttons ── */
.stButton > button {
    border-radius: 8px !important;
    font-weight: 500 !important;
    transition: opacity .15s !important;
}
.stButton > button:hover { opacity: .85 !important; }

/* ── Alert boxes ── */
.stSuccess, .stError, .stWarning, .stInfo {
    border-radius: 8px !important;
    font-size: 13px !important;
}

/* ── Badge pills ── */
.badge-ok    { background:#16a34a22; color:#4ade80; border:1px solid #16a34a44; padding:2px 8px; border-radius:4px; font-size:11px; font-weight:600; }
.badge-rev   { background:#f59e0b22; color:#fbbf24; border:1px solid #f59e0b44; padding:2px 8px; border-radius:4px; font-size:11px; font-weight:600; }
.badge-bank  { background:#3b82f622; color:#60a5fa; border:1px solid #3b82f644; padding:4px 14px; border-radius:20px; font-size:13px; font-weight:600; }

/* ── Section title ── */
.section-title {
    font-size: 11px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 1px; color: #5A6278; margin-bottom: 8px;
}
/* ── Login card ── */
.login-card {
    max-width: 400px; margin: 80px auto;
    background: #1A1E2A; border: 1px solid #2D3347;
    border-radius: 14px; padding: 36px;
    text-align: center;
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# 0)  Constants & Helpers
# ============================================================
FORCE_YEAR          = datetime.now().year
BALANCE_TOLERANCE   = 0.05
LOW_CONF_THRESHOLD  = 0.75

BANK_CONFIGS = {
    "DEFAULT": {
        "bank_keywords": [],
        "opening_keywords": ["ยอด ยก มา","ยอดยกมา","ยอด ยก","ยก มา","BALANCE BROUGHT FORWARD","BROUGHT FORWARD","BALANCE B/F","B/F"],
        "credit_keywords": ["รับ โอน","รับโอน","รับ โอน เงิน","ฝาก","ฝาก เงิน","ฝาก เงินสด","ฝากเงินสด","ดอกเบี้ย","คืนเงิน","เงิน เข้า","เงินเข้า","CREDIT","DEPOSIT","TRANSFER IN"],
        "debit_keywords":  ["โอน เงิน","โอนเงิน","โอน เงิน พร้อม เพ ย์","โอน เงิน พร้อมเพย์","จ่าย บิล","จ่าย คิว อา ร์","จ่าย QR","ชำระ","ชา ระ","ชํา ระ","ถอน","ถอนเงิน","หัก","ค่าธรรมเนียม","จ่าย","DEBIT","WITHDRAW","WITHDRAWAL","PAYMENT","FEE","TRANSFER OUT"],
    },
    "KBANK": {
        "bank_keywords": ["KASIKORN","KBANK","K PLUS","กสิกร","ธนาคาร กสิกร ไทย","ธนาคารกสิกรไทย"],
        "opening_keywords": ["ยอด ยก มา","ยอดยกมา","ยอด ยก","ยก มา","BALANCE BROUGHT FORWARD","BROUGHT FORWARD"],
        "credit_keywords": ["รับ โอน","รับโอน","ฝาก","ฝาก เงินสด","เงิน เข้า","เงินเข้า","ดอกเบี้ย","CREDIT","DEPOSIT"],
        "debit_keywords":  ["โอน เงิน","โอนเงิน","โอน เงิน พร้อม เพ ย์","โอน เงิน พร้อมเพย์","จ่าย คิว อา ร์","จ่าย QR","ถอน","ถอนเงิน","หัก","ชำระ","ค่าธรรมเนียม","DEBIT","WITHDRAW","PAYMENT","FEE"],
    },
    "KRUNGSRI": {
        "bank_keywords": ["KRUNGSRI","กรุงศรี","กรุง ศรี","BANK OF AYUDHYA","AYUDHYA","MOBILE KRUNGSRI","KRUNGSRI VISA","FIRSTCHOICE","FIRST CHOICE"],
        "opening_keywords": ["ยอด ยก มา","ยอดยกมา","ยอด ยก","ยก มา","BALANCE BROUGHT FORWARD","BROUGHT FORWARD","BALANCE B/F","B/F"],
        "credit_keywords": ["รับ โอน เงิน","รับ โอน","รับโอน","ฝาก เงิน","ฝาก","ดอกเบี้ย เงิน ฝาก","ดอกเบี้ย","คืนเงิน","เงิน เข้า","เงินเข้า","CREDIT","DEPOSIT","TRANSFER IN"],
        "debit_keywords":  ["จ่าย บิล","จ่าย คิว อา ร์","จ่าย QR","โอน เงิน พร้อม เพ ย์","โอน เงิน พร้อมเพย์","โอน เงิน","โอนเงิน","ถอนเงิน","ถอน","หัก","ค่าธรรมเนียม","ชำระ","จ่าย","DEBIT","WITHDRAW","WITHDRAWAL","PAYMENT","FEE"],
    },
    "BBL": {
        "bank_keywords": ["BANGKOK BANK","BANGKOKBANK","ธนาคารกรุงเทพ","ธนาคาร กรุงเทพ","บัวหลวง","BUALUANG","STATEMENT OF SAVING ACCOUNT"],
        "opening_keywords": ["B/F","B / F","BALANCE B/F","BALANCE BROUGHT FORWARD","BROUGHT FORWARD","ยอด ยก มา","ยอดยกมา","ยอด ยก","ยก มา"],
        "credit_keywords": ["TRF FR OTH BK","TRF FROTH BK","TRF FROM OTH BK","TRF FROM OTHER BANK","TRANSFER IN","SALARY","SMART","CREDIT","DEPOSIT","ฝาก","รับ โอน","รับโอน","เงิน เข้า","เงินเข้า"],
        "debit_keywords":  ["TRF TO OTH BK","TRF TOOTH BK","TRF TO OTHER BANK","TRF. PROMPTPAY","TRF . PROMPTPAY","PMT. PROMPTPAY","PMT . PROMPTPAY","PMT.PROMPTPAY","PMT FOR GOODS","CASH W/D ATM","CASH W / D ATM","WITHDRAWAL","WITHDRAW","TRANSFER","PAYMENT","DEBIT","FEE","ถอน","ถอนเงิน","โอน เงิน","โอนเงิน","ชำระ","จ่าย"],
    },
    "SCB": {
        "bank_keywords": ["SCB","ไทยพาณิชย์","ไทย พาณิชย์","ธนาคารไทยพาณิชย์","ธนาคาร ไทย พาณิชย์","THE SIAM COMMERCIAL BANK","SIAM COMMERCIAL BANK"],
        "opening_keywords": ["ยอดเงินคงเหลือยกมา","ยอด เงิน คงเหลือ ยก มา","BALANCE BROUGHT FORWARD","BROUGHT FORWARD","BALANCE B/F","B/F"],
        "credit_keywords": [" X1 ","X1","รับ โอน จาก","รับโอนจาก","รับ โอน","รับโอน","โอน จาก","โอนจาก","ฝาก","เงิน เข้า","เงินเข้า","CREDIT","DEPOSIT","TRANSFER IN"],
        "debit_keywords":  [" X2 ","X2","โอน ไป","โอนไป","PromptPay","PROMPTPAY","จ่าย บิล","จ่ายบิล","จ่าย","ถอน","ถอนเงิน","ชำระ","ชา ระ","ชํา ระ","PAYMENT","WITHDRAW","WITHDRAWAL","TRANSFER OUT","FEE"],
    },
}

NOISE_KEYWORDS = [
    "STATEMENT OF SAVING ACCOUNT","THE SIAM COMMERCIAL BANK","ACCOUNT NO","เลขที่บัญชี",
    "ชื่อ - สกุล","ADDRESS","สาขา","BRANCH","DATE | TIME","DEBIT/CREDIT","BALANCE/BAHT",
    "DESCRIPTION/NOTE","TOTAL AMOUNTS","TOTAL ITEMS","THIS DOCUMENT IS AUTO-GENERATED",
    "เอกสาร ฉบับ นี้","ออก โดย ระบบ อัตโนมัติ",
]

BANK_LABELS = {
    "AUTO":    "🔍 ตรวจจับอัตโนมัติ",
    "KBANK":   "🟢 กสิกรไทย (KBANK)",
    "KRUNGSRI":"🟡 กรุงศรี (Krungsri)",
    "BBL":     "🔵 กรุงเทพ (BBL)",
    "SCB":     "🟣 ไทยพาณิชย์ (SCB)",
}


# ============================================================
# 1)  Basic Helpers
# ============================================================
def clean_text(s: str) -> str:
    s = str(s).replace("\n", " ")
    return re.sub(r"\s+", " ", s).strip()

def get_config(bank: str) -> dict:
    return BANK_CONFIGS.get(bank, BANK_CONFIGS["DEFAULT"])

def normalize_amount_text(s: str) -> str:
    s = str(s).strip()
    s = re.sub(r"\s+", "", s)
    ocr_map = {"O":"0","o":"0","I":"1","l":"1","|":"1","S":"5","s":"5","B":"8","g":"9","q":"9","Z":"2","z":"2"}
    for c, d in ocr_map.items():
        s = s.replace(c, d)
    if re.match(r"^\d{1,3}(,\d{3})+,\d{2}$", s):
        s = s[:-3] + "." + s[-2:]
    if re.match(r"^\d+,\d{2}$", s):
        s = s.replace(",", ".")
    if re.match(r"^\d{1,3}\.\d{3}\.\d{2}$", s):
        parts = s.split(".")
        s = parts[0] + "," + parts[1] + "." + parts[2]
    return s

def parse_money(s) -> float | None:
    if s is None or (isinstance(s, float) and math.isnan(s)):
        return None
    s = normalize_amount_text(str(s)).replace(" ", "")
    s = re.sub(r"[^0-9,.\-]", "", s)
    if not s:
        return None
    try:
        return float(s.replace(",", ""))
    except:
        return None

def is_money_token(s: str) -> bool:
    s = normalize_amount_text(str(s).strip())
    pats = [
        r"^-?\d{1,3}(,\d{3})*(\.\d{2})$",
        r"^-?\d+\.\d{2}$",
        r"^-?\d+,\d{2}$",
        r"^-?\d{1,3}\.\d{3}\.\d{2}$",
    ]
    return any(re.match(p, s) for p in pats)

def extract_money_loose(token: str) -> list[float]:
    s = str(token)
    ocr_map = {"O":"0","o":"0","I":"1","l":"1","|":"1","S":"5","s":"5","B":"8","g":"9","q":"9","Z":"2","z":"2"}
    for c, d in ocr_map.items():
        s = s.replace(c, d)
    priority_patterns = [
        r"(?<![\d,])-?\d{1,3}(?:,\d{3})+\.\d{2}(?!\d)",
        r"(?<![\d.])-?\d{1,3}(?:\.\d{3})+,\d{2}(?!\d)",
        r"(?<![\d.])-?\d{1,3}(?:\.\d{3})+\.\d{2}(?!\d)",
        r"(?<![\d,])-?\d+\.\d{2}(?!\d)",
        r"(?<![\d,])-?\d+,\d{2}(?![\d.])",
    ]
    for pat in priority_patterns:
        matches = re.findall(pat, s)
        if matches:
            vals = [v for m in matches for v in [parse_money(m)] if v is not None]
            if vals:
                return vals
    return []

def money_to_key(value) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    try:
        return str(int(round(float(value) * 100)))
    except:
        return ""

def is_valid_number(x) -> bool:
    if x is None:
        return False
    try:
        if math.isnan(float(x)):
            return False
        return True
    except:
        return False


# ============================================================
# 2)  Date Helpers
# ============================================================
def fix_ocr_date(s: str) -> str:
    s = str(s).strip()
    s = s.replace("O","0").replace("o","0").replace("I","1").replace("l","1")
    s = s.replace(".","-").replace("/","-")
    return s

def is_date_token(s: str) -> bool:
    s = fix_ocr_date(s)
    return bool(re.match(r"^\d{1,2}[-]\d{1,2}[-]\d{2,4}$", s) or
                re.match(r"^\d{1,2}[-]\d{1,2}$", s))

def normalize_date(s: str) -> str | None:
    s = fix_ocr_date(s)
    m1 = re.match(r"^(\d{1,2})[-](\d{1,2})[-](\d{2,4})$", s)
    m2 = re.match(r"^(\d{1,2})[-](\d{1,2})$", s)
    if m1:
        day, month = int(m1.group(1)), int(m1.group(2))
    elif m2:
        day, month = int(m2.group(1)), int(m2.group(2))
    else:
        return None
    try:
        dt = datetime(FORCE_YEAR, month, day)
    except:
        return None
    return dt.strftime("%d-%m-%Y")

def extract_date(line_text: str) -> str | None:
    for tok in str(line_text).split():
        d = normalize_date(tok)
        if d:
            return d
    return None

def is_time_token(s: str) -> bool:
    return bool(re.match(r"^\d{1,2}:\d{2}(:\d{2})?$", str(s).strip()))

def extract_time(line_text: str) -> str | None:
    for tok in str(line_text).split():
        if is_time_token(tok):
            return tok
    return None

def parse_date_for_sort(date_text: str):
    if not date_text or pd.isna(date_text):
        return pd.NaT
    return pd.to_datetime(str(date_text).strip().replace("/","-"), format="%d-%m-%Y", errors="coerce")


# ============================================================
# 3)  Text / Transaction Helpers
# ============================================================
def has_opening_text(text: str, bank: str = "DEFAULT") -> bool:
    up = clean_text(text).upper()
    for kw in get_config(bank).get("opening_keywords", []):
        if kw.upper() in up:
            return True
    for kw in BANK_CONFIGS["DEFAULT"]["opening_keywords"]:
        if kw.upper() in up:
            return True
    return False

def classify_kw(text: str, bank: str = "DEFAULT") -> str:
    up = clean_text(text).upper()
    for kw in get_config(bank).get("credit_keywords", []):
        if kw.upper() in up: return "credit"
    for kw in get_config(bank).get("debit_keywords", []):
        if kw.upper() in up: return "debit"
    for kw in BANK_CONFIGS["DEFAULT"]["credit_keywords"]:
        if kw.upper() in up: return "credit"
    for kw in BANK_CONFIGS["DEFAULT"]["debit_keywords"]:
        if kw.upper() in up: return "debit"
    return "unknown"

def is_noise_line(text: str) -> bool:
    up = clean_text(text).upper()
    for kw in NOISE_KEYWORDS:
        if kw.upper() in up:
            return True
    if re.search(r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\s*[-–]\s*\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", up):
        return True
    if re.search(r"หน้า\s*\d+\s*/\s*\d+", up):
        return True
    return False

def detect_bank(full_text: str) -> str:
    text = clean_text(full_text).upper()
    best, best_score = "DEFAULT", 0
    for bank_name, cfg in BANK_CONFIGS.items():
        if bank_name == "DEFAULT":
            continue
        score = 0
        for kw in cfg.get("bank_keywords", []):
            if kw.upper() in text: score += 10
        for kw in cfg.get("debit_keywords", []):
            if kw.upper() in text: score += 1
        for kw in cfg.get("credit_keywords", []):
            if kw.upper() in text: score += 1
        if score > best_score:
            best_score = score
            best = bank_name
    return best


# ============================================================
# 4)  Healing Engine
# ============================================================
def try_healing_amount(prev_balance, read_amount, balance, tolerance=0.05):
    if any(v is None or (isinstance(v, float) and math.isnan(v))
           for v in [prev_balance, read_amount, balance]):
        return None, None
    exp_debit  = round(prev_balance - balance, 2)
    exp_credit = round(balance - prev_balance, 2)
    exp_amount = exp_debit if exp_debit > 0 else exp_credit
    tx_type    = "debit" if exp_debit > 0 else "credit"
    str_read, str_exp = f"{read_amount:.2f}", f"{exp_amount:.2f}"
    if abs(read_amount / 100 - exp_amount) <= tolerance:
        return exp_amount, f"healed_missing_dot_{tx_type}"
    if len(str_read) == len(str_exp):
        mismatches = sum(1 for a, b in zip(str_read, str_exp) if a != b)
        if mismatches <= 1:
            return exp_amount, f"healed_swapped_digit_{tx_type}"
    if len(str_exp) > len(str_read) and (str_read in str_exp or str_exp.endswith(str_read)):
        return exp_amount, f"healed_omitted_digit_{tx_type}"
    return None, None

def suggest_from_balance(prev_balance, balance):
    if any(v is None or (isinstance(v, float) and math.isnan(v)) for v in [prev_balance, balance]):
        return None, None
    diff = round(balance - prev_balance, 2)
    if diff > 0: return abs(diff), "credit"
    if diff < 0: return abs(diff), "debit"
    return 0.0, "zero"

def amount_balance_match(prev_balance, amount, balance, tolerance=BALANCE_TOLERANCE):
    if any(v is None or (isinstance(v, float) and math.isnan(v)) for v in [prev_balance, amount, balance]):
        return None
    debit_exp  = round(prev_balance - amount, 2)
    credit_exp = round(prev_balance + amount, 2)
    if abs(balance - debit_exp) <= tolerance:
        return {"type": "debit",   "expected_balance": debit_exp,  "diff": round(balance - debit_exp, 2),  "balance_check": "OK_DEBIT"}
    if abs(balance - credit_exp) <= tolerance:
        return {"type": "credit",  "expected_balance": credit_exp, "diff": round(balance - credit_exp, 2), "balance_check": "OK_CREDIT"}
    return None


# ============================================================
# 5)  Image Pre-processing
# ============================================================
def preprocess_image_variants(pil_img: Image.Image, base_name: str) -> list[tuple[str, Image.Image]]:
    img = pil_img.convert("RGB")
    variants = [("original", img)]

    # 2x + contrast + sharpen
    i1 = img.resize((img.width*2, img.height*2), Image.LANCZOS)
    i1 = ImageEnhance.Contrast(i1).enhance(1.9)
    i1 = ImageEnhance.Sharpness(i1).enhance(2.2)
    variants.append(("2x_contrast_sharp", i1))

    # 3x + grayscale + autocontrast
    i2 = img.resize((img.width*3, img.height*3), Image.LANCZOS)
    i2 = i2.convert("L"); i2 = ImageOps.autocontrast(i2)
    i2 = ImageEnhance.Contrast(i2).enhance(1.8)
    i2 = ImageEnhance.Sharpness(i2).enhance(2.0)
    variants.append(("3x_gray_autocontrast", i2))

    # 3x + UnsharpMask
    i3 = img.resize((img.width*3, img.height*3), Image.LANCZOS)
    i3 = i3.convert("L"); i3 = ImageOps.autocontrast(i3)
    i3 = i3.filter(ImageFilter.UnsharpMask(radius=1.2, percent=180, threshold=3))
    variants.append(("3x_unsharp", i3))

    # 3x BW threshold 175
    i4 = img.resize((img.width*3, img.height*3), Image.LANCZOS)
    i4 = i4.convert("L"); i4 = ImageOps.autocontrast(i4)
    i4 = ImageEnhance.Contrast(i4).enhance(2.0)
    i4 = i4.point(lambda x: 255 if x > 175 else 0)
    variants.append(("3x_bw_175", i4))

    # 3x BW threshold 145
    i5 = img.resize((img.width*3, img.height*3), Image.LANCZOS)
    i5 = i5.convert("L"); i5 = ImageOps.autocontrast(i5)
    i5 = ImageEnhance.Contrast(i5).enhance(1.7)
    i5 = i5.point(lambda x: 255 if x > 145 else 0)
    variants.append(("3x_bw_145", i5))

    return variants


# ============================================================
# 6)  Google Vision OCR
# ============================================================
def get_vision_client():
    """สร้าง Vision client จาก secret ใน .streamlit/secrets.toml"""
    from google.cloud import vision
    from google.oauth2 import service_account

    key_json_str = st.secrets.get("GOOGLE_VISION_KEY", "")
    if not key_json_str:
        st.error("❌ ไม่พบ GOOGLE_VISION_KEY ใน Streamlit Secrets")
        st.stop()

    try:
        info = json.loads(key_json_str)
    except json.JSONDecodeError as e:
        st.error(f"❌ GOOGLE_VISION_KEY ไม่ใช่ JSON ที่ถูกต้อง: {e}")
        st.stop()

    creds = service_account.Credentials.from_service_account_info(
        info,
        scopes=["https://www.googleapis.com/auth/cloud-platform"],
    )
    return vision.ImageAnnotatorClient(credentials=creds)


def vision_ocr_pil(pil_img: Image.Image, client, file_name: str) -> tuple[str, pd.DataFrame]:
    from google.cloud import vision as vis
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    content = buf.getvalue()

    image = vis.Image(content=content)
    response = client.document_text_detection(
        image=image,
        image_context=vis.ImageContext(language_hints=["th", "en"]),
    )
    if response.error.message:
        raise RuntimeError(response.error.message)

    rows = []
    for page in response.full_text_annotation.pages:
        for block in page.blocks:
            for para in block.paragraphs:
                for word in para.words:
                    word_text = "".join(s.text for s in word.symbols)
                    verts = word.bounding_box.vertices
                    xs = [v.x for v in verts]; ys = [v.y for v in verts]
                    xmin,xmax = min(xs),max(xs); ymin,ymax = min(ys),max(ys)
                    rows.append({
                        "file_name": file_name,
                        "page_width": page.width, "page_height": page.height,
                        "text": word_text,
                        "word_confidence": word.confidence,
                        "x_min": xmin, "y_min": ymin, "x_max": xmax, "y_max": ymax,
                        "x_center": (xmin+xmax)/2, "y_center": (ymin+ymax)/2,
                        "width": xmax-xmin, "height": ymax-ymin,
                    })
    return response.full_text_annotation.text, pd.DataFrame(rows)


# ============================================================
# 7)  Line Rebuild + Parse
# ============================================================
def rebuild_lines_by_y(words_df: pd.DataFrame, y_threshold: int = 10) -> pd.DataFrame:
    all_lines = []
    for file_name, fg in words_df.groupby("file_name"):
        if fg.empty: continue
        page_width = fg["page_width"].dropna().iloc[0]
        fg = fg[fg["x_center"] >= page_width * 0.04].copy()
        if fg.empty: continue
        fg = fg.sort_values(["y_center","x_center"]).reset_index(drop=True)
        current_line, current_y, line_id = [], None, 0
        for _, row in fg.iterrows():
            y = row["y_center"]
            if current_y is None:
                current_line, current_y = [row], y; continue
            if abs(y - current_y) <= y_threshold:
                current_line.append(row)
                current_y = (current_y + y) / 2
            else:
                ldf = pd.DataFrame(current_line); ldf["line_id"] = line_id
                all_lines.append(ldf); line_id += 1
                current_line, current_y = [row], y
        if current_line:
            ldf = pd.DataFrame(current_line); ldf["line_id"] = line_id
            all_lines.append(ldf)
    return pd.concat(all_lines, ignore_index=True) if all_lines else pd.DataFrame()


def parse_lines(words_df: pd.DataFrame, active_bank: str) -> pd.DataFrame:
    line_words = rebuild_lines_by_y(words_df, y_threshold=10)
    if line_words.empty: return pd.DataFrame()
    line_rows = []
    for (file_name, line_id), g in line_words.groupby(["file_name","line_id"]):
        g = g.sort_values("x_center").copy()
        raw = " ".join(g["text"].astype(str).tolist()).strip()
        raw = re.sub(r"\s+", " ", raw)
        if is_noise_line(raw): continue
        date = extract_date(raw)
        time = extract_time(raw)
        is_opening = has_opening_text(raw, active_bank)
        if not date and not is_opening: continue
        # collect money values
        money_vals = []
        for _, row in g.sort_values("x_center").iterrows():
            txt = normalize_amount_text(str(row["text"]))
            if is_money_token(txt) or extract_money_loose(txt):
                v = parse_money(txt)
                if v is None:
                    vs = extract_money_loose(txt)
                    if vs: v = vs[0]
                if v is not None:
                    money_vals.append(v)
        amount, balance = None, None
        if is_opening:
            if money_vals: balance = money_vals[-1]
        else:
            if len(money_vals) >= 2: amount, balance = money_vals[0], money_vals[-1]
            elif len(money_vals) == 1: amount = money_vals[0]
        avg_conf = float(g["word_confidence"].mean()) if "word_confidence" in g.columns else None
        min_conf = float(g["word_confidence"].min()) if "word_confidence" in g.columns else None
        line_rows.append({
            "file_name": file_name, "line_id": int(line_id),
            "date": date, "time": time, "code": None,
            "amount": amount, "balance": balance,
            "money_tokens": " | ".join(str(v) for v in money_vals),
            "money_count": len(money_vals),
            "raw_line_text": raw,
            "y_center": float(g["y_center"].mean()),
            "parsed_date": parse_date_for_sort(date),
            "is_opening_balance": is_opening,
            "line_confidence": avg_conf,
            "min_confidence": min_conf,
        })
    if not line_rows: return pd.DataFrame()
    df = pd.DataFrame(line_rows)
    return df.sort_values(["file_name","y_center"], na_position="last").reset_index(drop=True)


# ============================================================
# 8)  Dedup + Chain Check
# ============================================================
def build_duplicate_key(row) -> str:
    time_key    = str(row.get("time", "") or "").strip()
    amount_key  = money_to_key(row.get("amount"))
    balance_key = money_to_key(row.get("balance"))
    raw         = str(row.get("raw_line_text", ""))
    if has_opening_text(raw):
        return "OPENING|" + balance_key
    if time_key:
        return f"TX|{time_key}|{amount_key}|{balance_key}"
    return f"TX|{amount_key}|{balance_key}"

def deduplicate(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return df
    df = df.copy()
    df["_dup_key"] = df.apply(build_duplicate_key, axis=1)
    df = df.drop_duplicates(subset=["_dup_key"], keep="first").reset_index(drop=True)
    return df.drop(columns=["_dup_key"], errors="ignore")


def add_debit_credit_check(ordered_df: pd.DataFrame, active_bank: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    check_rows = []
    prev_balance = None
    seq = 0
    for _, row in ordered_df.iterrows():
        seq += 1
        date         = row.get("date")
        time         = row.get("time")
        code         = row.get("code")
        amount       = row.get("amount")
        bal_val      = row.get("balance")
        raw          = row.get("raw_line_text", "")
        money_tokens = row.get("money_tokens", "")
        line_conf    = row.get("line_confidence")
        min_conf     = row.get("min_confidence")
        order_method = row.get("order_method", "")
        is_opening   = row.get("is_opening_balance", False)

        debit = credit = expected_balance = diff_ = suggested_amount = suggested_type = None
        balance_check = ""

        low_conf = (
            (line_conf is not None and not math.isnan(line_conf) and line_conf < LOW_CONF_THRESHOLD) or
            (min_conf  is not None and not math.isnan(min_conf)  and min_conf  < LOW_CONF_THRESHOLD)
        )

        if is_opening or code == "OPENING":
            balance_check = "OPENING_BALANCE"
        else:
            match = amount_balance_match(prev_balance, amount, bal_val)
            if match:
                expected_balance = match["expected_balance"]
                diff_            = match["diff"]
                balance_check    = match["balance_check"]
                if match["type"] == "debit": debit  = amount
                else:                        credit = amount
            else:
                h_amt, h_type = try_healing_amount(prev_balance, amount, bal_val)
                if h_amt is not None:
                    if "debit"  in h_type: debit  = h_amt
                    else:                  credit = h_amt
                    expected_balance = bal_val; diff_ = 0.0
                    balance_check = f"OK_{h_type.upper()}"
                else:
                    tx_type = classify_kw(raw, active_bank)
                    s_amt, s_type = suggest_from_balance(prev_balance, bal_val)
                    if s_amt is not None and s_type in ("debit","credit"):
                        suggested_amount, suggested_type = s_amt, s_type
                        expected_balance = round(
                            (prev_balance - s_amt) if s_type=="debit" else (prev_balance + s_amt), 2
                        ) if prev_balance is not None else None
                        diff_ = round(bal_val - expected_balance, 2) if expected_balance is not None else None
                        if tx_type == s_type:
                            balance_check = "SUGGEST_AMOUNT_BY_KEYWORD_REVIEW"
                            if s_type == "debit": debit  = s_amt
                            else:                 credit = s_amt
                        else:
                            balance_check = "SUGGEST_AMOUNT_REVIEW"
                    else:
                        balance_check = "CHAIN_BROKEN_OCR_REVIEW"

        if low_conf and "OK" in balance_check:
            balance_check += "_LOW_CONFIDENCE"

        check_rows.append({
            "seq": seq, "date": date, "time": time, "code": code,
            "debit": debit, "credit": credit, "balance": bal_val,
            "prev_balance": prev_balance, "amount": amount,
            "suggested_amount": suggested_amount, "suggested_type": suggested_type,
            "expected_balance": expected_balance, "diff": diff_,
            "line_confidence": line_conf, "min_confidence": min_conf,
            "money_tokens": money_tokens, "raw_line_text": raw,
            "order_method": order_method, "balance_check": balance_check,
        })
        if bal_val is not None and not (isinstance(bal_val, float) and math.isnan(bal_val)):
            prev_balance = bal_val

    check_df = pd.DataFrame(check_rows)
    parsed_df = check_df[
        (check_df["balance_check"] != "OPENING_BALANCE") &
        (check_df["balance"].apply(is_valid_number)) &
        (check_df["debit"].apply(is_valid_number) | check_df["credit"].apply(is_valid_number))
    ][["date","time","debit","credit","balance","balance_check","raw_line_text"]].copy()
    parsed_df = parsed_df.reset_index(drop=True)
    return parsed_df, check_df


# ============================================================
# 9)  Score + Best Variant
# ============================================================
def score_words(words_df: pd.DataFrame, active_bank: str = "DEFAULT") -> float:
    if words_df.empty: return -999999.0
    parsed = parse_lines(words_df, active_bank)
    if parsed.empty: return 0.0
    conf = parsed["line_confidence"].fillna(0).mean() * 100 if "line_confidence" in parsed.columns else 0
    return (
        len(parsed) * 100 +
        parsed["money_count"].fillna(0).sum() * 20 +
        parsed["date"].notna().sum() * 50 +
        conf
    )


# ============================================================
# 10)  Full Pipeline
# ============================================================
def run_pipeline(
    uploaded_files,
    selected_bank: str,
    progress_cb=None,
    status_cb=None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, str]:
    """
    Returns: (parsed_df, check_df, review_df, active_bank)
    """
    client = get_vision_client()
    all_words_list, all_full_texts = [], []
    n = len(uploaded_files)

    for i, uf in enumerate(uploaded_files):
        fname = uf.name
        if status_cb: status_cb(f"📡 OCR ภาพที่ {i+1}/{n}: {fname}")
        if progress_cb: progress_cb((i) / n * 0.6)

        pil_img = Image.open(uf).convert("RGB")
        variants = preprocess_image_variants(pil_img, fname)

        best_score, best_words = -999999, None

        for v_name, v_img in variants:
            try:
                full_text, word_df = vision_ocr_pil(v_img, client, fname)
                if word_df.empty: continue
                score = score_words(word_df, selected_bank if selected_bank != "AUTO" else "DEFAULT")
                if score > best_score:
                    best_score = score
                    best_words = word_df
                    best_text  = full_text
            except Exception as e:
                pass  # skip failed variant

        if best_words is None:
            raise RuntimeError(f"OCR ล้มเหลวทุก variant สำหรับ {fname}")

        all_words_list.append(best_words)
        all_full_texts.append(best_text)

    if progress_cb: progress_cb(0.65)
    if status_cb: status_cb("🔍 ตรวจจับธนาคาร + Parse รายการ...")

    combined_text = "\n".join(all_full_texts)
    active_bank   = selected_bank if selected_bank != "AUTO" else detect_bank(combined_text)

    words_df = pd.concat(all_words_list, ignore_index=True)

    if progress_cb: progress_cb(0.75)
    parsed_lines = parse_lines(words_df, active_bank)
    dedup_lines  = deduplicate(parsed_lines)
    dedup_lines  = dedup_lines.sort_values(["file_name","y_center"], na_position="last").reset_index(drop=True)

    if progress_cb: progress_cb(0.9)
    if status_cb: status_cb("✅ คำนวณ Balance Chain...")

    parsed_df, check_df = add_debit_credit_check(dedup_lines, active_bank)

    review_df = check_df[
        check_df["balance_check"].astype(str).str.contains(
            "REVIEW|BROKEN|UNKNOWN|NO_PREV|MISSING|SUGGEST|HEALED", regex=True
        )
    ].copy().reset_index(drop=True)

    if progress_cb: progress_cb(1.0)
    return parsed_df, check_df, review_df, active_bank


# ============================================================
# 11)  Excel Export
# ============================================================
def create_excel(parsed_df: pd.DataFrame, check_df: pd.DataFrame,
                 review_df: pd.DataFrame, active_bank: str) -> bytes:
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.utils import get_column_letter

    buf = io.BytesIO()

    # Build summary
    db_s  = pd.to_numeric(parsed_df.get("debit",  pd.Series(dtype=float)), errors="coerce")
    cr_s  = pd.to_numeric(parsed_df.get("credit", pd.Series(dtype=float)), errors="coerce")
    summary_df = pd.DataFrame([
        ["detected_bank",    active_bank],
        ["transaction_rows", int(parsed_df["date"].notna().sum()) if not parsed_df.empty else 0],
        ["debit_total",      float(db_s.fillna(0).sum())],
        ["credit_total",     float(cr_s.fillna(0).sum())],
        ["net_change",       float(cr_s.fillna(0).sum() - db_s.fillna(0).sum())],
        ["ok_checks",        int(check_df["balance_check"].astype(str).str.contains("OK").sum()) if not check_df.empty else 0],
        ["anomalies_found",  int(check_df["balance_check"].astype(str).str.contains("REVIEW|BROKEN|MISSING").sum()) if not check_df.empty else 0],
    ], columns=["metric","value"])

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        parsed_df.to_excel(writer, sheet_name="parsed_stm", index=False)
        check_df.to_excel(writer,  sheet_name="check",      index=False)
        summary_df.to_excel(writer,sheet_name="summary",    index=False)
        review_df.to_excel(writer, sheet_name="ocr_review", index=False)

    buf.seek(0)
    wb = load_workbook(buf)

    # Format "check" sheet
    if "check" in wb.sheetnames:
        ws = wb["check"]
        headers = {str(cell.value): cell.column for cell in ws[1] if cell.value}
        hfill = PatternFill("solid", fgColor="1E3A5F")
        hfont = Font(color="FFFFFF", bold=True)
        thin  = Side(style="thin", color="2D3347")
        for cell in ws[1]:
            cell.fill = hfill; cell.font = hfont
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(bottom=thin)
        num_cols = ["debit","credit","balance","prev_balance","amount","suggested_amount","expected_balance","diff"]
        for col in num_cols:
            if col in headers:
                for row in range(2, ws.max_row+1):
                    ws.cell(row=row, column=headers[col]).number_format = "#,##0.00"
        ws.freeze_panes = "A2"
        light_red = PatternFill("solid", fgColor="FCE4E4")
        dark_font = Font(color="9C0006")
        if "balance_check" in headers:
            bc_col = get_column_letter(headers["balance_check"])
            ws.conditional_formatting.add(
                f"A2:{get_column_letter(ws.max_column)}{ws.max_row}",
                FormulaRule(
                    formula=[f'=AND(${bc_col}2<>"OK_DEBIT",${bc_col}2<>"OK_CREDIT",${bc_col}2<>"OPENING_BALANCE",${bc_col}2<>"")'],
                    fill=light_red, font=dark_font
                )
            )
        for col_idx in range(1, ws.max_column+1):
            cl = get_column_letter(col_idx)
            max_len = max((len(str(c.value or "")) for c in ws[cl]), default=10)
            ws.column_dimensions[cl].width = min(max(max_len+2, 12), 40)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ============================================================
# 12)  Session State Init
# ============================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "results" not in st.session_state:
    st.session_state.results = None
if "uploaded_file_names" not in st.session_state:
    st.session_state.uploaded_file_names = []


# ============================================================
# 13)  LOGIN PAGE
# ============================================================
def login_page():
    st.markdown("""
    <div style="text-align:center;margin-top:60px;">
      <div style="font-size:56px;margin-bottom:16px;">🔐</div>
      <h2 style="margin-bottom:4px;">STM Image Parser</h2>
      <p style="color:#8B93A8;font-size:14px;margin-bottom:32px;">
        ป้อนรหัสผ่านเพื่อเข้าใช้งาน
      </p>
    </div>
    """, unsafe_allow_html=True)

    col_l, col_c, col_r = st.columns([1, 1.2, 1])
    with col_c:
        password = st.text_input(
            "รหัสผ่าน",
            type="password",
            placeholder="••••••••",
            label_visibility="collapsed",
        )
        if st.button("🔓  เข้าสู่ระบบ", use_container_width=True, type="primary"):
            correct_pw = st.secrets.get("APP_PASSWORD", "stm2025")
            if password == correct_pw:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ รหัสผ่านไม่ถูกต้อง")


# ============================================================
# 14)  MAIN APP PAGE
# ============================================================
def main_app():
    # ── Sidebar ──────────────────────────────────────────────
    with st.sidebar:
        st.markdown("""
        <div style="text-align:center;padding:20px 0 8px;">
          <div style="font-size:40px;">📊</div>
          <div style="font-size:18px;font-weight:700;margin-top:6px;">STM Parser</div>
          <div style="font-size:11px;color:#5A6278;margin-top:4px;">
            KBANK · KRUNGSRI · BBL · SCB
          </div>
        </div>
        <hr style="border:none;border-top:1px solid #2D3347;margin:12px 0 20px;">
        """, unsafe_allow_html=True)

        # ── Year badge ──
        st.markdown(f"""
        <div style="background:#1A1E2A;border:1px solid #2D3347;border-radius:8px;
                    padding:10px 14px;margin-bottom:16px;text-align:center;">
          <div style="font-size:10px;color:#5A6278;text-transform:uppercase;letter-spacing:1px;">
            บังคับใช้ปี (FORCE_YEAR)
          </div>
          <div style="font-size:22px;font-weight:700;color:#3B82F6;font-family:monospace;">
            {FORCE_YEAR}
          </div>
        </div>
        """, unsafe_allow_html=True)

        # ── Bank Selector ──
        st.markdown('<div class="section-title">เลือกธนาคาร</div>', unsafe_allow_html=True)
        bank_choice = st.radio(
            "bank",
            options=list(BANK_LABELS.keys()),
            format_func=lambda x: BANK_LABELS[x],
            index=0,
            label_visibility="collapsed",
        )

        st.markdown('<hr style="border:none;border-top:1px solid #2D3347;margin:16px 0;">', unsafe_allow_html=True)

        # ── Upload ──
        st.markdown('<div class="section-title">อัปโหลดภาพ Statement</div>', unsafe_allow_html=True)
        uploaded_files = st.file_uploader(
            "upload",
            type=["jpg","jpeg","png"],
            accept_multiple_files=True,
            label_visibility="collapsed",
            help="รองรับ JPG / PNG · หลายไฟล์พร้อมกันได้",
        )

        # ── Clear uploaded files ──
        if uploaded_files:
            for uf in uploaded_files:
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.markdown(f"<div style='font-size:11px;color:#8B93A8;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;'>"
                                f"📄 {uf.name}</div>", unsafe_allow_html=True)
                with col2:
                    pass  # file_uploader มีปุ่ม X ในตัวอยู่แล้ว

        st.markdown('<hr style="border:none;border-top:1px solid #2D3347;margin:16px 0;">', unsafe_allow_html=True)

        # ── Run Button ──
        run_clicked = st.button(
            "▶  เริ่มประมวลผล",
            type="primary",
            use_container_width=True,
            disabled=not bool(uploaded_files),
        )

        st.markdown('<hr style="border:none;border-top:1px solid #2D3347;margin:16px 0;">', unsafe_allow_html=True)

        # ── Logout ──
        if st.button("🚪  ออกจากระบบ", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.results = None
            st.rerun()

        st.markdown(f"""
        <div style="margin-top:20px;font-size:10px;color:#3B6278;text-align:center;line-height:1.8;">
          Google Cloud Vision OCR<br>
          KBANK · KRUNGSRI · BBL · SCB<br>
          ปีที่ใช้: {FORCE_YEAR}
        </div>
        """, unsafe_allow_html=True)

    # ── Main Content ──────────────────────────────────────────
    st.markdown(f"""
    <div style="display:flex;align-items:center;gap:14px;margin-bottom:28px;">
      <div style="font-size:36px;">📊</div>
      <div>
        <div style="font-size:24px;font-weight:700;letter-spacing:-0.3px;">STM Image Parser</div>
        <div style="font-size:13px;color:#8B93A8;margin-top:2px;">
          Google Cloud Vision OCR · ปีบังคับ: <b style="color:#3B82F6;">{FORCE_YEAR}</b>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Run Pipeline ──
    if run_clicked and uploaded_files:
        st.session_state.results = None
        progress_bar = st.progress(0, text="กำลังเริ่มต้น...")
        status_box   = st.empty()

        def prog_cb(val):
            progress_bar.progress(val, text=f"{int(val*100)}%")
        def stat_cb(msg):
            status_box.info(msg)

        try:
            parsed_df, check_df, review_df, active_bank = run_pipeline(
                uploaded_files, bank_choice, prog_cb, stat_cb
            )
            st.session_state.results = {
                "parsed": parsed_df, "check": check_df,
                "review": review_df, "active_bank": active_bank,
            }
            progress_bar.progress(1.0, text="✅ เสร็จสิ้น!")
            status_box.success(f"✅ ประมวลผลสำเร็จ · {len(parsed_df)} รายการ · ธนาคาร: {active_bank}")
        except Exception as e:
            progress_bar.empty()
            status_box.error(f"❌ Error: {e}")
            st.exception(e)

    # ── Display Results ──
    if st.session_state.results:
        res = st.session_state.results
        parsed_df  = res["parsed"]
        check_df   = res["check"]
        review_df  = res["review"]
        active_bank= res["active_bank"]

        # ── Bank badge ──
        bank_emoji = {"KBANK":"🟢","KRUNGSRI":"🟡","BBL":"🔵","SCB":"🟣","DEFAULT":"🔍"}.get(active_bank,"🔍")
        st.markdown(f'<span class="badge-bank">{bank_emoji} ตรวจพบ: {active_bank}</span>', unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # ── Summary Metrics ──
        db_s = pd.to_numeric(parsed_df.get("debit",  pd.Series(dtype=float)), errors="coerce")
        cr_s = pd.to_numeric(parsed_df.get("credit", pd.Series(dtype=float)), errors="coerce")
        net  = cr_s.fillna(0).sum() - db_s.fillna(0).sum()
        ok_n = check_df["balance_check"].astype(str).str.contains("OK").sum() if not check_df.empty else 0
        rv_n = len(review_df)

        c1,c2,c3,c4,c5,c6 = st.columns(6)
        c1.metric("📋 รายการทั้งหมด", f"{len(parsed_df)}")
        c2.metric("💚 ยอดฝากรวม", f"฿{cr_s.fillna(0).sum():,.2f}")
        c3.metric("❤️ ยอดถอนรวม",  f"฿{db_s.fillna(0).sum():,.2f}")
        c4.metric("⚖️ ยอดสุทธิ",   f"฿{net:,.2f}")
        c5.metric("✅ ผ่านตรวจสอบ", f"{ok_n}")
        c6.metric("⚠️ ต้องตรวจ",   f"{rv_n}")

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Tabs ──
        tab1, tab2, tab3, tab4 = st.tabs(["📋 parsed_stm", "🔍 check", "⚠️ OCR Review", "📥 Export"])

        # ── Tab 1: parsed_stm ──
        with tab1:
            if parsed_df.empty:
                st.warning("ไม่พบรายการธุรกรรม")
            else:
                display_df = parsed_df.copy()
                # Format numbers
                for col in ["debit","credit","balance"]:
                    if col in display_df.columns:
                        display_df[col] = pd.to_numeric(display_df[col], errors="coerce")

                def highlight_row(row):
                    check = str(row.get("balance_check",""))
                    if "REVIEW" in check or ("OK" not in check and check not in ("OPENING_BALANCE","")):
                        return ["background-color: rgba(245,158,11,0.08)"]*len(row)
                    return [""]*len(row)

                styled = display_df.style.apply(highlight_row, axis=1)
                styled = styled.format({
                    "debit":   lambda v: f"{v:,.2f}" if pd.notna(v) else "",
                    "credit":  lambda v: f"{v:,.2f}" if pd.notna(v) else "",
                    "balance": lambda v: f"{v:,.2f}" if pd.notna(v) else "",
                })
                st.dataframe(styled, use_container_width=True, height=500)

        # ── Tab 2: check ──
        with tab2:
            if check_df.empty:
                st.info("ไม่มีข้อมูล")
            else:
                show_cols = ["seq","date","time","debit","credit","balance",
                             "prev_balance","expected_balance","diff","balance_check","raw_line_text"]
                show_cols = [c for c in show_cols if c in check_df.columns]
                st.dataframe(check_df[show_cols], use_container_width=True, height=500)

        # ── Tab 3: OCR Review ──
        with tab3:
            if review_df.empty:
                st.success("✅ ไม่มีรายการที่ต้องตรวจสอบ — ทุกแถวผ่าน Balance Chain")
            else:
                st.warning(f"⚠️ พบ {len(review_df)} รายการที่ต้องตรวจสอบ")
                show_cols = ["seq","date","debit","credit","balance","balance_check","raw_line_text"]
                show_cols = [c for c in show_cols if c in review_df.columns]
                st.dataframe(review_df[show_cols], use_container_width=True, height=400)

        # ── Tab 4: Export ──
        with tab4:
            st.markdown("#### 📥 ดาวน์โหลดผลลัพธ์")
            col_a, col_b, col_c = st.columns(3)

            with col_a:
                if not parsed_df.empty:
                    csv_parsed = parsed_df.to_csv(index=False, encoding="utf-8-sig")
                    st.download_button(
                        "📄 parsed_stm.csv",
                        data=csv_parsed.encode("utf-8-sig"),
                        file_name="parsed_stm.csv",
                        mime="text/csv",
                        use_container_width=True,
                    )

            with col_b:
                if not check_df.empty:
                    csv_check = check_df.to_csv(index=False, encoding="utf-8-sig")
                    st.download_button(
                        "🔍 check.csv",
                        data=csv_check.encode("utf-8-sig"),
                        file_name="check.csv",
                        mime="text/csv",
                        use_container_width=True,
                    )

            with col_c:
                if not review_df.empty:
                    csv_rev = review_df.to_csv(index=False, encoding="utf-8-sig")
                    st.download_button(
                        "⚠️ ocr_review.csv",
                        data=csv_rev.encode("utf-8-sig"),
                        file_name="ocr_review.csv",
                        mime="text/csv",
                        use_container_width=True,
                    )

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("##### 📊 Excel (พร้อม Conditional Formatting)")
            try:
                excel_bytes = create_excel(parsed_df, check_df, review_df, active_bank)
                st.download_button(
                    "📊 ดาวน์โหลด Excel (stm_result.xlsx)",
                    data=excel_bytes,
                    file_name="stm_result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary",
                )
            except Exception as e:
                st.error(f"สร้าง Excel ไม่ได้: {e}")

    else:
        # ── Welcome State ──
        if not uploaded_files:
            st.markdown("""
            <div style="text-align:center;padding:60px 20px;color:#5A6278;">
              <div style="font-size:56px;margin-bottom:16px;">📤</div>
              <div style="font-size:16px;font-weight:600;color:#8B93A8;margin-bottom:8px;">
                อัปโหลดภาพ Statement ทางซ้าย
              </div>
              <div style="font-size:13px;">
                รองรับ JPG / PNG · หลายไฟล์พร้อมกัน<br>
                KBANK · KRUNGSRI · BBL · SCB
              </div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.info(f"📁 เลือกไฟล์แล้ว {len(uploaded_files)} ไฟล์ — กด **▶ เริ่มประมวลผล** ในแถบซ้าย")


# ============================================================
# 15)  Entry Point
# ============================================================
if not st.session_state.authenticated:
    login_page()
else:
    main_app()
