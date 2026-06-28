# ============================================================
# STM Image to Excel Parser — Streamlit Edition
# KBANK + KRUNGSRI + BBL + SCB
# รัน: streamlit run app.py
# ============================================================

import os
import re
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
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ─── CSS ────────────────────────────────────────────────────
st.markdown("""
<style>
    /* ใช้ Theme ของ Streamlit เป็นหลัก เพื่อให้เข้ากับ Light / Dark mode */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 980px;
    }

    h1 {
        font-size: 2.35rem !important;
        font-weight: 800 !important;
        letter-spacing: -0.5px;
        margin-bottom: 0.25rem;
    }

    div[data-testid="stCaptionContainer"] {
        color: var(--text-color-secondary, #8c8c8c);
        font-size: 0.95rem;
    }

    .section-title {
        display: flex;
        align-items: center;
        gap: 8px;
        font-weight: 800;
        font-size: 1.15rem;
        line-height: 1.4;
        margin: 18px 0 10px 0;
        color: inherit !important;
        background: transparent !important;
        border: none !important;
        padding: 0 !important;
    }

    .section-title .step-number {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-width: 30px;
        height: 30px;
        border-radius: 999px;
        background: #4c82fb;
        color: #ffffff;
        font-size: 0.95rem;
        font-weight: 800;
    }

    .section-title .step-text { color: inherit; }

    .badge {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        background: #e8f0fe;
        color: #1a56db;
        border-radius: 999px;
        padding: 7px 18px;
        font-size: 1.05rem;
        font-weight: 900;
        line-height: 1.2;
        margin-left: 10px;
        min-width: 68px;
        min-height: 34px;
        letter-spacing: 0.2px;
    }

    @media (prefers-color-scheme: dark) {
        .badge {
            background: rgba(76, 130, 251, 0.18);
            color: #a9c4ff;
            border: 1px solid rgba(169, 196, 255, 0.25);
        }
    }

    div[data-testid="stRadio"] label {
        font-size: 1.02rem !important;
        font-weight: 600 !important;
        padding: 4px 8px;
        border-radius: 8px;
    }

    div[data-testid="stRadio"] div[role="radiogroup"] { gap: 14px; }

    div[data-testid="stFileUploader"] { margin-top: 4px; }
    div[data-testid="stFileUploader"] section { border-radius: 10px; }

    div[data-testid="stAlert"] { border-radius: 10px; }

    div[data-testid="stButton"] button,
    div[data-testid="stDownloadButton"] button {
        border-radius: 10px;
        font-weight: 750;
    }

    div[data-testid="metric-container"] {
        border-radius: 10px;
        padding: 14px 16px !important;
    }

    div[data-testid="stDataFrame"] {
        border-radius: 10px;
        overflow: hidden;
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
    "AUTO":    "ตรวจจับอัตโนมัติ",
    "KBANK":   "กสิกรไทย (KBANK)",
    "KRUNGSRI":"กรุงศรี (Krungsri)",
    "BBL":     "กรุงเทพ (BBL)",
    "SCB":     "ไทยพาณิชย์ (SCB)",
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
    from google.cloud import vision
    from google.oauth2 import service_account

    if "gcp_service_account" not in st.secrets:
        st.error("ไม่พบ [gcp_service_account] ใน Streamlit Secrets")
        st.stop()

    try:
        info = dict(st.secrets["gcp_service_account"])
        if "private_key" in info and isinstance(info["private_key"], str):
            info["private_key"] = info["private_key"].replace("\\n", "\n")

        creds = service_account.Credentials.from_service_account_info(
            info,
            scopes=["https://www.googleapis.com/auth/cloud-platform"],
        )
    except Exception as e:
        st.error(f"อ่าน [gcp_service_account] ไม่สำเร็จ: {e}")
        st.stop()

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
    parsed = parse_lines(words_df, active
