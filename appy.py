# ============================================================
# STM Image to Excel Parser - Streamlit Version
# Manual Bank Selection: KBANK / KRUNGSRI / BBL / SCB / DEFAULT
# Google Cloud Vision OCR
# ============================================================

import os
import re
import json
import tempfile
from datetime import datetime

import pandas as pd
import streamlit as st

from PIL import Image, ImageEnhance, ImageFilter, ImageOps
from google.cloud import vision

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter


# ============================================================
# 0) Streamlit Page
# ============================================================

st.set_page_config(
    page_title="STM Image to Excel Parser",
    page_icon="📄",
    layout="wide"
)

st.title("📄 STM Image to Excel Parser")
st.caption("OCR Statement รูปภาพ → Excel | เลือกธนาคารเอง ไม่ต้อง Auto Detect")


# ============================================================
# 1) Settings
# ============================================================

BALANCE_TOLERANCE = 0.05
LOW_CONFIDENCE_THRESHOLD = 0.75

ACTIVE_BANK = "DEFAULT"
FORCE_YEAR = datetime.now().year


# ============================================================
# 2) Bank Configs
# ============================================================

BANK_CONFIGS = {
    "DEFAULT": {
        "amount_mode": "amount_balance",
        "bank_keywords": [],
        "opening_keywords": [
            "ยอด ยก มา", "ยอดยกมา", "ยอด ยก", "ยก มา",
            "BALANCE BROUGHT FORWARD", "BROUGHT FORWARD", "BALANCE B/F", "B/F",
        ],
        "credit_keywords": [
            "รับ โอน", "รับโอน", "รับ โอน เงิน", "ฝาก", "ฝาก เงิน", "ฝาก เงินสด",
            "ฝากเงินสด", "ดอกเบี้ย", "คืนเงิน", "เงิน เข้า", "เงินเข้า",
            "CREDIT", "DEPOSIT", "TRANSFER IN",
        ],
        "debit_keywords": [
            "โอน เงิน", "โอนเงิน", "โอน เงิน พร้อม เพ ย์", "โอน เงิน พร้อมเพย์",
            "จ่าย บิล", "จ่าย คิว อา ร์", "จ่าย QR", "ชำระ", "ชา ระ", "ชํา ระ",
            "ถอน", "ถอนเงิน", "หัก", "ค่าธรรมเนียม", "จ่าย",
            "DEBIT", "WITHDRAW", "WITHDRAWAL", "PAYMENT", "FEE", "TRANSFER OUT",
        ],
    },
    "KBANK": {
        "amount_mode": "amount_balance",
        "bank_keywords": ["KASIKORN", "KBANK", "K PLUS", "กสิกร", "ธนาคาร กสิกร ไทย", "ธนาคารกสิกรไทย"],
        "opening_keywords": ["ยอด ยก มา", "ยอดยกมา", "ยอด ยก", "ยก มา", "BALANCE BROUGHT FORWARD", "BROUGHT FORWARD"],
        "credit_keywords": ["รับ โอน", "รับโอน", "รับ โอน เงิน", "ฝาก", "ฝาก เงินสด", "เงิน เข้า", "เงินเข้า", "ดอกเบี้ย", "CREDIT", "DEPOSIT"],
        "debit_keywords": ["โอน เงิน", "โอนเงิน", "โอน เงิน พร้อม เพ ย์", "โอน เงิน พร้อมเพย์", "จ่าย คิว อา ร์", "จ่าย QR", "ถอน", "ถอนเงิน", "หัก", "ชำระ", "ค่าธรรมเนียม", "DEBIT", "WITHDRAW", "PAYMENT", "FEE"],
    },
    "KRUNGSRI": {
        "amount_mode": "amount_balance",
        "bank_keywords": ["KRUNGSRI", "กรุงศรี", "กรุง ศรี", "BANK OF AYUDHYA", "AYUDHYA", "MOBILE KRUNGSRI", "KRUNGSRI VISA", "FIRSTCHOICE", "FIRST CHOICE"],
        "opening_keywords": ["ยอด ยก มา", "ยอดยกมา", "ยอด ยก", "ยก มา", "BALANCE BROUGHT FORWARD", "BROUGHT FORWARD", "BALANCE B/F", "B/F"],
        "credit_keywords": ["รับ โอน เงิน", "รับ โอน", "รับโอน", "ฝาก เงิน", "ฝาก", "ดอกเบี้ย เงิน ฝาก", "ดอกเบี้ย", "คืนเงิน", "เงิน เข้า", "เงินเข้า", "CREDIT", "DEPOSIT", "TRANSFER IN"],
        "debit_keywords": ["จ่าย บิล", "จ่าย คิว อา ร์", "จ่าย QR", "โอน เงิน พร้อม เพ ย์", "โอน เงิน พร้อมเพย์", "โอน เงิน", "โอนเงิน", "ถอนเงิน", "ถอน", "หัก", "ค่าธรรมเนียม", "ชำระ", "จ่าย", "DEBIT", "WITHDRAW", "WITHDRAWAL", "PAYMENT", "FEE"],
    },
    "BBL": {
        "amount_mode": "amount_balance",
        "bank_keywords": ["BANGKOK BANK", "BANGKOKBANK", "ธนาคารกรุงเทพ", "ธนาคาร กรุงเทพ", "บัวหลวง", "BUALUANG", "STATEMENT OF SAVING ACCOUNT"],
        "opening_keywords": ["B/F", "B / F", "BALANCE B/F", "BALANCE BROUGHT FORWARD", "BROUGHT FORWARD", "ยอด ยก มา", "ยอดยกมา", "ยอด ยก", "ยก มา"],
        "credit_keywords": ["TRF FR OTH BK", "TRF FROTH BK", "TRF FROM OTH BK", "TRF FROM OTHER BANK", "TRANSFER IN", "SALARY", "SMART", "CREDIT", "DEPOSIT", "ฝาก", "รับ โอน", "รับโอน", "เงิน เข้า", "เงินเข้า"],
        "debit_keywords": ["TRF TO OTH BK", "TRF TOOTH BK", "TRF TO OTHER BANK", "TRF. PROMPTPAY", "TRF . PROMPTPAY", "PMT. PROMPTPAY", "PMT . PROMPTPAY", "PMT.PROMPTPAY", "PMT FOR GOODS", "CASH W/D ATM", "CASH W / D ATM", "WITHDRAWAL", "WITHDRAW", "TRANSFER", "PAYMENT", "DEBIT", "FEE", "ถอน", "ถอนเงิน", "โอน เงิน", "โอนเงิน", "ชำระ", "จ่าย"],
    },
    "SCB": {
        "amount_mode": "amount_balance",
        "bank_keywords": ["SCB", "ไทยพาณิชย์", "ไทย พาณิชย์", "ธนาคารไทยพาณิชย์", "ธนาคาร ไทย พาณิชย์", "THE SIAM COMMERCIAL BANK", "SIAM COMMERCIAL BANK", "STATEMENT OF SAVING ACCOUNT"],
        "opening_keywords": ["ยอดเงินคงเหลือยกมา", "ยอด เงิน คงเหลือ ยก มา", "BALANCE BROUGHT FORWARD", "BROUGHT FORWARD", "BALANCE B/F", "B/F"],
        "credit_keywords": [" X1 ", "X1", "รับ โอน จาก", "รับโอนจาก", "รับ โอน", "รับโอน", "โอน จาก", "โอนจาก", "ฝาก", "เงิน เข้า", "เงินเข้า", "CREDIT", "DEPOSIT", "TRANSFER IN"],
        "debit_keywords": [" X2 ", "X2", "โอน ไป", "โอนไป", "PromptPay", "PROMPTPAY", "จ่าย บิล", "จ่ายบิล", "จ่าย", "ถอน", "ถอนเงิน", "ชำระ", "ชา ระ", "ชํา ระ", "PAYMENT", "WITHDRAW", "WITHDRAWAL", "TRANSFER OUT", "FEE"],
    },
}


# ============================================================
# 3) Basic Helpers
# ============================================================

def clean_text(s):
    s = str(s).replace("\n", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def get_box_info(vertices):
    xs = [v.x for v in vertices]
    ys = [v.y for v in vertices]
    x_min, y_min, x_max, y_max = min(xs), min(ys), max(xs), max(ys)
    return {
        "x_min": x_min,
        "y_min": y_min,
        "x_max": x_max,
        "y_max": y_max,
        "x_center": (x_min + x_max) / 2,
        "y_center": (y_min + y_max) / 2,
        "width": x_max - x_min,
        "height": y_max - y_min,
    }


def get_bank_config(bank_name=None):
    if bank_name is None:
        bank_name = ACTIVE_BANK
    return BANK_CONFIGS.get(bank_name, BANK_CONFIGS["DEFAULT"])


# ============================================================
# 4) Image Preprocessing
# ============================================================

def preprocess_image_variants(input_path):
    img = Image.open(input_path).convert("RGB")
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    variants = [("original", input_path)]

    img1 = img.resize((img.width * 2, img.height * 2), Image.LANCZOS)
    img1 = ImageEnhance.Contrast(img1).enhance(1.9)
    img1 = ImageEnhance.Sharpness(img1).enhance(2.2)
    p1 = os.path.join(tempfile.gettempdir(), f"pre_{base_name}_2x_contrast_sharp.png")
    img1.save(p1)
    variants.append(("2x_contrast_sharp", p1))

    img2 = img.resize((img.width * 3, img.height * 3), Image.LANCZOS).convert("L")
    img2 = ImageOps.autocontrast(img2)
    img2 = ImageEnhance.Contrast(img2).enhance(1.8)
    img2 = ImageEnhance.Sharpness(img2).enhance(2.0)
    p2 = os.path.join(tempfile.gettempdir(), f"pre_{base_name}_3x_gray_autocontrast.png")
    img2.save(p2)
    variants.append(("3x_gray_autocontrast", p2))

    img3 = img.resize((img.width * 3, img.height * 3), Image.LANCZOS).convert("L")
    img3 = ImageOps.autocontrast(img3)
    img3 = img3.filter(ImageFilter.UnsharpMask(radius=1.2, percent=180, threshold=3))
    p3 = os.path.join(tempfile.gettempdir(), f"pre_{base_name}_3x_unsharp.png")
    img3.save(p3)
    variants.append(("3x_unsharp", p3))

    img4 = img.resize((img.width * 3, img.height * 3), Image.LANCZOS).convert("L")
    img4 = ImageOps.autocontrast(img4)
    img4 = ImageEnhance.Contrast(img4).enhance(2.0)
    img4 = img4.point(lambda x: 255 if x > 175 else 0)
    p4 = os.path.join(tempfile.gettempdir(), f"pre_{base_name}_3x_bw_threshold_175.png")
    img4.save(p4)
    variants.append(("3x_bw_threshold_175", p4))

    img5 = img.resize((img.width * 3, img.height * 3), Image.LANCZOS).convert("L")
    img5 = ImageOps.autocontrast(img5)
    img5 = ImageEnhance.Contrast(img5).enhance(1.7)
    img5 = img5.point(lambda x: 255 if x > 145 else 0)
    p5 = os.path.join(tempfile.gettempdir(), f"pre_{base_name}_3x_bw_threshold_145.png")
    img5.save(p5)
    variants.append(("3x_bw_threshold_145", p5))

    img6 = img.resize((img.width * 3, img.height * 3), Image.LANCZOS).convert("L")
    img6 = ImageOps.autocontrast(img6)
    img6 = img6.filter(ImageFilter.MedianFilter(size=3))
    img6 = img6.filter(ImageFilter.UnsharpMask(radius=1.0, percent=160, threshold=2))
    p6 = os.path.join(tempfile.gettempdir(), f"pre_{base_name}_3x_denoise_sharp.png")
    img6.save(p6)
    variants.append(("3x_denoise_sharp", p6))

    return variants


# ============================================================
# 5) Money Helpers
# ============================================================

def normalize_amount_text(s):
    s = str(s).strip()
    s = re.sub(r"\s+", "", s)
    ocr_character_map = {"O": "0", "o": "0", "I": "1", "l": "1", "|": "1", "S": "5", "s": "5", "B": "8", "g": "9", "q": "9", "Z": "2", "z": "2"}
    for char, digit in ocr_character_map.items():
        s = s.replace(char, digit)
    if re.match(r"^\d{1,3}(,\d{3})+,\d{2}$", s):
        s = s[:-3] + "." + s[-2:]
    if re.match(r"^\d+,\d{2}$", s):
        s = s.replace(",", ".")
    if re.match(r"^\d{1,3}\.\d{3}\.\d{2}$", s):
        parts = s.split(".")
        s = parts[0] + "," + parts[1] + "." + parts[2]
    return s


def parse_money(s):
    if pd.isna(s):
        return None
    s = normalize_amount_text(str(s)).replace(" ", "")
    s = re.sub(r"[^0-9,.\-]", "", s)
    if not s:
        return None
    try:
        return float(s.replace(",", ""))
    except Exception:
        return None


def is_money_token(s):
    s = normalize_amount_text(str(s).strip())
    patterns = [
        r"^-?\d{1,3}(,\d{3})*(\.\d{2})$",
        r"^-?\d+\.\d{2}$",
        r"^-?\d+,\d{2}$",
        r"^-?\d{1,3}\.\d{3}\.\d{2}$",
    ]
    return any(re.match(p, s) for p in patterns)


def money_to_key(value):
    if value is None or pd.isna(value):
        return ""
    try:
        return str(int(round(float(value) * 100)))
    except Exception:
        return ""


def extract_money_values_loose(token):
    token = str(token).strip()
    ocr_character_map = {"O": "0", "o": "0", "I": "1", "l": "1", "|": "1", "S": "5", "s": "5", "B": "8", "g": "9", "q": "9", "Z": "2", "z": "2"}
    for char, digit in ocr_character_map.items():
        token = token.replace(char, digit)
    priority_patterns = [
        r"(?<![\d,])-?\d{1,3}(?:,\d{3})+\.\d{2}(?!\d)",
        r"(?<![\d.])-?\d{1,3}(?:\.\d{3})+,\d{2}(?!\d)",
        r"(?<![\d.])-?\d{1,3}(?:\.\d{3})+\.\d{2}(?!\d)",
        r"(?<![\d,])-?\d+\.\d{2}(?!\d)",
        r"(?<![\d,])-?\d+,\d{2}(?![\d.])",
    ]
    for pattern in priority_patterns:
        matches = re.findall(pattern, token)
        if matches:
            values = []
            for m in matches:
                value = parse_money(m)
                if value is not None:
                    values.append(value)
            return values
    return []


def score_money_candidate(text, value):
    text = str(text)
    score = 0
    if re.search(r"\d{1,3},\d{3}\.\d{2}", text):
        score += 40
    if re.search(r"\d+[.,]\d{2}", text):
        score += 20
    digits = re.sub(r"\D", "", text)
    score += min(len(digits), 10)
    if value is not None and not pd.isna(value):
        if abs(float(value)) >= 1000:
            score += 8
        elif abs(float(value)) >= 100:
            score += 5
    return score


def pick_best_money_candidate(candidates):
    if not candidates:
        return None
    candidates = sorted(candidates, key=lambda x: (score_money_candidate(x.get("text", ""), x.get("value")), x.get("x_center", 0)), reverse=True)
    return candidates[0].get("value")


def amount_balance_match(prev_balance, amount, balance, tolerance=BALANCE_TOLERANCE):
    if prev_balance is None or amount is None or balance is None or pd.isna(prev_balance) or pd.isna(amount) or pd.isna(balance):
        return None
    debit_exp = round(prev_balance - amount, 2)
    credit_exp = round(prev_balance + amount, 2)
    if abs(balance - debit_exp) <= tolerance:
        return {"type": "debit", "expected_balance": debit_exp, "diff": round(balance - debit_exp, 2), "balance_check": "OK_DEBIT"}
    if abs(balance - credit_exp) <= tolerance:
        return {"type": "credit", "expected_balance": credit_exp, "diff": round(balance - credit_exp, 2), "balance_check": "OK_CREDIT"}
    return None


def try_healing_amount(prev_balance, read_amount, balance, tolerance=0.05):
    if prev_balance is None or balance is None or pd.isna(prev_balance) or pd.isna(balance):
        return None, None
    expected_debit = round(prev_balance - balance, 2)
    expected_credit = round(balance - prev_balance, 2)
    expected_amount = expected_debit if expected_debit > 0 else expected_credit
    tx_type = "debit" if expected_debit > 0 else "credit"
    if read_amount is None or pd.isna(read_amount):
        return None, None
    str_read = f"{read_amount:.2f}"
    str_exp = f"{expected_amount:.2f}"
    if abs(read_amount / 100 - expected_amount) <= tolerance:
        return expected_amount, f"healed_missing_dot_{tx_type}"
    if len(str_read) == len(str_exp):
        mismatches = sum(1 for a, b in zip(str_read, str_exp) if a != b)
        if mismatches <= 1:
            return expected_amount, f"healed_swapped_digit_{tx_type}"
    if len(str_exp) > len(str_read) and (str_read in str_exp or str_exp.endswith(str_read)):
        return expected_amount, f"healed_omitted_digit_{tx_type}"
    return None, None


def suggest_amount_from_balance(prev_balance, balance):
    if prev_balance is None or balance is None or pd.isna(prev_balance) or pd.isna(balance):
        return None, None
    diff = round(balance - prev_balance, 2)
    if diff > 0:
        return abs(diff), "credit"
    if diff < 0:
        return abs(diff), "debit"
    return 0.0, "zero"


# ============================================================
# 6) Date Helpers
# ============================================================

def fix_ocr_date_token(token):
    s = str(token).strip()
    s = s.replace("O", "0").replace("o", "0").replace("I", "1").replace("l", "1")
    s = s.replace(".", "-").replace("/", "-")
    return s


def is_date_token(s):
    s = fix_ocr_date_token(s)
    patterns = [r"^\d{1,2}[-]\d{1,2}[-]\d{2,4}$", r"^\d{1,2}[-]\d{1,2}$"]
    return any(re.match(p, s) for p in patterns)


def normalize_date(s):
    s = fix_ocr_date_token(s)
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
    except Exception:
        return None
    return dt.strftime("%d-%m-%Y")


def extract_date_from_line(line_text):
    for token in str(line_text).split():
        token_fixed = fix_ocr_date_token(token)
        if is_date_token(token_fixed):
            return normalize_date(token_fixed)
    return None


def parse_date_for_sort(date_text):
    if pd.isna(date_text) or not str(date_text).strip():
        return pd.NaT
    return pd.to_datetime(str(date_text).strip().replace("/", "-"), format="%d-%m-%Y", errors="coerce")


def is_time_token(s):
    return bool(re.match(r"^\d{1,2}:\d{2}(:\d{2})?$", str(s).strip()))


def extract_time_from_line(line_text):
    for token in str(line_text).split():
        if is_time_token(token):
            return token
    return None


# ============================================================
# 7) Text / Transaction Helpers
# ============================================================

def join_words(g):
    if g.empty:
        return ""
    return clean_text(" ".join(g.sort_values("x_center")["text"].astype(str).tolist()))


def extract_money_from_line_words(g):
    money_items = []
    for _, row in g.sort_values("x_center").iterrows():
        text = normalize_amount_text(str(row["text"]))
        if is_money_token(text) or extract_money_values_loose(text):
            value = parse_money(text)
            if value is not None:
                money_items.append({"text": text, "value": value, "x_center": row["x_center"], "confidence": row.get("word_confidence", None)})
    return money_items


def has_opening_balance_text(text, bank_name=None):
    text = clean_text(text).upper()
    cfg = get_bank_config(bank_name)
    for kw in cfg.get("opening_keywords", []):
        if kw.upper() in text:
            return True
    for kw in BANK_CONFIGS["DEFAULT"].get("opening_keywords", []):
        if kw.upper() in text:
            return True
    return False


def classify_by_keyword(text, bank_name=None):
    text = clean_text(text).upper()
    cfg = get_bank_config(bank_name)
    for kw in cfg.get("credit_keywords", []):
        if kw.upper() in text:
            return "credit"
    for kw in cfg.get("debit_keywords", []):
        if kw.upper() in text:
            return "debit"
    for kw in BANK_CONFIGS["DEFAULT"].get("credit_keywords", []):
        if kw.upper() in text:
            return "credit"
    for kw in BANK_CONFIGS["DEFAULT"].get("debit_keywords", []):
        if kw.upper() in text:
            return "debit"
    return "unknown"


def is_noise_line(text):
    text_upper = clean_text(text).upper()
    noise_keywords = [
        "STATEMENT OF SAVING ACCOUNT", "THE SIAM COMMERCIAL BANK", "ACCOUNT NO", "เลขที่บัญชี",
        "ชื่อ - สกุล", "ADDRESS", "สาขา", "BRANCH", "DATE | TIME", "DEBIT/CREDIT", "BALANCE/BAHT",
        "DESCRIPTION/NOTE", "TOTAL AMOUNTS", "TOTAL ITEMS", "THIS DOCUMENT IS AUTO-GENERATED",
        "เอกสาร ฉบับ นี้", "ออก โดย ระบบ อัตโนมัติ",
    ]
    if any(kw.upper() in text_upper for kw in noise_keywords):
        return True
    if re.search(r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\s*[-–]\s*\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", text_upper):
        return True
    if re.search(r"หน้า\s*\d+\s*/\s*\d+", text_upper):
        return True
    return False


def extract_page_no_from_text(text):
    text = clean_text(text)
    m = re.search(r"หน้า\s*(\d+)\s*/\s*(\d+)", text)
    if m:
        return int(m.group(1))
    m = re.search(r"\b(\d+)\s*/\s*(\d+)\b", text)
    if m:
        current, total = int(m.group(1)), int(m.group(2))
        if 1 <= current <= total <= 20:
            return current
    return None


def infer_page_no_map_from_line_words(line_words_df):
    page_map = {}
    if line_words_df.empty:
        return page_map
    for file_name, fg in line_words_df.groupby("file_name"):
        found_page = None
        for _, lg in fg.groupby("line_id"):
            line_text = join_words(lg)
            page_no = extract_page_no_from_text(line_text)
            if page_no is not None:
                found_page = page_no
                break
        if found_page is not None:
            page_map[file_name] = found_page
    return page_map


# ============================================================
# 8) SCB Specific Helpers
# ============================================================

def extract_scb_code_from_line(text):
    for token in re.split(r"\s+", clean_text(text).upper()):
        token = token.strip("|:;,.()[]{}")
        if token in ["X1", "X2"]:
            return token
    return None


def extract_scb_money_by_column(g):
    if g.empty or "page_width" not in g.columns:
        return {"amount": None, "balance": None, "money_texts": []}
    page_width = g["page_width"].dropna().iloc[0]
    amount_x_min, amount_x_max = page_width * 0.25, page_width * 0.52
    balance_x_min, balance_x_max = page_width * 0.52, page_width * 0.70
    amount_candidates, balance_candidates, money_texts = [], [], []
    for _, row in g.sort_values("x_center").iterrows():
        x, text = row["x_center"], str(row["text"])
        if amount_x_min <= x < amount_x_max:
            for value in extract_money_values_loose(text):
                amount_candidates.append({"value": value, "x_center": x, "text": text})
                money_texts.append(str(text))
        elif balance_x_min <= x < balance_x_max:
            for value in extract_money_values_loose(text):
                balance_candidates.append({"value": value, "x_center": x, "text": text})
                money_texts.append(str(text))
    return {"amount": pick_best_money_candidate(amount_candidates), "balance": pick_best_money_candidate(balance_candidates), "money_texts": money_texts}


def parse_scb_line(g, page_no=None):
    g = g.sort_values("x_center").copy()
    raw_line_text = join_words(g)
    if is_noise_line(raw_line_text):
        return None
    date, time = extract_date_from_line(raw_line_text), extract_time_from_line(raw_line_text)
    line_confidence, min_confidence = None, None
    if "word_confidence" in g.columns:
        line_confidence = float(g["word_confidence"].mean())
        min_confidence = float(g["word_confidence"].min())
    if has_opening_balance_text(raw_line_text, "SCB"):
        money_info = extract_scb_money_by_column(g)
        balance = money_info.get("balance")
        if balance is None:
            money_items = extract_money_from_line_words(g)
            balance = money_items[-1]["value"] if money_items else None
            money_texts = [m["text"] for m in money_items]
        else:
            money_texts = money_info.get("money_texts", [])
        if balance is None:
            return None
        return {
            "file_name": g["file_name"].iloc[0], "line_id": int(g["line_id"].iloc[0]), "page_no": page_no,
            "date": None, "time": None, "code": "OPENING", "amount": None, "balance": balance,
            "money_tokens": " | ".join(money_texts), "money_count": len(money_texts), "raw_line_text": raw_line_text,
            "y_center": g["y_center"].mean(), "parsed_date": pd.NaT, "is_opening_balance": True,
            "line_confidence": line_confidence, "min_confidence": min_confidence,
        }
    code = extract_scb_code_from_line(raw_line_text)
    if not date or code not in ["X1", "X2"]:
        return None
    money_info = extract_scb_money_by_column(g)
    return {
        "file_name": g["file_name"].iloc[0], "line_id": int(g["line_id"].iloc[0]), "page_no": page_no,
        "date": date, "time": time, "code": code, "amount": money_info.get("amount"), "balance": money_info.get("balance"),
        "money_tokens": " | ".join(money_info.get("money_texts", [])), "money_count": len(money_info.get("money_texts", [])),
        "raw_line_text": raw_line_text, "y_center": g["y_center"].mean(), "parsed_date": parse_date_for_sort(date),
        "is_opening_balance": False, "line_confidence": line_confidence, "min_confidence": min_confidence,
    }


# ============================================================
# 9) Balance Chain Sorting
# ============================================================

def advanced_build_chain_with_healing(work_df, start_index, tolerance=0.05):
    candidates = work_df.copy().reset_index(drop=True)
    start_row = candidates.loc[start_index].to_dict()
    candidates = candidates.drop(index=start_index).reset_index(drop=True)
    ordered_rows = []
    chain_match_count = 0
    start_row["order_method"] = "CHAIN_START"
    ordered_rows.append(start_row)
    prev_balance = start_row.get("balance")
    while not candidates.empty:
        found_index, found_match, is_healed_match, healed_info = None, None, False, {}
        for idx, row in candidates.iterrows():
            match = amount_balance_match(prev_balance=prev_balance, amount=row.get("amount"), balance=row.get("balance"), tolerance=tolerance)
            if match is not None:
                found_index, found_match = idx, match
                break
        if found_index is None and prev_balance is not None:
            for idx, row in candidates.iterrows():
                h_amount, h_type = try_healing_amount(prev_balance, row.get("amount"), row.get("balance"), tolerance)
                if h_amount is not None:
                    found_index = idx
                    is_healed_match = True
                    healed_info = {"amount": h_amount, "type": h_type}
                    break
        if found_index is None:
            break
        next_row = candidates.loc[found_index].to_dict()
        if is_healed_match:
            next_row["amount"] = healed_info["amount"]
            next_row["order_method"] = f"BALANCE_CHAIN_{healed_info['type'].upper()}"
            next_row["chain_type"] = "debit" if "debit" in healed_info["type"] else "credit"
        else:
            next_row["order_method"] = "BALANCE_CHAIN_PERFECT"
            next_row["chain_type"] = found_match["type"]
        ordered_rows.append(next_row)
        chain_match_count += 1
        if next_row.get("balance") is not None and not pd.isna(next_row.get("balance")):
            prev_balance = next_row.get("balance")
        candidates = candidates.drop(index=found_index).reset_index(drop=True)
    return pd.DataFrame(ordered_rows), candidates, chain_match_count


def smart_order_by_balance_chain(df):
    if df.empty:
        return df
    if ACTIVE_BANK == "SCB":
        sort_cols = []
        if "page_no" in df.columns and df["page_no"].notna().any():
            sort_cols.append("page_no")
        sort_cols.extend(["file_name", "y_center"])
        result = df.copy().sort_values(sort_cols, na_position="last").reset_index(drop=True)
        result["order_method"] = "SCB_PAGE_LAYOUT"
        return result
    work = df.copy().reset_index(drop=True)
    opening_df = work[work["is_opening_balance"] == True].copy()
    tx_df = work[work["is_opening_balance"] == False].copy()
    possible_starts = []
    if not opening_df.empty:
        opening_df = opening_df.reset_index(drop=True)
        for _, opening_row in opening_df.iterrows():
            temp_df = pd.concat([pd.DataFrame([opening_row]), tx_df], ignore_index=True)
            possible_starts.append({"df": temp_df, "start_index": 0, "start_type": "OPENING_START"})
    tx_temp = tx_df.reset_index(drop=True)
    for idx in range(len(tx_temp)):
        possible_starts.append({"df": tx_temp, "start_index": idx, "start_type": "TX_START"})
    best_ordered, best_remaining, best_score = None, None, -999999
    for option in possible_starts:
        test_df = option["df"].copy().reset_index(drop=True)
        ordered_df, remaining_df, match_count = advanced_build_chain_with_healing(test_df, option["start_index"], tolerance=BALANCE_TOLERANCE)
        score = match_count * 1000
        if option["start_type"] == "OPENING_START":
            score += 500
        score -= len(remaining_df) * 50
        if "parsed_date" in ordered_df.columns:
            score += ordered_df["parsed_date"].notna().sum()
        if score > best_score:
            best_score, best_ordered, best_remaining = score, ordered_df.copy(), remaining_df.copy()
    if best_ordered is None:
        return work
    if best_remaining is not None and not best_remaining.empty:
        remaining = best_remaining.copy()
        sort_fallback = []
        if "page_no" in remaining.columns and remaining["page_no"].notna().any():
            sort_fallback.append("page_no")
        sort_fallback.extend(["parsed_date", "time", "y_center"])
        remaining = remaining.sort_values(by=sort_fallback, na_position="last").reset_index(drop=True)
        remaining["order_method"] = "CHAIN_BROKEN_FALLBACK_REVIEW"
        result = pd.concat([best_ordered, remaining], ignore_index=True)
    else:
        result = best_ordered
    return result.reset_index(drop=True)


# ============================================================
# 10) OCR
# ============================================================

def vision_ocr_image(image_path, client, language_hints=None):
    if language_hints is None:
        language_hints = ["th", "en"]
    with open(image_path, "rb") as f:
        content = f.read()
    image = vision.Image(content=content)
    response = client.document_text_detection(image=image, image_context=vision.ImageContext(language_hints=language_hints))
    if response.error.message:
        raise Exception(response.error.message)
    rows = []
    for page_idx, page in enumerate(response.full_text_annotation.pages):
        for block_idx, block in enumerate(page.blocks):
            for paragraph_idx, paragraph in enumerate(block.paragraphs):
                for word_idx, word in enumerate(paragraph.words):
                    word_text = "".join([s.text for s in word.symbols])
                    word_box = get_box_info(word.bounding_box.vertices)
                    rows.append({
                        "file_name": os.path.basename(image_path), "page_idx": page_idx,
                        "page_width": page.width, "page_height": page.height, "block_idx": block_idx,
                        "paragraph_idx": paragraph_idx, "word_idx": word_idx, "text": word_text,
                        "word_confidence": word.confidence, "x_min": word_box["x_min"], "y_min": word_box["y_min"],
                        "x_max": word_box["x_max"], "y_max": word_box["y_max"], "x_center": word_box["x_center"],
                        "y_center": word_box["y_center"], "width": word_box["width"], "height": word_box["height"],
                    })
    return response.full_text_annotation.text, pd.DataFrame(rows)


def rebuild_lines_by_y(words_df, y_threshold=10):
    all_lines = []
    for file_name, g in words_df.groupby("file_name"):
        if g.empty:
            continue
        page_width = g["page_width"].dropna().iloc[0]
        g = g[g["x_center"] >= page_width * 0.04].copy()
        if g.empty:
            continue
        g = g.sort_values(["y_center", "x_center"]).reset_index(drop=True)
        current_line, current_y, line_id = [], None, 0
        for _, row in g.iterrows():
            y = row["y_center"]
            if current_y is None:
                current_line, current_y = [row], y
                continue
            if abs(y - current_y) <= y_threshold:
                current_line.append(row)
                current_y = (current_y + y) / 2
            else:
                line_df = pd.DataFrame(current_line)
                line_df["line_id"] = line_id
                all_lines.append(line_df)
                line_id += 1
                current_line, current_y = [row], y
        if current_line:
            line_df = pd.DataFrame(current_line)
            line_df["line_id"] = line_id
            all_lines.append(line_df)
    return pd.concat(all_lines, ignore_index=True) if all_lines else pd.DataFrame()


def parse_simple_stm(words_df):
    line_words_df = rebuild_lines_by_y(words_df, y_threshold=10)
    if line_words_df.empty:
        return pd.DataFrame()
    line_rows = []
    page_no_map = infer_page_no_map_from_line_words(line_words_df)
    for (file_name, line_id), g in line_words_df.groupby(["file_name", "line_id"]):
        g = g.sort_values("x_center").copy()
        page_no = page_no_map.get(file_name)
        if ACTIVE_BANK == "SCB":
            parsed_line = parse_scb_line(g, page_no=page_no)
            if parsed_line is not None:
                line_rows.append(parsed_line)
            continue
        raw_line_text = join_words(g)
        if is_noise_line(raw_line_text):
            continue
        date, time = extract_date_from_line(raw_line_text), extract_time_from_line(raw_line_text)
        if not date and not has_opening_balance_text(raw_line_text):
            continue
        money_items = extract_money_from_line_words(g)
        money_values = [m["value"] for m in money_items]
        money_texts = [m["text"] for m in money_items]
        line_confidence, min_confidence = None, None
        if "word_confidence" in g.columns:
            line_confidence = float(g["word_confidence"].mean())
            min_confidence = float(g["word_confidence"].min())
        amount, balance = None, None
        if has_opening_balance_text(raw_line_text):
            if len(money_values) >= 1:
                balance = money_values[-1]
        else:
            if len(money_values) >= 2:
                amount, balance = money_values[0], money_values[-1]
            elif len(money_values) == 1:
                amount = money_values[0]
        line_rows.append({
            "file_name": file_name, "line_id": int(line_id), "page_no": page_no, "date": date, "time": time, "code": None,
            "amount": amount, "balance": balance, "money_tokens": " | ".join(money_texts), "money_count": len(money_values),
            "raw_line_text": raw_line_text, "y_center": g["y_center"].mean(), "parsed_date": parse_date_for_sort(date),
            "is_opening_balance": has_opening_balance_text(raw_line_text), "line_confidence": line_confidence, "min_confidence": min_confidence,
        })
    parsed_df = pd.DataFrame(line_rows)
    if parsed_df.empty:
        return parsed_df
    sort_cols = []
    if "page_no" in parsed_df.columns and parsed_df["page_no"].notna().any():
        sort_cols.append("page_no")
    return parsed_df.sort_values(sort_cols + ["file_name", "y_center"], na_position="last").reset_index(drop=True)


def score_parsed_result(parsed_df):
    if parsed_df is None or parsed_df.empty:
        return -999999
    conf_score = parsed_df["line_confidence"].fillna(0).mean() * 100 if "line_confidence" in parsed_df.columns else 0
    return len(parsed_df) * 100 + parsed_df["money_count"].fillna(0).sum() * 20 + parsed_df["date"].notna().sum() * 50 + conf_score


def ocr_best_variant_for_image(image_path, client, status_box=None):
    variants = preprocess_image_variants(image_path)
    best, best_score = None, None
    logs = []
    for v_name, v_path in variants:
        try:
            full_text, word_df = vision_ocr_image(v_path, client)
            if word_df.empty:
                logs.append(f"{v_name}: empty")
                continue
            word_df["file_name"] = os.path.basename(image_path)
            word_df["ocr_variant"] = v_name
            parsed_preview = parse_simple_stm(word_df)
            score = score_parsed_result(parsed_preview)
            logs.append(f"{v_name}: score={round(score, 2)}, rows={len(parsed_preview)}")
            if best_score is None or score > best_score:
                best_score = score
                best = {"variant_name": v_name, "variant_path": v_path, "full_text": full_text, "word_df": word_df, "score": score}
        except Exception as e:
            logs.append(f"{v_name}: OCR error: {e}")
    if best is None:
        raise Exception(f"OCR พังบนรูป {image_path}")
    if status_box is not None:
        status_box.write({"image": os.path.basename(image_path), "selected_variant": best["variant_name"], "score": round(best["score"], 2), "logs": logs})
    return best


# ============================================================
# 11) Final Check Logic
# ============================================================

def is_valid_number(x):
    if x is None:
        return False
    try:
        if pd.isna(x):
            return False
        float(x)
        return True
    except Exception:
        return False


def is_opening_row(row):
    raw_line_text = str(row.get("raw_line_text", ""))
    code = str(row.get("code", ""))
    balance_check = str(row.get("balance_check", ""))
    return has_opening_balance_text(raw_line_text) or code == "OPENING" or balance_check == "OPENING_BALANCE"


def build_final_parsed_stm_from_check(check_df):
    parsed_cols = ["date", "debit", "credit", "balance"]
    if check_df is None or check_df.empty:
        return pd.DataFrame(columns=parsed_cols)
    df = check_df.copy()
    required_cols = ["seq", "date", "debit", "credit", "balance", "balance_check", "raw_line_text", "code"]
    for col in required_cols:
        if col not in df.columns:
            df[col] = None
    opening_mask = df.apply(is_opening_row, axis=1)
    df = df[~opening_mask].copy()
    df = df[df["balance"].apply(is_valid_number)].copy()
    has_debit_or_credit = df["debit"].apply(is_valid_number) | df["credit"].apply(is_valid_number)
    df = df[has_debit_or_credit].copy()
    safe_mask = df["balance_check"].astype(str).str.contains("OK|SUGGEST", regex=True, na=False)
    df = df[safe_mask].copy()
    if "seq" in df.columns:
        df = df.sort_values("seq").reset_index(drop=True)
    else:
        df = df.reset_index(drop=True)
    final_df = df[parsed_cols].copy()
    for col in ["debit", "credit", "balance"]:
        final_df[col] = pd.to_numeric(final_df[col], errors="coerce")
    return final_df.reset_index(drop=True)


def add_debit_credit_and_check(ordered_df):
    parsed_cols = ["date", "debit", "credit", "balance"]
    check_cols = [
        "seq", "date", "time", "code", "debit", "credit", "balance", "prev_balance", "amount",
        "suggested_amount", "suggested_type", "expected_balance", "diff", "line_confidence", "min_confidence",
        "money_tokens", "raw_line_text", "order_method", "balance_check",
    ]
    if ordered_df is None or ordered_df.empty:
        return pd.DataFrame(columns=parsed_cols), pd.DataFrame(columns=check_cols)
    check_rows = []
    prev_balance = None
    seq = 0
    for _, row in ordered_df.iterrows():
        seq += 1
        date, time, code = row.get("date"), row.get("time"), row.get("code")
        amount, balance_val = row.get("amount"), row.get("balance")
        raw_line_text, money_tokens, order_method = row.get("raw_line_text", ""), row.get("money_tokens", ""), row.get("order_method", "")
        line_confidence, min_confidence = row.get("line_confidence"), row.get("min_confidence")
        debit = credit = expected_balance = diff = suggested_amount = suggested_type = None
        balance_check = ""
        low_conf = ((line_confidence is not None and not pd.isna(line_confidence) and line_confidence < LOW_CONFIDENCE_THRESHOLD) or (min_confidence is not None and not pd.isna(min_confidence) and min_confidence < LOW_CONFIDENCE_THRESHOLD))
        if has_opening_balance_text(raw_line_text) or code == "OPENING":
            balance_check = "OPENING_BALANCE"
        elif ACTIVE_BANK == "SCB" and code in ["X1", "X2"]:
            if prev_balance is None:
                balance_check = "SCB_NO_PREV_BALANCE_REVIEW"
            elif balance_val is None or pd.isna(balance_val):
                balance_check = "SCB_NO_BALANCE_REVIEW"
            elif amount is None or pd.isna(amount):
                if code == "X1":
                    suggested_amount = round(balance_val - prev_balance, 2)
                    suggested_type = "credit"
                    credit = suggested_amount
                    expected_balance = balance_val
                    diff = 0.0
                    balance_check = "SCB_X1_NO_AMOUNT_SUGGEST_CREDIT_REVIEW"
                else:
                    suggested_amount = round(prev_balance - balance_val, 2)
                    suggested_type = "debit"
                    debit = suggested_amount
                    expected_balance = balance_val
                    diff = 0.0
                    balance_check = "SCB_X2_NO_AMOUNT_SUGGEST_DEBIT_REVIEW"
            else:
                if code == "X1":
                    expected_balance = round(prev_balance + amount, 2)
                    diff = round(balance_val - expected_balance, 2)
                    if abs(diff) <= BALANCE_TOLERANCE:
                        credit = amount
                        balance_check = "OK_CREDIT"
                    else:
                        healed_amt, heal_type = try_healing_amount(prev_balance, amount, balance_val, BALANCE_TOLERANCE)
                        if healed_amt is not None:
                            credit = healed_amt
                            expected_balance = balance_val
                            diff = 0.0
                            balance_check = f"OK_{heal_type.upper()}"
                        else:
                            credit = amount
                            balance_check = "SCB_X1_SUGGEST_AMOUNT_REVIEW"
                else:
                    expected_balance = round(prev_balance - amount, 2)
                    diff = round(balance_val - expected_balance, 2)
                    if abs(diff) <= BALANCE_TOLERANCE:
                        debit = amount
                        balance_check = "OK_DEBIT"
                    else:
                        healed_amt, heal_type = try_healing_amount(prev_balance, amount, balance_val, BALANCE_TOLERANCE)
                        if healed_amt is not None:
                            debit = healed_amt
                            expected_balance = balance_val
                            diff = 0.0
                            balance_check = f"OK_{heal_type.upper()}"
                        else:
                            debit = amount
                            balance_check = "SCB_X2_SUGGEST_AMOUNT_REVIEW"
        else:
            match = amount_balance_match(prev_balance, amount, balance_val, BALANCE_TOLERANCE)
            if match is not None:
                expected_balance = match["expected_balance"]
                diff = match["diff"]
                balance_check = match["balance_check"]
                if match["type"] == "debit":
                    debit = amount
                else:
                    credit = amount
            else:
                healed_amt, heal_type = try_healing_amount(prev_balance, amount, balance_val, BALANCE_TOLERANCE)
                if healed_amt is not None:
                    if "debit" in heal_type:
                        debit = healed_amt
                    else:
                        credit = healed_amt
                    expected_balance = balance_val
                    diff = 0.0
                    balance_check = f"OK_{heal_type.upper()}"
                else:
                    tx_type = classify_by_keyword(raw_line_text, ACTIVE_BANK)
                    suggested_amount, suggested_type = suggest_amount_from_balance(prev_balance, balance_val)
                    if suggested_amount is not None and suggested_type in ["debit", "credit"]:
                        expected_balance = round(prev_balance - suggested_amount if suggested_type == "debit" else prev_balance + suggested_amount, 2)
                        diff = round(balance_val - expected_balance, 2)
                        if tx_type == suggested_type:
                            balance_check = "SUGGEST_AMOUNT_BY_KEYWORD_REVIEW"
                            if suggested_type == "debit":
                                debit = suggested_amount
                            else:
                                credit = suggested_amount
                        else:
                            balance_check = "SUGGEST_AMOUNT_REVIEW"
                    else:
                        balance_check = "CHAIN_BROKEN_OCR_REVIEW"
        if low_conf and "OK" in balance_check:
            balance_check += "_LOW_CONFIDENCE"
        check_rows.append({
            "seq": seq, "date": date, "time": time, "code": code, "debit": debit, "credit": credit,
            "balance": balance_val, "prev_balance": prev_balance, "amount": amount, "suggested_amount": suggested_amount,
            "suggested_type": suggested_type, "expected_balance": expected_balance, "diff": diff,
            "line_confidence": line_confidence, "min_confidence": min_confidence, "money_tokens": money_tokens,
            "raw_line_text": raw_line_text, "order_method": order_method, "balance_check": balance_check,
        })
        if balance_val is not None and not pd.isna(balance_val):
            prev_balance = balance_val
    check_df = pd.DataFrame(check_rows)
    final_df = build_final_parsed_stm_from_check(check_df)
    return final_df, check_df


def build_duplicate_key(row):
    date_key = str(row.get("date", "")).strip()
    time_key = str(row.get("time", "")).strip()
    page_key = str(row.get("page_no", "")).strip()
    amount_key = money_to_key(row.get("amount"))
    balance_key = money_to_key(row.get("balance"))
    raw_line_text = str(row.get("raw_line_text", ""))
    if has_opening_balance_text(raw_line_text):
        return "OPENING|" + page_key + "|" + balance_key
    if time_key:
        return f"TX|{page_key}|{date_key}|{time_key}|{amount_key}|{balance_key}"
    return f"TX|{page_key}|{date_key}|{amount_key}|{balance_key}"


def deduplicate_rows(parsed_df):
    if parsed_df is None or parsed_df.empty:
        return parsed_df
    df = parsed_df.copy()
    df["duplicate_key"] = df.apply(build_duplicate_key, axis=1)
    df = df.drop_duplicates(subset=["duplicate_key"], keep="first").reset_index(drop=True)
    return df.drop(columns=["duplicate_key"], errors="ignore")


# ============================================================
# 12) Excel Export
# ============================================================

def create_summary_df(final_df, check_df, selected_bank):
    if final_df.empty:
        return pd.DataFrame(columns=["metric", "value"])
    db_s = pd.to_numeric(final_df["debit"], errors="coerce")
    cr_s = pd.to_numeric(final_df["credit"], errors="coerce")
    bl_s = pd.to_numeric(final_df["balance"], errors="coerce")
    return pd.DataFrame([
        ["selected_bank", selected_bank],
        ["transaction_rows", int(final_df["date"].notna().sum())],
        ["debit_total", float(db_s.fillna(0).sum())],
        ["credit_total", float(cr_s.fillna(0).sum())],
        ["net_change", float(cr_s.fillna(0).sum() - db_s.fillna(0).sum())],
        ["first_balance", float(bl_s.dropna().iloc[0]) if not bl_s.dropna().empty else None],
        ["last_balance", float(bl_s.dropna().iloc[-1]) if not bl_s.dropna().empty else None],
        ["ok_checks", int(check_df["balance_check"].astype(str).str.contains("OK").sum())],
        ["review_rows", int(check_df["balance_check"].astype(str).str.contains("REVIEW|BROKEN|MISSING|NO_PREV", regex=True).sum())],
    ], columns=["metric", "value"])


def prepare_check_export_df(check_df):
    if check_df.empty:
        return check_df
    df = check_df.drop(columns=["amount", "line_confidence", "min_confidence"], errors="ignore")
    df = df.rename(columns={"expected_balance": "expected_balance_py", "diff": "diff_py"})
    df["expected_balance_excel"] = None
    df["anomaly_reason_excel"] = None
    return df


def apply_check_sheet_formulas_and_format(output_excel, sheet_name="check"):
    wb = load_workbook(output_excel)
    if sheet_name not in wb.sheetnames:
        return
    ws = wb[sheet_name]
    headers = {str(cell.value): cell.column for cell in ws[1] if cell.value}
    required = ["debit", "credit", "balance", "prev_balance", "suggested_amount", "expected_balance_excel", "anomaly_reason_excel"]
    for col in required:
        if col not in headers:
            wb.save(output_excel)
            return
    debit_col = get_column_letter(headers["debit"])
    credit_col = get_column_letter(headers["credit"])
    balance_col = get_column_letter(headers["balance"])
    prev_col = get_column_letter(headers["prev_balance"])
    suggested_col = get_column_letter(headers["suggested_amount"])
    expected_col = get_column_letter(headers["expected_balance_excel"])
    reason_col = get_column_letter(headers["anomaly_reason_excel"])
    tol = BALANCE_TOLERANCE
    for row in range(2, ws.max_row + 1):
        ws[f"{expected_col}{row}"] = f'=IF({prev_col}{row}="","",ROUND({prev_col}{row}-IF({debit_col}{row}="",0,{debit_col}{row})+IF({credit_col}{row}="",0,{credit_col}{row}),2))'
        ws[f"{reason_col}{row}"] = f'=IF({prev_col}{row}="","OPENING_OR_NO_PREV",IF({balance_col}{row}="","MISSING_BALANCE",IF({suggested_col}{row}<>"","กระทบยอดให้ตรวจสอบ",IF(AND({debit_col}{row}="",{credit_col}{row}=""),"MISSING_DEBIT_CREDIT",IF(ABS({balance_col}{row}-{expected_col}{row})>{tol},"BALANCE_NOT_MATCH","OK")))))'
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    header_font = Font(color="000000", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin)
    number_cols = ["debit", "credit", "balance", "prev_balance", "suggested_amount", "expected_balance_py", "diff_py", "expected_balance_excel"]
    for col in number_cols:
        if col in headers:
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=headers[col]).number_format = "#,##0.00"
    ws.freeze_panes = "A2"
    light_red_fill = PatternFill("solid", fgColor="FCE4E4")
    dark_red_font = Font(color="9C0006")
    ws.conditional_formatting.add(
        f"A2:{get_column_letter(ws.max_column)}{ws.max_row}",
        FormulaRule(formula=[f'=AND(${reason_col}2<>"OK",${reason_col}2<>"OPENING_OR_NO_PREV",${reason_col}2<>"")'], fill=light_red_fill, font=dark_red_font),
    )
    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_len = max([len(str(cell.value or "")) for cell in ws[col_letter]])
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 45)
    wb.save(output_excel)


def build_excel_bytes(final_df, check_df, selected_bank):
    summary_df = create_summary_df(final_df, check_df, selected_bank)
    ocr_review_df = check_df[check_df["balance_check"].astype(str).str.contains("REVIEW|BROKEN|UNKNOWN|NO_PREV|MISSING|SUGGEST|HEALED", regex=True, na=False)].copy().reset_index(drop=True)
    check_export_df = prepare_check_export_df(check_df)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        output_excel = tmp.name
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="parsed_stm", index=False)
        check_export_df.to_excel(writer, sheet_name="check", index=False)
        summary_df.to_excel(writer, sheet_name="summary", index=False)
        ocr_review_df.to_excel(writer, sheet_name="ocr_review", index=False)
    apply_check_sheet_formulas_and_format(output_excel, sheet_name="check")
    with open(output_excel, "rb") as f:
        excel_bytes = f.read()
    return excel_bytes, summary_df, ocr_review_df


# ============================================================
# 13) Google Vision Credential
# ============================================================

def create_vision_client_from_uploaded_key(uploaded_key):
    if uploaded_key is None:
        raise Exception("กรุณาอัปโหลดไฟล์ Google Vision Service Account JSON ก่อน")
    key_bytes = uploaded_key.read()
    key_data = json.loads(key_bytes.decode("utf-8"))
    with tempfile.NamedTemporaryFile(delete=False, suffix=".json", mode="w", encoding="utf-8") as tmp:
        json.dump(key_data, tmp, ensure_ascii=False, indent=2)
        credential_path = tmp.name
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = credential_path
    return vision.ImageAnnotatorClient()


# ============================================================
# 14) Streamlit UI
# ============================================================

with st.sidebar:
    st.header("ตั้งค่าการประมวลผล")
    selected_bank = st.selectbox("เลือกธนาคาร", ["KBANK", "KRUNGSRI", "BBL", "SCB", "DEFAULT"], index=0, help="เลือกธนาคารเอง ระบบจะไม่ Auto Detect")
    selected_year = st.number_input("ปีของ Statement", min_value=2000, max_value=2100, value=datetime.now().year, step=1, help="ถ้า Statement ไม่มีปีในรายการ ระบบจะใช้ปีนี้")
    st.divider()
    uploaded_key = st.file_uploader("อัปโหลด Google Vision Service Account JSON", type=["json"])
    st.caption("ใช้ไฟล์ JSON จาก Google Cloud Service Account")

uploaded_images = st.file_uploader("อัปโหลดรูป Statement", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
process_btn = st.button("เริ่มประมวลผล", type="primary", use_container_width=True)


# ============================================================
# 15) Run Pipeline
# ============================================================

if process_btn:
    if not uploaded_images:
        st.error("กรุณาอัปโหลดรูป Statement ก่อน")
        st.stop()
    if uploaded_key is None:
        st.error("กรุณาอัปโหลด Google Vision Service Account JSON ก่อน")
        st.stop()

    ACTIVE_BANK = selected_bank
    FORCE_YEAR = int(selected_year)
    st.info(f"ธนาคารที่เลือก: {ACTIVE_BANK} | ปี Statement: {FORCE_YEAR}")

    try:
        client = create_vision_client_from_uploaded_key(uploaded_key)
        all_words = []
        all_full_text = []
        progress = st.progress(0)
        log_area = st.container()

        with st.spinner("กำลัง OCR และเลือก image variant ที่ดีที่สุด..."):
            for i, img_file in enumerate(uploaded_images):
                suffix = os.path.splitext(img_file.name)[1].lower()
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                    tmp.write(img_file.read())
                    image_path = tmp.name
                with log_area:
                    st.write(f"กำลังประมวลผล: {img_file.name}")
                best = ocr_best_variant_for_image(image_path, client, status_box=log_area)
                best["word_df"]["file_name"] = img_file.name
                all_words.append(best["word_df"])
                all_full_text.append(best["full_text"])
                progress.progress((i + 1) / len(uploaded_images))

        if not all_words:
            st.error("OCR ไม่พบข้อมูลจากรูป")
            st.stop()

        words_raw_df = pd.concat(all_words, ignore_index=True)
        with st.spinner("กำลังแยกรายการ Statement และตรวจสอบ balance chain..."):
            parsed_df = parse_simple_stm(words_raw_df)
            dedup_df = deduplicate_rows(parsed_df)
            ordered_df = smart_order_by_balance_chain(dedup_df)
            final_df, check_df = add_debit_credit_and_check(ordered_df)
            final_df = build_final_parsed_stm_from_check(check_df)

        excel_bytes, summary_df, ocr_review_df = build_excel_bytes(final_df, check_df, ACTIVE_BANK)
        st.success("ประมวลผลสำเร็จ")

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Parsed Rows", len(final_df))
        with col2:
            ok_count = int(check_df["balance_check"].astype(str).str.contains("OK", na=False).sum())
            st.metric("OK Checks", ok_count)
        with col3:
            review_count = int(check_df["balance_check"].astype(str).str.contains("REVIEW|BROKEN|MISSING|NO_PREV", regex=True, na=False).sum())
            st.metric("Review Rows", review_count)
        with col4:
            st.metric("Bank", ACTIVE_BANK)

        st.subheader("Preview parsed_stm")
        st.dataframe(final_df, use_container_width=True)

        st.subheader("Debug check")
        display_cols = ["seq", "date", "time", "debit", "credit", "balance", "prev_balance", "expected_balance", "diff", "balance_check", "raw_line_text"]
        display_cols = [c for c in display_cols if c in check_df.columns]
        st.dataframe(check_df[display_cols], use_container_width=True)

        st.subheader("Summary")
        st.dataframe(summary_df, use_container_width=True)

        st.download_button(
            label="ดาวน์โหลด Excel",
            data=excel_bytes,
            file_name=f"stm_image_{ACTIVE_BANK}_{FORCE_YEAR}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as e:
        st.exception(e)


