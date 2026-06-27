# ============================================================
# STM Image to Excel Parser - Streamlit Version (Production Ready)
# Manual Bank Selection: KBANK / KRUNGSRI / BBL / SCB / DEFAULT
# Google Cloud Vision OCR — Credentials from st.secrets
# ============================================================

import os
import re
import tempfile
import shutil
from datetime import datetime

import pandas as pd
import streamlit as st

from PIL import Image, ImageEnhance, ImageFilter, ImageOps
from google.cloud import vision
from google.oauth2 import service_account

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

# ============================================================
# 0) Streamlit Page Config
# ============================================================
st.set_page_config(page_title="STM Image to Excel Parser", page_icon="📄", layout="wide")
st.title("📄 STM Image to Excel Parser")
st.caption("OCR Statement รูปภาพ → Excel | เลือกธนาคารเอง ไม่ต้อง Auto Detect")

# ============================================================
# 1) Constants & Bank Configs
# ============================================================
BALANCE_TOLERANCE = 0.05
LOW_CONFIDENCE_THRESHOLD = 0.75

BANK_CONFIGS = {
    "DEFAULT": {
        "opening_keywords": ["ยอด ยก มา", "ยอดยกมา", "ยอด ยก", "ยก มา", "BALANCE BROUGHT FORWARD", "BROUGHT FORWARD", "BALANCE B/F", "B/F"],
        "credit_keywords": ["รับ โอน", "รับโอน", "รับ โอน เงิน", "ฝาก", "ฝาก เงิน", "ฝาก เงินสด", "ฝากเงินสด", "ดอกเบี้ย", "คืนเงิน", "เงิน เข้า", "เงินเข้า", "CREDIT", "DEPOSIT", "TRANSFER IN"],
        "debit_keywords": ["โอน เงิน", "โอนเงิน", "โอน เงิน พร้อม เพ ย์", "โอน เงิน พร้อมเพย์", "จ่าย บิล", "จ่าย คิว อา ร์", "จ่าย QR", "ชำระ", "ชา ระ", "ชํา ระ", "ถอน", "ถอนเงิน", "หัก", "ค่าธรรมเนียม", "จ่าย", "DEBIT", "WITHDRAW", "WITHDRAWAL", "PAYMENT", "FEE", "TRANSFER OUT"],
    },
    "KBANK": {
        "opening_keywords": ["ยอด ยก มา", "ยอดยกมา", "ยอด ยก", "ยก มา", "BALANCE BROUGHT FORWARD", "BROUGHT FORWARD"],
        "credit_keywords": ["รับ โอน", "รับโอน", "รับ โอน เงิน", "ฝาก", "ฝาก เงินสด", "เงิน เข้า", "เงินเข้า", "ดอกเบี้ย", "CREDIT", "DEPOSIT"],
        "debit_keywords": ["โอน เงิน", "โอนเงิน", "โอน เงิน พร้อม เพ ย์", "โอน เงิน พร้อมเพย์", "จ่าย คิว อา ร์", "จ่าย QR", "ถอน", "ถอนเงิน", "หัก", "ชำระ", "ค่าธรรมเนียม", "DEBIT", "WITHDRAW", "PAYMENT", "FEE"],
    },
    "KRUNGSRI": {
        "opening_keywords": ["ยอด ยก มา", "ยอดยกมา", "ยอด ยก", "ยก มา", "BALANCE BROUGHT FORWARD", "BROUGHT FORWARD", "BALANCE B/F", "B/F"],
        "credit_keywords": ["รับ โอน เงิน", "รับ โอน", "รับโอน", "ฝาก เงิน", "ฝาก", "ดอกเบี้ย เงิน ฝาก", "ดอกเบี้ย", "คืนเงิน", "เงิน เข้า", "เงินเข้า", "CREDIT", "DEPOSIT", "TRANSFER IN"],
        "debit_keywords": ["จ่าย บิล", "จ่าย คิว อา ร์", "จ่าย QR", "โอน เงิน พร้อม เพ ย์", "โอน เงิน พร้อมเพย์", "โอน เงิน", "โอนเงิน", "ถอนเงิน", "ถอน", "หัก", "ค่าธรรมเนียม", "ชำระ", "จ่าย", "DEBIT", "WITHDRAW", "WITHDRAWAL", "PAYMENT", "FEE"],
    },
    "BBL": {
        "opening_keywords": ["B/F", "B / F", "BALANCE B/F", "BALANCE BROUGHT FORWARD", "BROUGHT FORWARD", "ยอด ยก มา", "ยอดยกมา", "ยอด ยก", "ยก มา"],
        "credit_keywords": ["TRF FR OTH BK", "TRF FROTH BK", "TRF FROM OTH BK", "TRF FROM OTHER BANK", "TRANSFER IN", "SALARY", "SMART", "CREDIT", "DEPOSIT", "ฝาก", "รับ โอน", "รับโอน", "เงิน เข้า", "เงินเข้า"],
        "debit_keywords": ["TRF TO OTH BK", "TRF TOOTH BK", "TRF TO OTHER BANK", "TRF. PROMPTPAY", "TRF . PROMPTPAY", "PMT. PROMPTPAY", "PMT . PROMPTPAY", "PMT.PROMPTPAY", "PMT FOR GOODS", "CASH W/D ATM", "CASH W / D ATM", "WITHDRAWAL", "WITHDRAW", "TRANSFER", "PAYMENT", "DEBIT", "FEE", "ถอน", "ถอนเงิน", "โอน เงิน", "โอนเงิน", "ชำระ", "จ่าย"],
    },
    "SCB": {
        "opening_keywords": ["ยอดเงินคงเหลือยกมา", "ยอด เงิน คงเหลือ ยก มา", "BALANCE BROUGHT FORWARD", "BROUGHT FORWARD", "BALANCE B/F", "B/F"],
        "credit_keywords": [" X1 ", "X1", "รับ โอน จาก", "รับโอนจาก", "รับ โอน", "รับโอน", "โอน จาก", "โอนจาก", "ฝาก", "เงิน เข้า", "เงินเข้า", "CREDIT", "DEPOSIT", "TRANSFER IN"],
        "debit_keywords": [" X2 ", "X2", "โอน ไป", "โอนไป", "PromptPay", "PROMPTPAY", "จ่าย บิล", "จ่ายบิล", "จ่าย", "ถอน", "ถอนเงิน", "ชำระ", "ชา ระ", "ชํา ระ", "PAYMENT", "WITHDRAW", "WITHDRAWAL", "TRANSFER OUT", "FEE"],
    },
}

def get_bank_config(bank_name):
    return BANK_CONFIGS.get(bank_name, BANK_CONFIGS["DEFAULT"])

# ============================================================
# 2) Google Vision Client
# ============================================================
@st.cache_resource
def create_vision_client():
    try:
        key_data = dict(st.secrets["gcp_service_account"])
        creds = service_account.Credentials.from_service_account_info(key_data)
        return vision.ImageAnnotatorClient(credentials=creds)
    except Exception as e:
        st.error(f"❌ ไม่สามารถเชื่อมต่อ Google Vision API ได้: {e}\n(กรุณาตั้งค่า st.secrets ให้ถูกต้อง)")
        st.stop()

# ============================================================
# 3) Text & Math Helpers
# ============================================================
def clean_text(s):
    return re.sub(r"\s+", " ", str(s).replace("\n", " ")).strip()

def normalize_amount_text(s):
    s = re.sub(r"\s+", "", str(s).strip())
    ocr_map = {"O": "0", "o": "0", "I": "1", "l": "1", "|": "1", "S": "5", "s": "5", "B": "8", "g": "9", "q": "9", "Z": "2", "z": "2"}
    for char, digit in ocr_map.items(): s = s.replace(char, digit)
    if re.match(r"^\d{1,3}(,\d{3})+,\d{2}$", s): s = s[:-3] + "." + s[-2:]
    if re.match(r"^\d+,\d{2}$", s): s = s.replace(",", ".")
    if re.match(r"^\d{1,3}\.\d{3}\.\d{2}$", s):
        parts = s.split(".")
        s = parts[0] + "," + parts[1] + "." + parts[2]
    return s

def parse_money(s):
    if pd.isna(s): return None
    s = re.sub(r"[^0-9,.\-]", "", normalize_amount_text(s))
    if not s: return None
    try: return float(s.replace(",", ""))
    except Exception: return None

def is_money_token(s):
    s = normalize_amount_text(str(s).strip())
    patterns = [r"^-?\d{1,3}(,\d{3})*(\.\d{2})$", r"^-?\d+\.\d{2}$", r"^-?\d+,\d{2}$", r"^-?\d{1,3}\.\d{3}\.\d{2}$"]
    return any(re.match(p, s) for p in patterns)

def extract_money_values_loose(token):
    token = str(token).strip()
    ocr_map = {"O": "0", "o": "0", "I": "1", "l": "1", "|": "1", "S": "5", "s": "5", "B": "8", "g": "9", "q": "9", "Z": "2", "z": "2"}
    for char, digit in ocr_map.items(): token = token.replace(char, digit)
    priority_patterns = [
        r"(?<![\d,])-?\d{1,3}(?:,\d{3})+\.\d{2}(?!\d)", r"(?<![\d.])-?\d{1,3}(?:\.\d{3})+,\d{2}(?!\d)",
        r"(?<![\d.])-?\d{1,3}(?:\.\d{3})+\.\d{2}(?!\d)", r"(?<![\d,])-?\d+\.\d{2}(?!\d)", r"(?<![\d,])-?\d+,\d{2}(?![\d.])"
    ]
    for pattern in priority_patterns:
        matches = re.findall(pattern, token)
        if matches:
            values = [parse_money(m) for m in matches if parse_money(m) is not None]
            if values: return values
    return []

def score_money_candidate(text, value):
    text, score = str(text), 0
    if re.search(r"\d{1,3},\d{3}\.\d{2}", text): score += 40
    if re.search(r"\d+[.,]\d{2}", text): score += 20
    score += min(len(re.sub(r"\D", "", text)), 10)
    if value is not None and not pd.isna(value):
        if abs(float(value)) >= 1000: score += 8
        elif abs(float(value)) >= 100: score += 5
    return score

def pick_best_money_candidate(candidates):
    if not candidates: return None
    return sorted(candidates, key=lambda x: (score_money_candidate(x.get("text", ""), x.get("value")), x.get("x_center", 0)), reverse=True)[0].get("value")

def money_to_key(value):
    if value is None or pd.isna(value): return ""
    try: return str(int(round(float(value) * 100)))
    except: return ""

def try_healing_amount(prev_balance, read_amount, balance, tolerance=0.05):
    if any(v is None or pd.isna(v) for v in [prev_balance, balance]): return None, None
    expected_amount = round(prev_balance - balance, 2) if round(prev_balance - balance, 2) > 0 else round(balance - prev_balance, 2)
    tx_type = "debit" if round(prev_balance - balance, 2) > 0 else "credit"
    if read_amount is None or pd.isna(read_amount): return None, None
    str_read, str_exp = f"{read_amount:.2f}", f"{expected_amount:.2f}"
    
    if abs(read_amount / 100 - expected_amount) <= tolerance: return expected_amount, f"healed_missing_dot_{tx_type}"
    if len(str_read) == len(str_exp) and sum(1 for a, b in zip(str_read, str_exp) if a != b) <= 1: return expected_amount, f"healed_swapped_digit_{tx_type}"
    if len(str_exp) > len(str_read) and (str_read in str_exp or str_exp.endswith(str_read)): return expected_amount, f"healed_omitted_digit_{tx_type}"
    return None, None

def amount_balance_match(prev_balance, amount, balance, tolerance=BALANCE_TOLERANCE):
    if any(v is None or pd.isna(v) for v in [prev_balance, amount, balance]): return None
    d_exp, c_exp = round(prev_balance - amount, 2), round(prev_balance + amount, 2)
    if abs(balance - d_exp) <= tolerance: return {"type": "debit", "expected_balance": d_exp, "diff": round(balance - d_exp, 2), "balance_check": "OK_DEBIT"}
    if abs(balance - c_exp) <= tolerance: return {"type": "credit", "expected_balance": c_exp, "diff": round(balance - c_exp, 2), "balance_check": "OK_CREDIT"}
    return None

def suggest_amount_from_balance(prev_balance, balance):
    if any(v is None or pd.isna(v) for v in [prev_balance, balance]): return None, None
    diff = round(balance - prev_balance, 2)
    if diff > 0: return abs(diff), "credit"
    if diff < 0: return abs(diff), "debit"
    return 0.0, "zero"

# ============================================================
# 4) Date & Transaction Helpers
# ============================================================
def fix_ocr_date_token(token):
    return str(token).strip().replace("O", "0").replace("o", "0").replace("I", "1").replace("l", "1").replace(".", "-").replace("/", "-")

def is_date_token(s):
    return any(re.match(p, fix_ocr_date_token(s)) for p in [r"^\d{1,2}[-]\d{1,2}[-]\d{2,4}$", r"^\d{1,2}[-]\d{1,2}$"])

def normalize_date(s, force_year):
    s = fix_ocr_date_token(s)
    m1, m2 = re.match(r"^(\d{1,2})[-](\d{1,2})[-](\d{2,4})$", s), re.match(r"^(\d{1,2})[-](\d{1,2})$", s)
    if m1: day, month = int(m1.group(1)), int(m1.group(2))
    elif m2: day, month = int(m2.group(1)), int(m2.group(2))
    else: return None
    try: return datetime(force_year, month, day).strftime("%d-%m-%Y")
    except: return None

def extract_date_from_line(line_text, force_year):
    for t in str(line_text).split():
        if is_date_token(t): return normalize_date(t, force_year)
    return None

def parse_date_for_sort(date_text):
    if pd.isna(date_text) or not str(date_text).strip(): return pd.NaT
    return pd.to_datetime(str(date_text).strip().replace("/", "-"), format="%d-%m-%Y", errors="coerce")

def is_time_token(s): return bool(re.match(r"^\d{1,2}:\d{2}(:\d{2})?$", str(s).strip()))

def extract_time_from_line(line_text):
    for t in str(line_text).split():
        if is_time_token(t): return t
    return None

def join_words(g): return clean_text(" ".join(g.sort_values("x_center")["text"].astype(str).tolist())) if not g.empty else ""

def extract_money_from_line_words(g):
    items = []
    for _, row in g.sort_values("x_center").iterrows():
        text = normalize_amount_text(str(row["text"]))
        if is_money_token(text) or extract_money_values_loose(text):
            val = parse_money(text)
            if val is not None: items.append({"text": text, "value": val, "x_center": row["x_center"], "confidence": row.get("word_confidence")})
    return items

def has_opening_balance_text(text, bank_name):
    text = clean_text(text).upper()
    for kw in get_bank_config(bank_name).get("opening_keywords", []) + BANK_CONFIGS["DEFAULT"]["opening_keywords"]:
        if kw.upper() in text: return True
    return False

def classify_by_keyword(text, bank_name):
    text = clean_text(text).upper()
    cfg = get_bank_config(bank_name)
    for kw in cfg.get("credit_keywords", []) + BANK_CONFIGS["DEFAULT"]["credit_keywords"]:
        if kw.upper() in text: return "credit"
    for kw in cfg.get("debit_keywords", []) + BANK_CONFIGS["DEFAULT"]["debit_keywords"]:
        if kw.upper() in text: return "debit"
    return "unknown"

def is_noise_line(text):
    text_u = clean_text(text).upper()
    noise = ["STATEMENT OF SAVING ACCOUNT", "THE SIAM COMMERCIAL BANK", "ACCOUNT NO", "เลขที่บัญชี", "ชื่อ - สกุล", "ADDRESS", "สาขา", "BRANCH", "DATE | TIME", "DEBIT/CREDIT", "BALANCE/BAHT", "DESCRIPTION/NOTE", "TOTAL AMOUNTS", "TOTAL ITEMS", "THIS DOCUMENT IS AUTO-GENERATED", "เอกสาร ฉบับ นี้", "ออก โดย ระบบ อัตโนมัติ"]
    if any(kw in text_u for kw in noise): return True
    if re.search(r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\s*[-–]\s*\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", text_u): return True
    if re.search(r"หน้า\s*\d+\s*/\s*\d+", text_u): return True
    return False

def extract_page_no_from_text(text):
    text = clean_text(text)
    m = re.search(r"หน้า\s*(\d+)\s*/\s*(\d+)", text)
    if m: return int(m.group(1))
    m = re.search(r"\b(\d+)\s*/\s*(\d+)\b", text)
    if m and 1 <= int(m.group(1)) <= int(m.group(2)) <= 20: return int(m.group(1))
    return None

def infer_page_no_map_from_line_words(df):
    page_map = {}
    if df.empty: return page_map
    for f_name, fg in df.groupby("file_name"):
        found = None
        for _, lg in fg.groupby("line_id"):
            page_no = extract_page_no_from_text(join_words(lg))
            if page_no is not None:
                found = page_no
                break
        if found is not None: page_map[f_name] = found
    return page_map

# ============================================================
# 5) SCB Specific
# ============================================================
def extract_scb_code_from_line(text):
    for token in re.split(r"\s+", clean_text(text).upper()):
        if token.strip("|:;,.()[]{}") in ["X1", "X2"]: return token.strip("|:;,.()[]{}")
    return None

def extract_scb_money_by_column(g):
    if g.empty or "page_width" not in g.columns: return {"amount": None, "balance": None, "money_texts": []}
    pw = g["page_width"].dropna().iloc[0]
    a_min, a_max, b_min, b_max = pw * 0.25, pw * 0.52, pw * 0.52, pw * 0.70
    a_cands, b_cands, m_texts = [], [], []
    for _, row in g.sort_values("x_center").iterrows():
        x, t = row["x_center"], str(row["text"])
        if a_min <= x < a_max:
            for v in extract_money_values_loose(t): a_cands.append({"value": v, "x_center": x, "text": t}); m_texts.append(t)
        elif b_min <= x < b_max:
            for v in extract_money_values_loose(t): b_cands.append({"value": v, "x_center": x, "text": t}); m_texts.append(t)
    return {"amount": pick_best_money_candidate(a_cands), "balance": pick_best_money_candidate(b_cands), "money_texts": m_texts}

def parse_scb_line(g, active_bank, force_year, page_no=None):
    g = g.sort_values("x_center").copy()
    raw_line = join_words(g)
    if is_noise_line(raw_line): return None
    date, time = extract_date_from_line(raw_line, force_year), extract_time_from_line(raw_line)
    l_conf = float(g["word_confidence"].mean()) if "word_confidence" in g.columns else None
    m_conf = float(g["word_confidence"].min()) if "word_confidence" in g.columns else None
    
    if has_opening_balance_text(raw_line, active_bank):
        m_info = extract_scb_money_by_column(g)
        bal = m_info.get("balance")
        if bal is None:
            m_items = extract_money_from_line_words(g)
            bal = m_items[-1]["value"] if m_items else None
            m_texts = [m["text"] for m in m_items]
        else: m_texts = m_info.get("money_texts", [])
        if bal is None: return None
        return {"file_name": g["file_name"].iloc[0], "line_id": int(g["line_id"].iloc[0]), "page_no": page_no, "date": None, "time": None, "code": "OPENING", "amount": None, "balance": bal, "money_tokens": " | ".join(m_texts), "money_count": len(m_texts), "raw_line_text": raw_line, "y_center": g["y_center"].mean(), "parsed_date": pd.NaT, "is_opening_balance": True, "line_confidence": l_conf, "min_confidence": m_conf}

    code = extract_scb_code_from_line(raw_line)
    if not date or code not in ["X1", "X2"]: return None
    m_info = extract_scb_money_by_column(g)
    return {"file_name": g["file_name"].iloc[0], "line_id": int(g["line_id"].iloc[0]), "page_no": page_no, "date": date, "time": time, "code": code, "amount": m_info.get("amount"), "balance": m_info.get("balance"), "money_tokens": " | ".join(m_info.get("money_texts", [])), "money_count": len(m_info.get("money_texts", [])), "raw_line_text": raw_line, "y_center": g["y_center"].mean(), "parsed_date": parse_date_for_sort(date), "is_opening_balance": False, "line_confidence": l_conf, "min_confidence": m_conf}

# ============================================================
# 6) Image Preprocessing & OCR
# ============================================================
def get_box_info(vertices):
    xs, ys = [v.x for v in vertices], [v.y for v in vertices]
    return {"x_min": min(xs), "y_min": min(ys), "x_max": max(xs), "y_max": max(ys), "x_center": sum(xs)/len(xs), "y_center": sum(ys)/len(ys), "width": max(xs)-min(xs), "height": max(ys)-min(ys)}

def preprocess_image_variants(input_path, tmp_dir):
    img = Image.open(input_path).convert("RGB")
    b_name = os.path.splitext(os.path.basename(input_path))[0]
    def _save(img_obj, suffix):
        p = os.path.join(tmp_dir, f"pre_{b_name}_{suffix}.png")
        img_obj.save(p); return p

    v = [("original", input_path)]
    i1 = ImageEnhance.Sharpness(ImageEnhance.Contrast(img.resize((img.width*2, img.height*2), Image.LANCZOS)).enhance(1.9)).enhance(2.2)
    v.append(("2x_contrast_sharp", _save(i1, "2x_contrast_sharp")))
    
    i2 = ImageEnhance.Sharpness(ImageEnhance.Contrast(ImageOps.autocontrast(img.resize((img.width*3, img.height*3), Image.LANCZOS).convert("L"))).enhance(1.8)).enhance(2.0)
    v.append(("3x_gray_autocontrast", _save(i2, "3x_gray_autocontrast")))
    
    i3 = ImageOps.autocontrast(img.resize((img.width*3, img.height*3), Image.LANCZOS).convert("L")).filter(ImageFilter.UnsharpMask(radius=1.2, percent=180, threshold=3))
    v.append(("3x_unsharp", _save(i3, "3x_unsharp")))
    
    i4 = ImageEnhance.Contrast(ImageOps.autocontrast(img.resize((img.width*3, img.height*3), Image.LANCZOS).convert("L"))).enhance(2.0).point(lambda x: 255 if x > 175 else 0)
    v.append(("3x_bw_threshold_175", _save(i4, "3x_bw_threshold_175")))
    return v

def vision_ocr_image(image_path, client):
    with open(image_path, "rb") as f: content = f.read()
    response = client.document_text_detection(image=vision.Image(content=content), image_context=vision.ImageContext(language_hints=["th", "en"]))
    if response.error.message: raise Exception(response.error.message)
    rows = []
    for p_idx, page in enumerate(response.full_text_annotation.pages):
        for b_idx, block in enumerate(page.blocks):
            for pr_idx, paragraph in enumerate(block.paragraphs):
                for w_idx, word in enumerate(paragraph.words):
                    wb = get_box_info(word.bounding_box.vertices)
                    rows.append({"file_name": os.path.basename(image_path), "page_idx": p_idx, "page_width": page.width, "page_height": page.height, "block_idx": b_idx, "paragraph_idx": pr_idx, "word_idx": w_idx, "text": "".join([s.text for s in word.symbols]), "word_confidence": word.confidence, **wb})
    return response.full_text_annotation.text, pd.DataFrame(rows)

def ocr_best_variant_for_image(image_path, client, active_bank, force_year, tmp_dir, status_box=None):
    variants = preprocess_image_variants(image_path, tmp_dir)
    best, best_score = None, None
    for v_name, v_path in variants:
        try:
            full_text, word_df = vision_ocr_image(v_path, client)
            if word_df.empty: continue
            word_df["file_name"], word_df["ocr_variant"] = os.path.basename(image_path), v_name
            parsed_preview = parse_simple_stm(word_df, active_bank, force_year)
            score = (len(parsed_preview)*100 + parsed_preview["money_count"].fillna(0).sum()*20 + parsed_preview["date"].notna().sum()*50 + (parsed_preview["line_confidence"].fillna(0).mean()*100 if "line_confidence" in parsed_preview else 0)) if not parsed_preview.empty else -999
            if best_score is None or score > best_score:
                best_score = score
                best = {"variant_name": v_name, "word_df": word_df}
        except Exception: pass
    if best is None: raise Exception(f"OCR พังบนรูป {image_path}")
    if status_box: status_box.write(f"✓ เลือกความชัดระดับ: {best['variant_name']} สำหรับไฟล์ {os.path.basename(image_path)}")
    return best

# ============================================================
# 7) Build Lines & Parse
# ============================================================
def rebuild_lines_by_y(words_df, y_threshold=10):
    all_lines = []
    for f_name, g in words_df.groupby("file_name"):
        if g.empty: continue
        g = g[g["x_center"] >= g["page_width"].dropna().iloc[0] * 0.04].sort_values(["y_center", "x_center"]).reset_index(drop=True)
        cur_line, cur_y, line_id = [], None, 0
        for _, r in g.iterrows():
            y = r["y_center"]
            if cur_y is None: cur_line, cur_y = [r], y; continue
            if abs(y - cur_y) <= y_threshold: cur_line.append(r); cur_y = (cur_y + y) / 2
            else:
                cur_line_df = pd.DataFrame(cur_line); cur_line_df["line_id"] = line_id; all_lines.append(cur_line_df)
                line_id += 1; cur_line, cur_y = [r], y
        if cur_line:
            cur_line_df = pd.DataFrame(cur_line); cur_line_df["line_id"] = line_id; all_lines.append(cur_line_df)
    return pd.concat(all_lines, ignore_index=True) if all_lines else pd.DataFrame()

def parse_simple_stm(words_df, active_bank, force_year):
    line_df = rebuild_lines_by_y(words_df, 10)
    if line_df.empty: return pd.DataFrame()
    rows = []
    page_map = infer_page_no_map_from_line_words(line_df)
    for (f_name, l_id), g in line_df.groupby(["file_name", "line_id"]):
        g = g.sort_values("x_center").copy()
        p_no = page_map.get(f_name)
        if active_bank == "SCB":
            pl = parse_scb_line(g, active_bank, force_year, p_no)
            if pl: rows.append(pl)
            continue
        
        raw_text = join_words(g)
        if is_noise_line(raw_text): continue
        date, time = extract_date_from_line(raw_text, force_year), extract_time_from_line(raw_text)
        if not date and not has_opening_balance_text(raw_text, active_bank): continue
        
        m_items = extract_money_from_line_words(g)
        m_vals, m_texts = [m["value"] for m in m_items], [m["text"] for m in m_items]
        l_conf = float(g["word_confidence"].mean()) if "word_confidence" in g.columns else None
        m_conf = float(g["word_confidence"].min()) if "word_confidence" in g.columns else None
        
        amt, bal = None, None
        if has_opening_balance_text(raw_text, active_bank):
            if len(m_vals) >= 1: bal = m_vals[-1]
        else:
            if len(m_vals) >= 2: amt, bal = m_vals[0], m_vals[-1]
            elif len(m_vals) == 1: amt = m_vals[0]
            
        rows.append({"file_name": f_name, "line_id": int(l_id), "page_no": p_no, "date": date, "time": time, "code": None, "amount": amt, "balance": bal, "money_tokens": " | ".join(m_texts), "money_count": len(m_vals), "raw_line_text": raw_text, "y_center": g["y_center"].mean(), "parsed_date": parse_date_for_sort(date), "is_opening_balance": has_opening_balance_text(raw_text, active_bank), "line_confidence": l_conf, "min_confidence": m_conf})
        
    pdf = pd.DataFrame(rows)
    if pdf.empty: return pdf
    sort_cols = ["page_no"] if "page_no" in pdf.columns and pdf["page_no"].notna().any() else []
    return pdf.sort_values(sort_cols + ["file_name", "y_center"], na_position="last").reset_index(drop=True)

# ============================================================
# 8) Advanced Balance Chain
# ============================================================
def advanced_build_chain_with_healing(work_df, start_index, tolerance=0.05):
    cands = work_df.copy().reset_index(drop=True)
    start_row = cands.loc[start_index].to_dict()
    cands = cands.drop(index=start_index).reset_index(drop=True)
    ord_rows = []
    match_count = 0
    start_row["order_method"] = "CHAIN_START"
    ord_rows.append(start_row)
    prev_bal = start_row.get("balance")
    
    while not cands.empty:
        f_idx, f_match, is_healed, h_info = None, None, False, {}
        for idx, row in cands.iterrows():
            m = amount_balance_match(prev_bal, row.get("amount"), row.get("balance"), tolerance)
            if m: f_idx, f_match = idx, m; break
        if f_idx is None and prev_bal is not None:
            for idx, row in cands.iterrows():
                ha, ht = try_healing_amount(prev_bal, row.get("amount"), row.get("balance"), tolerance)
                if ha is not None: f_idx, is_healed, h_info = idx, True, {"amount": ha, "type": ht}; break
        if f_idx is None: break
        
        nxt = cands.loc[f_idx].to_dict()
        if is_healed:
            nxt["amount"], nxt["order_method"], nxt["chain_type"] = h_info["amount"], f"BALANCE_CHAIN_{h_info['type'].upper()}", "debit" if "debit" in h_info["type"] else "credit"
        else:
            nxt["order_method"], nxt["chain_type"] = "BALANCE_CHAIN_PERFECT", f_match["type"]
            
        ord_rows.append(nxt)
        match_count += 1
        if nxt.get("balance") is not None and not pd.isna(nxt.get("balance")): prev_bal = nxt.get("balance")
        cands = cands.drop(index=f_idx).reset_index(drop=True)
    return pd.DataFrame(ord_rows), cands, match_count

def smart_order_by_balance_chain(df, active_bank):
    if df.empty: return df
    if active_bank == "SCB":
        sc = ["page_no"] if "page_no" in df.columns and df["page_no"].notna().any() else []
        sc.extend(["file_name", "y_center"])
        res = df.copy().sort_values(sc, na_position="last").reset_index(drop=True)
        res["order_method"] = "SCB_PAGE_LAYOUT"
        return res

    work = df.copy().reset_index(drop=True)
    op_df = work[work["is_opening_balance"] == True].copy()
    tx_df = work[work["is_opening_balance"] == False].copy()
    
    starts = []
    if not op_df.empty:
        for _, r in op_df.reset_index(drop=True).iterrows():
            starts.append({"df": pd.concat([pd.DataFrame([r]), tx_df], ignore_index=True), "start_index": 0, "start_type": "OPENING_START"})
    tx_temp = tx_df.reset_index(drop=True)
    for idx in range(len(tx_temp)): starts.append({"df": tx_temp, "start_index": idx, "start_type": "TX_START"})
    
    best_ord, best_rem, best_score = None, None, -999999
    for opt in starts:
        ord_df, rem_df, m_count = advanced_build_chain_with_healing(opt["df"].copy(), opt["start_index"], BALANCE_TOLERANCE)
        score = m_count * 1000 + (500 if opt["start_type"] == "OPENING_START" else 0) - len(rem_df) * 50 + (ord_df["parsed_date"].notna().sum() if "parsed_date" in ord_df.columns else 0)
        if score > best_score: best_score, best_ord, best_rem = score, ord_df.copy(), rem_df.copy()
        
    if best_ord is None: return work
    if best_rem is not None and not best_rem.empty:
        sf = ["page_no"] if "page_no" in best_rem.columns and best_rem["page_no"].notna().any() else []
        sf.extend(["parsed_date", "time", "y_center"])
        best_rem = best_rem.sort_values(by=sf, na_position="last").reset_index(drop=True)
        best_rem["order_method"] = "CHAIN_BROKEN_FALLBACK_REVIEW"
        return pd.concat([best_ord, best_rem], ignore_index=True).reset_index(drop=True)
    return best_ord.reset_index(drop=True)

# ============================================================
# 9) Deduplication & Debit/Credit Check Builders
# ============================================================
def build_duplicate_key(row):
    t_key, a_key, b_key, raw = str(row.get("time", "")).strip(), money_to_key(row.get("amount")), money_to_key(row.get("balance")), str(row.get("raw_line_text", ""))
    if "ยอดยกมา" in raw.lower() or "brought forward" in raw.upper() or "ยอด ยก มา" in raw: return "OPENING|" + b_key
    if t_key: return f"TX|{t_key}|{a_key}|{b_key}"
    return f"TX|{a_key}|{b_key}"

def deduplicate_rows(df):
    if df is None or df.empty: return df
    df = df.copy()
    df["duplicate_key"] = df.apply(build_duplicate_key, axis=1)
    return df.drop_duplicates(subset=["duplicate_key"], keep="first").drop(columns=["duplicate_key"]).reset_index(drop=True)

def add_debit_credit_and_check(ordered_df, active_bank):
    if ordered_df is None or ordered_df.empty: return pd.DataFrame()
    rows, prev_bal, seq = [], None, 0
    for _, r in ordered_df.iterrows():
        seq += 1
        date, time, code, amt, bal, raw, mt, om, l_conf, m_conf = r.get("date"), r.get("time"), r.get("code"), r.get("amount"), r.get("balance"), r.get("raw_line_text", ""), r.get("money_tokens", ""), r.get("order_method", ""), r.get("line_confidence"), r.get("min_confidence")
        deb = cred = exp_bal = diff = sug_amt = sug_type = None
        chk = ""
        low_c = (l_conf is not None and not pd.isna(l_conf) and l_conf < LOW_CONFIDENCE_THRESHOLD) or (m_conf is not None and not pd.isna(m_conf) and m_conf < LOW_CONFIDENCE_THRESHOLD)
        
        if has_opening_balance_text(raw, active_bank) or code == "OPENING": chk = "OPENING_BALANCE"
        elif active_bank == "SCB" and code in ["X1", "X2"]:
            if prev_bal is None: chk = "SCB_NO_PREV_BALANCE_REVIEW"
            elif bal is None or pd.isna(bal): chk = "SCB_NO_BALANCE_REVIEW"
            elif amt is None or pd.isna(amt):
                if code == "X1": sug_amt, sug_type, cred, exp_bal, diff, chk = round(bal - prev_bal, 2), "credit", round(bal - prev_bal, 2), bal, 0.0, "SCB_X1_NO_AMOUNT_SUGGEST_CREDIT_REVIEW"
                else: sug_amt, sug_type, deb, exp_bal, diff, chk = round(prev_bal - bal, 2), "debit", round(prev_bal - bal, 2), bal, 0.0, "SCB_X2_NO_AMOUNT_SUGGEST_DEBIT_REVIEW"
            else:
                if code == "X1":
                    exp_bal, diff = round(prev_bal + amt, 2), round(bal - round(prev_bal + amt, 2), 2)
                    if abs(diff) <= BALANCE_TOLERANCE: cred, chk = amt, "OK_CREDIT"
                    else:
                        ha, ht = try_healing_amount(prev_bal, amt, bal, BALANCE_TOLERANCE)
                        if ha is not None: cred, exp_bal, diff, chk = ha, bal, 0.0, f"OK_{ht.upper()}"
                        else: cred, chk = amt, "SCB_X1_SUGGEST_AMOUNT_REVIEW"
                else:
                    exp_bal, diff = round(prev_bal - amt, 2), round(bal - round(prev_bal - amt, 2), 2)
                    if abs(diff) <= BALANCE_TOLERANCE: deb, chk = amt, "OK_DEBIT"
                    else:
                        ha, ht = try_healing_amount(prev_bal, amt, bal, BALANCE_TOLERANCE)
                        if ha is not None: deb, exp_bal, diff, chk = ha, bal, 0.0, f"OK_{ht.upper()}"
                        else: deb, chk = "SCB_X2_SUGGEST_AMOUNT_REVIEW"
        else:
            m = amount_balance_match(prev_bal, amt, bal, BALANCE_TOLERANCE)
            if m:
                exp_bal, diff, chk = m["expected_balance"], m["diff"], m["balance_check"]
                if m["type"] == "debit": deb = amt
                else: cred = amt
            else:
                ha, ht = try_healing_amount(prev_bal, amt, bal, BALANCE_TOLERANCE)
                if ha is not None:
                    if "debit" in ht: deb = ha
                    else: cred = ha
                    exp_bal, diff, chk = bal, 0.0, f"OK_{ht.upper()}"
                else:
                    tx_type, (sug_amt, sug_type) = classify_by_keyword(raw, active_bank), suggest_amount_from_balance(prev_bal, bal)
                    if sug_amt is not None and sug_type in ["debit", "credit"]:
                        exp_bal = round(prev_bal - sug_amt if sug_type == "debit" else prev_bal + sug_amt, 2)
                        diff = round(bal - exp_bal, 2)
                        if tx_type == sug_type:
                            chk = "SUGGEST_AMOUNT_BY_KEYWORD_REVIEW"
                            if sug_type == "debit": deb = sug_amt
                            else: cred = sug_amt
                        else: chk = "SUGGEST_AMOUNT_REVIEW"
                    else: chk = "CHAIN_BROKEN_OCR_REVIEW"
        
        if low_c and "OK" in chk: chk += "_LOW_CONFIDENCE"
        rows.append({"seq": seq, "date": date, "time": time, "code": code, "debit": deb, "credit": cred, "balance": bal, "prev_balance": prev_bal, "amount": amt, "suggested_amount": sug_amt, "suggested_type": sug_type, "expected_balance": exp_bal, "diff": diff, "line_confidence": l_conf, "min_confidence": m_conf, "money_tokens": mt, "raw_line_text": raw, "order_method": om, "balance_check": chk})
        if bal is not None and not pd.isna(bal): prev_bal = bal
    return pd.DataFrame(rows)

def is_valid_num(x):
    if x is None: return False
    try: return not pd.isna(x) and float(x) is not None
    except: return False

def build_final_parsed_stm_from_check(df):
    if df is None or df.empty: return pd.DataFrame(columns=["date", "debit", "credit", "balance"])
    df = df.copy()
    op_mask = df.apply(lambda r: "ยอดยกมา" in str(r.get("raw_line_text", "")).lower() or "brought forward" in str(r.get("raw_line_text", "")).upper() or "ยอด ยก มา" in str(r.get("raw_line_text", "")) or str(r.get("code", "")) == "OPENING" or str(r.get("balance_check", "")) == "OPENING_BALANCE", axis=1)
    df = df[~op_mask].copy()
    df = df[df["balance"].apply(is_valid_num)].copy()
    df = df[(df["debit"].apply(is_valid_num)) | (df["credit"].apply(is_valid_num))].copy()
    if "seq" in df.columns: df = df.sort_values("seq")
    res = df[["date", "debit", "credit", "balance"]].copy()
    for c in ["debit", "credit", "balance"]: res[c] = pd.to_numeric(res[c], errors="coerce")
    return res.reset_index(drop=True)

# ============================================================
# 10) Exact Excel Export Logic (Matched to Colab 100%)
# ============================================================
def create_summary_df(final_df, check_df, detected_bank):
    if final_df.empty: return pd.DataFrame(columns=["metric", "value"])
    db_s, cr_s = pd.to_numeric(final_df["debit"]), pd.to_numeric(final_df["credit"])
    return pd.DataFrame([
        ["detected_bank", detected_bank], 
        ["transaction_rows", int(final_df["date"].notna().sum())],
        ["debit_total", float(db_s.fillna(0).sum())], 
        ["credit_total", float(cr_s.fillna(0).sum())],
        ["net_change", float(cr_s.fillna(0).sum() - db_s.fillna(0).sum())],
        ["ok_checks", int(check_df["balance_check"].astype(str).str.contains("OK").sum())],
        ["anomalies_found", int(check_df["balance_check"].astype(str).str.contains("REVIEW|BROKEN|MISSING").sum())]
    ], columns=["metric", "value"])

def prepare_check_export_df(check_df):
    if check_df.empty: return check_df
    # ลบคอลัมน์ที่ไม่ได้ใช้งานในตารางตรวจสอบ เพื่อให้ดูสะอาดตา เหมือนกับ Colab เป๊ะ
    df = check_df.drop(columns=["amount", "line_confidence", "min_confidence", "money_tokens", "raw_line_text", "balance_check"], errors="ignore")
    df = df.rename(columns={"expected_balance": "expected_balance_py", "diff": "diff_py"})
    df["expected_balance_excel"], df["anomaly_reason_excel"] = None, None
    return df

def apply_check_sheet_formulas_and_format(output_excel, sheet_name="check"):
    wb = load_workbook(output_excel)
    if sheet_name not in wb.sheetnames: return
    ws = wb[sheet_name]
    headers = {str(cell.value): cell.column for cell in ws[1] if cell.value}
    
    req = ["debit", "credit", "balance", "prev_balance", "suggested_amount", "expected_balance_excel", "anomaly_reason_excel"]
    for c in req:
        if c not in headers: return
        
    d_c, c_c, b_c, p_c, s_c, e_c, r_c = [get_column_letter(headers[x]) for x in req]
    tol = BALANCE_TOLERANCE

    for row in range(2, ws.max_row + 1):
        ws[f"{e_c}{row}"] = f'=IF({p_c}{row}="","",ROUND({p_c}{row}-IF({d_c}{row}="",0,{d_c}{row})+IF({c_c}{row}="",0,{c_c}{row}),2))'
        ws[f"{r_c}{row}"] = (f'=IF({p_c}{row}="","OPENING_OR_NO_PREV",IF({b_c}{row}="","MISSING_BALANCE",IF({s_c}{row}<>"","กระทบยอดให้ตรวจสอบ",IF(AND({d_c}{row}="",{c_c}{row}=""),"MISSING_DEBIT_CREDIT",IF(ABS({b_c}{row}-{e_c}{row})>{tol},"BALANCE_NOT_MATCH","OK")))))')

    h_fill, h_font, thin = PatternFill("solid", fgColor="D9EAF7"), Font(color="000000", bold=True), Side(style="thin", color="D9E2F3")
    for cell in ws[1]: cell.fill, cell.font, cell.alignment, cell.border = h_fill, h_font, Alignment(horizontal="center", vertical="center"), Border(bottom=thin)
    
    for col in ["debit", "credit", "balance", "prev_balance", "suggested_amount", "expected_balance_py", "diff_py", "expected_balance_excel"]:
        if col in headers:
            for row in range(2, ws.max_row + 1): ws.cell(row=row, column=headers[col]).number_format = '#,##0.00'
            
    ws.freeze_panes = "A2"
    ws.conditional_formatting.add(f"A2:{get_column_letter(ws.max_column)}{ws.max_row}", FormulaRule(formula=[f'=AND(${r_c}2<>"OK",${r_c}2<>"OPENING_OR_NO_PREV",${r_c}2<>"")'], fill=PatternFill("solid", fgColor="FCE4E4"), font=Font(color="9C0006")))
    
    for c_idx in range(1, ws.max_column + 1):
        cl = get_column_letter(c_idx)
        max_len = max([len(str(cell.value or '')) for cell in ws[cl]])
        ws.column_dimensions[cl].width = min(max(max_len + 2, 12), 35)
    wb.save(output_excel)

def build_excel_bytes(final_df, check_df, selected_bank, tmp_dir):
    sum_df = create_summary_df(final_df, check_df, selected_bank)
    ocr_rev_df = check_df[check_df["balance_check"].astype(str).str.contains("REVIEW|BROKEN|UNKNOWN|NO_PREV|MISSING|SUGGEST|HEALED", regex=True, na=False)].copy().reset_index(drop=True)
    check_export = prepare_check_export_df(check_df)
    
    out_path = os.path.join(tmp_dir, "output.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="parsed_stm", index=False)
        check_export.to_excel(writer, sheet_name="check", index=False)
        sum_df.to_excel(writer, sheet_name="summary", index=False)
        ocr_rev_df.to_excel(writer, sheet_name="ocr_review", index=False)
        
    apply_check_sheet_formulas_and_format(out_path, "check")
    with open(out_path, "rb") as f: bytes_data = f.read()
    return bytes_data, sum_df, ocr_rev_df

# ============================================================
# 11) Streamlit UI & Main Logic
# ============================================================
with st.sidebar:
    st.header("⚙️ ตั้งค่าการประมวลผล")
    selected_bank = st.selectbox("เลือกธนาคาร", ["KBANK", "KRUNGSRI", "BBL", "SCB", "DEFAULT"], help="เลือกระบบธนาคารที่ตรงกับ Statement")
    selected_year = st.number_input("ปีของ Statement", min_value=2000, max_value=2100, value=datetime.now().year, help="กรณีที่ปีไม่มีในภาพ ระบบจะใช้ปีที่กำหนดนี้")
    st.divider()
    st.caption("🔐 Credentials load from st.secrets")

uploaded_images = st.file_uploader("🖼️ อัปโหลดรูป Statement", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
process_btn = st.button("🚀 เริ่มประมวลผล", type="primary", use_container_width=True)

if process_btn:
    if not uploaded_images:
        st.warning("⚠️ กรุณาอัปโหลดรูป Statement ก่อนกดเริ่มทำงาน")
        st.stop()

    client = create_vision_client()
    tmp_dir = tempfile.mkdtemp()
    
    try:
        all_words = []
        p_bar = st.progress(0)
        log_box = st.container()

        with st.spinner(f"กำลังแยกแยะข้อความและคำนวณ Balance Chain สำหรับ {selected_bank}..."):
            for i, img_f in enumerate(uploaded_images):
                ext = os.path.splitext(img_f.name)[1].lower() or ".png"
                i_path = os.path.join(tmp_dir, f"input_{i}{ext}")
                with open(i_path, "wb") as f: f.write(img_f.read())
                
                best_v = ocr_best_variant_for_image(i_path, client, selected_bank, int(selected_year), tmp_dir, status_box=log_box)
                best_v["word_df"]["file_name"] = img_f.name
                all_words.append(best_v["word_df"])
                p_bar.progress((i + 1) / len(uploaded_images))

        if not all_words:
            st.error("❌ ไม่พบข้อความที่สามารถอ่านได้จากรูปภาพ")
            st.stop()

        words_raw_df = pd.concat(all_words, ignore_index=True)
        
        parsed_df = parse_simple_stm(words_raw_df, selected_bank, int(selected_year))
        dedup_df = deduplicate_rows(parsed_df)
        ordered_df = smart_order_by_balance_chain(dedup_df, selected_bank)
        check_df = add_debit_credit_and_check(ordered_df, selected_bank)
        final_df = build_final_parsed_stm_from_check(check_df)

        excel_bytes, summary_df, ocr_review_df = build_excel_bytes(final_df, check_df, selected_bank, tmp_dir)

        st.session_state.update({
            "ready": True, "final": final_df, "check": check_df,
            "summary": summary_df, "bytes": excel_bytes, "bank": selected_bank
        })

    except Exception as e: st.exception(e)
    finally: shutil.rmtree(tmp_dir, ignore_errors=True)

# ============================================================
# 12) Output Display
# ============================================================
if st.session_state.get("ready"):
    st.success("✅ สร้างไฟล์ Excel สมบูรณ์แบบเรียบร้อยแล้ว!")
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Rows (ใช้งานได้)", len(st.session_state["final"]))
    c2.metric("OK Checks", int(st.session_state["check"]["balance_check"].astype(str).str.contains("OK", na=False).sum()))
    c3.metric("Anomalies Found", int(st.session_state["check"]["balance_check"].astype(str).str.contains("REVIEW|BROKEN|MISSING", regex=True, na=False).sum()))
    c4.metric("Bank Processed", st.session_state["bank"])

    st.subheader("📊 ข้อมูลพร้อมใช้งาน (parsed_stm)")
    st.dataframe(st.session_state["final"], use_container_width=True)

    st.download_button(
        label="⬇️ ดาวน์โหลดไฟล์ Excel (สมบูรณ์ 100%)",
        data=st.session_state["bytes"],
        file_name=f"stm_image_{st.session_state['bank']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
