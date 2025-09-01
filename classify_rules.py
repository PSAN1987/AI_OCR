# classify_rules.py
# OCRテキストから分類・項目抽出・命名を行うロジックをまとめる

import re
from datetime import datetime

# 分類ルール
CAT_RULES = [
    ("同意書",     r"同意|承諾|Consent|署名|サイン"),
    ("保険証",     r"保険証|被保険者|保険者番号|記号|番号|有効期限"),
    ("治療報告書", r"治療報告書|報告書|診療報酬|所見|診断|経過"),
    ("患者リスト", r"患者リスト|患者一覧|Patient List|患者台帳"),
    ("実績",       r"治療実績|実績|処置|施術|点数|算定"),
    ("請求書",     r"請求書|請求(?!先)|請求金額|振込先|INVOICE|請求合計"),
]

DATE_PATTERNS = [
    r"(20\d{2})[./年-](\d{1,2})[./月-](\d{1,2})日?",
    r"(20\d{2})-(\d{1,2})-(\d{1,2})",
    r"(20\d{2})/(\d{1,2})/(\d{1,2})",
]

def detect_category(text: str) -> str:
    t = text.replace("　", " ")
    for name, pat in CAT_RULES:
        if re.search(pat, t):
            return name
    return "その他"

def extract_date(text: str):
    for p in DATE_PATTERNS:
        m = re.search(p, text)
        if m:
            y, mo, d = m.groups()
            return f"{int(y):04d}{int(mo):02d}{int(d):02d}"
    return None

def extract_patient(text: str):
    for label in ["患者氏名", "患者名", "氏名", "お名前", "患者", "被保険者氏名"]:
        m = re.search(label + r"\s*[:：]?\s*([^\n\r\t 　]{2,30})", text)
        if m:
            return m.group(1).strip()
    m = re.search(r"([^\s\n\r]{2,30})\s*様", text)
    return m.group(1).strip() if m else None

def extract_doctor(text: str):
    for label in ["医師名", "担当医", "先生", "Dr", "Doctor"]:
        m = re.search(label + r"\s*[:：]?\s*([^\n\r\t 　]{2,30})", text, flags=re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return None

def extract_client(text: str):
    m = re.search(r"(営業先|会社名|取引先)\s*[:：]?\s*([^\n\r\t 　]{2,50})", text)
    return m.group(2).strip() if m else None

def extract_client_dept(text: str):
    m = re.search(r"(担当|担当区|部署|部|課)\s*[:：]?\s*([^\n\r\t 　]{2,50})", text)
    return m.group(2).strip() if m else None

def extract_clinic(text: str):
    m = re.search(r"([^\s\n\r]{2,30}(治療院|クリニック|医院))", text)
    return m.group(1).strip() if m else None

def extract_staff(text: str):
    m = re.search(r"(スタッフ|担当者|施術者|作成者)\s*[:：]?\s*([^\n\r\t 　]{2,30})", text)
    return m.group(2).strip() if m else None

def _sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip() or "不明"

def _ym_from_dt(dt: str) -> str:
    return f"{dt[0:4]}年{dt[4:6]}月"

def build_filename(category: str,
                   patient: str,
                   doctor: str,
                   date_str: str,
                   ext: str,
                   text: str) -> str:
    p = _sanitize_filename(patient or "不明")
    d = _sanitize_filename(doctor or "不明")
    dt = date_str or datetime.now().strftime("%Y%m%d")
    ym = _ym_from_dt(dt)

    client = _sanitize_filename(extract_client(text) or "営業先不明")
    client_dept = _sanitize_filename(extract_client_dept(text) or "担当区不明")
    clinic = _sanitize_filename(extract_clinic(text) or "治療院不明")
    staff = _sanitize_filename(extract_staff(text) or "スタッフ不明")

    if category == "同意書":
        return f"同意書_{p}_{d}_{dt}{ext}"
    if category == "保険証":
        return f"保険証_{p}_{dt}{ext}"
    if category == "治療報告書":
        return f"{p}_{client}_{client_dept}_{ym}_{clinic}_治療報告書_{staff}{ext}"
    if category == "患者リスト":
        return f"患者リスト_{p}_{d}_{dt}{ext}"
    if category == "請求書":
        return f"請求書_{clinic}_{ym}{ext}"
    if category == "実績":
        return f"実績_{clinic}_{ym}{ext}"
    return f"{category}_{p}_{dt}{ext}"




