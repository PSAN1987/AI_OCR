# classify_rules.py
# OCRテキストから分類・項目抽出・命名を行うロジック

import re
from datetime import datetime
import unicodedata  # ← 追加

# --- カテゴリごとのキーワードセット ---
KEYWORDS = {
    "同意書": [
        r"同意書", r"同意", r"承諾", r"署名", r"サイン", r"Consent",
    ],
    "保険証": [
        r"健康保険証", r"\b保険証\b", r"保険者番号", r"記号[・\s]*番号",
        r"記号\s*[:：]?\s*\S+", r"番号\s*[:：]?\s*\S+", r"有効期限",
        r"交付日", r"発行者", r"保険者名",
    ],
    "治療報告書": [
        r"治療報告書", r"報告書", r"所見", r"診断", r"経過",
        r"再評価", r"施術計画|治療計画", r"症状|疼痛|ROM|機能評価",
    ],
    # 🔹 患者リストを強化
    "患者リスト": [
        r"患者リスト", r"患者一覧", r"Patient\s*List", r"患者台帳",
        r"フェイスシート", r"利用者情報", r"患者情報", r"ご利用者様",
        r"介護状況", r"入居者情報", r"要介護", r"認定日",
        r"電話番号", r"住所", r"生年月日",
    ],
    "請求書": [
        r"\b請求書\b", r"INVOICE",
        r"請求日", r"請求書番号", r"請求金額|ご請求金額",
        r"振込先|お振込", r"銀行|支店|口座番号",
        r"件名", r"御中",
        r"内訳", r"数量", r"単価", r"消費税", r"合計",
        r"合計金額|ご請求合計",
    ],
    "実績": [
        r"療養費支給申請書", r"あんま|マッサージ", r"施術内訳",
        r"施術日|施術年月日", r"往療", r"通院", r"施術回数|回数",
        r"単価", r"小計|合計|総計|総費用", r"施術者|施術管理者",
        r"申請者", r"審査", r"公費負担", r"受給者番号", r"摘要",
    ],
}

WEIGHTED_KEYWORDS = {
    "同意書": [
        (r"\b同意書\b", 5, True),
        (r"\b同意\b", 2, False),
        (r"署名|サイン|承諾|Consent", 2, False),
    ],
    "保険証": [
        (r"\b(健康)?保険証\b", 5, True),
        (r"保険者番号|記号\s*[:：]?\s*\S+|番号\s*[:：]?\s*\S+", 3, False),
        (r"有効期限|交付日|発行者|保険者名", 2, False),
    ],
    "治療報告書": [
        (r"\b治療報告書\b", 5, True),
        (r"\b報告書\b", 2, False),
        (r"所見|診断|施術計画|治療計画|症状|疼痛|ROM|機能評価", 2, False),
    ],
    "患者リスト": [
        (r"患者(リスト|一覧)|Patient\s*List|患者台帳|フェイスシート", 5, True),
        (r"利用者情報|ご利用者様|入居者情報", 3, False),
        (r"要介護|認定日|電話番号|住所|生年月日", 2, False),
    ],
    "請求書": [
        (r"\b請求書\b|\bINVOICE\b", 5, True),
        (r"請求日|請求書番号|ご?請求金額|振込先|銀行|支店|口座番号|内訳|数量|単価|消費税|合計(金額|)", 2, False),
    ],
    "実績": [
        (r"\b療養費支給申請書\b", 5, True),
        (r"施術内訳|施術日|施術年月日|往療|通院|施術回数|回数|小計|合計|総計|総費用|施術者|施術管理者|申請者|公費負担|受給者番号|摘要", 2, False),
    ],
}

# 誤爆抑制の減点語（任意）
NEGATIVE_HINTS = {
    "請求書": [r"見積書", r"納品書"],
    "同意書": [r"説明書", r"注意書き"],
}


DATE_PATTERNS = [
    r"(20\d{2})[./年-](\d{1,2})[./月-](\d{1,2})日?",
    r"(20\d{2})-(\d{1,2})-(\d{1,2})",
    r"(20\d{2})/(\d{1,2})/(\d{1,2})",
]

def normalize_text(t: str) -> str:
    """OCRノイズを吸収する前処理"""
    if not t:
        return ""
    t = unicodedata.normalize("NFKC", t)          # 全角半角統一
    t = t.replace("ー", "-").replace("―", "-")    # ダッシュ揺れ
    t = t.replace("・", " ")                      # 中黒→空白
    t = re.sub(r"[ \t\u3000]+", " ", t)           # 連続空白を1つに
    return t

# --- 分類 ---
def detect_category(text: str) -> str:
    """
    複数キーワード一致の重み付け＋アンカー語でスコアリング。
    アンカー未ヒット or 僅差は「その他」へ。
    """
    t = normalize_text(text)
    scores = {k: 0 for k in WEIGHTED_KEYWORDS.keys()}
    anchors_hit = {k: 0 for k in WEIGHTED_KEYWORDS.keys()}

    for cat, pats in WEIGHTED_KEYWORDS.items():
        for pat, w, is_anchor in pats:
            hits = re.findall(pat, t, flags=re.IGNORECASE)
            if hits:
                scores[cat] += w * len(hits)
                if is_anchor:
                    anchors_hit[cat] += len(hits)

    # 減点（ネガティブヒント）
    for cat, pats in NEGATIVE_HINTS.items():
        for pat in pats:
            if re.search(pat, t, flags=re.IGNORECASE):
                scores[cat] -= 2

    best = max(scores, key=lambda k: scores[k])
    # アンカー未ヒット or スコア非正 → その他
    if anchors_hit[best] == 0 or scores[best] <= 0:
        return "その他"

    # 僅差（不確実）なら その他
    top2 = sorted(scores.items(), key=lambda kv: kv[1], reverse=True)[:2]
    if len(top2) == 2 and (top2[0][1] - top2[1][1]) < 3:
        return "その他"

    return best


# --- 日付抽出 ---
def extract_date(text: str):
    for p in DATE_PATTERNS:
        m = re.search(p, text)
        if m:
            y, mo, d = m.groups()
            return f"{int(y):04d}{int(mo):02d}{int(d):02d}"
    return None

# --- 各種項目抽出 ---
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

def extract_invoice_clinic(text: str):
    m = re.search(r"([^\s\n\r]{2,50}(治療院|クリニック|医院))[ 　]*(御中)?", text)
    if m:
        return m.group(1).strip()
    return None

# --- ファイル名生成 ---
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
        invoice_clinic = _sanitize_filename(extract_invoice_clinic(text) or clinic)
        return f"請求書_{invoice_clinic}_{ym}{ext}"
    if category == "実績":
        return f"実績_{clinic}_{ym}{ext}"
    return f"{category}_{p}_{dt}{ext}"
