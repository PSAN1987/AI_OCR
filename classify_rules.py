# classify_rules.py
# OCRテキストから分類・項目抽出・命名を行うロジック

import re
from datetime import datetime
import unicodedata

def normalize_text(t: str) -> str:
    if not t:
        return ""
    t = unicodedata.normalize("NFKC", t)
    t = t.replace("　", " ")
    t = re.sub(r"[ \t]+", " ", t)
    # ラベルの分割（例: 患者 氏 名 / 保険医 氏 名）を結合
    t = re.sub(r"(患者|被保険者|保険医|医師)\s*氏\s*名", r"\1氏名", t)
    t = re.sub(r"氏\s*名", "氏名", t)
    return t

# 2. フルネーム判定用のパターン
NAME_TOKEN = r"[一-龥々〆ヵヶァ-ンーA-Za-z]{1,15}"
# 区切りありのフルネーム（空白/中黒）
FULLNAME_SEP = rf"({NAME_TOKEN})[ ･・]+({NAME_TOKEN})"
# 区切りなしのフルネーム（例: 山田太郎／カタカナ連結）
# → 2トークン連結とみなす（最短一致で頭2トークンを拾う）
FULLNAME_CONTIG = rf"({NAME_TOKEN})({NAME_TOKEN})"

def _join_fullname(g1: str, g2: str) -> str:
    # ファイル名用にスペース・中黒は除去（「可知美恵子」形式）
    return f"{g1}{g2}"


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

DATE_PATTERNS = [
    r"(20\d{2})[./年-](\d{1,2})[./月-](\d{1,2})日?",
    r"(20\d{2})-(\d{1,2})-(\d{1,2})",
    r"(20\d{2})/(\d{1,2})/(\d{1,2})",
]

# --- 分類 ---
def detect_category(text: str) -> str:
    """
    複数キーワード一致数でスコアリングし、最もスコアが高いカテゴリを返す。
    """
    t = text.replace("　", " ")
    scores = {k: 0 for k in KEYWORDS.keys()}
    for cat, pats in KEYWORDS.items():
        for pat in pats:
            matches = re.findall(pat, t, flags=re.IGNORECASE)
            if matches:
                # ヒット回数を加点
                scores[cat] += len(matches)
    best = max(scores, key=lambda k: scores[k])
    return best if scores[best] > 0 else "その他"

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
    t = normalize_text(text)

    # まずは明示ラベルを優先
    label_patterns = [r"患者氏名", r"患者名", r"被保険者氏名", r"被保険者名"]
    for lb in label_patterns:
        # 例: 患者氏名 佐藤 太郎 ／ 患者名 佐藤太郎
        m = re.search(lb + r"\s*[:：]?\s*" + FULLNAME_SEP, t)
        if m:
            return _join_fullname(m.group(1), m.group(2))
        m2 = re.search(lb + r"\s*[:：]?\s*" + FULLNAME_CONTIG, t)
        if m2:
            return _join_fullname(m2.group(1), m2.group(2))

    # 汎用「氏名」だが、保険医/医師/施術者/担当者/保険者/被保険者 由来は除外
    excl = r"(?<!保険医)(?<!医師)(?<!施術者)(?<!担当者)(?<!保険者)(?<!被保険者)"
    m = re.search(excl + r"氏名\s*[:：]?\s*" + FULLNAME_SEP, t)
    if m:
        return _join_fullname(m.group(1), m.group(2))
    m2 = re.search(excl + r"氏名\s*[:：]?\s*" + FULLNAME_CONTIG, t)
    if m2:
        return _join_fullname(m2.group(1), m2.group(2))

    # 「フルネーム様」パターン（例: 佐藤 太郎 様 / 佐藤太郎 様）
    m = re.search(FULLNAME_SEP + r"\s*様\b", t)
    if m:
        return _join_fullname(m.group(1), m.group(2))
    m2 = re.search(FULLNAME_CONTIG + r"\s*様\b", t)
    if m2:
        return _join_fullname(m2.group(1), m2.group(2))

    # ここまで見つからなければ “不明” 扱い（フルネーム未満は返さない）
    return None

def extract_doctor(text: str):
    t = normalize_text(text)
    # 医師系ラベルを網羅（保険医氏名を最優先）
    label_patterns = [r"保険医氏名", r"医師氏名", r"医師名", r"担当医", r"先生", r"Dr", r"Doctor"]

    for lb in label_patterns:
        m = re.search(lb + r"\s*[:：]?\s*" + FULLNAME_SEP, t, flags=re.IGNORECASE)
        if m:
            return _join_fullname(m.group(1), m.group(2))
        m2 = re.search(lb + r"\s*[:：]?\s*" + FULLNAME_CONTIG, t, flags=re.IGNORECASE)
        if m2:
            return _join_fullname(m2.group(1), m2.group(2))

    # 医師名に「様」は通常付かないので “様” フォールバックは無し
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
# 追加: 先頭付近のimportsのままでOK（re, datetimeは既にあります）

def _sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip() or "不明"

def _ym_from_dt(dt: str) -> str:
    return f"{dt[0:4]}年{dt[4:6]}月"

def _compact(s: str) -> str:
    # 連続した空白/アンダースコアを縮め、全角空白も吸収
    s = re.sub(r"[ \u3000]+", " ", s).strip()
    s = re.sub(r"_+", "_", s)
    s = re.sub(r"[ _]{2,}", " ", s)
    return s

def _tokens(text: str, patient: str, doctor: str, date_str: str) -> dict:
    dt = date_str or datetime.now().strftime("%Y%m%d")
    ym = _ym_from_dt(dt)
    tokens = {
        "patient": _sanitize_filename(patient or "不明"),
        "doctor": _sanitize_filename(doctor or "不明"),
        "date": dt,
        "ym": ym,
        # 既存抽出器を再利用
        "client": _sanitize_filename(extract_client(text) or "営業先不明"),
        "client_dept": _sanitize_filename(extract_client_dept(text) or "担当区不明"),
        "clinic": _sanitize_filename(extract_clinic(text) or "治療院不明"),
        "staff": _sanitize_filename(extract_staff(text) or "スタッフ不明"),
        "invoice_clinic": _sanitize_filename(extract_invoice_clinic(text) or (extract_clinic(text) or "治療院不明")),
    }
    return tokens

# カテゴリ別テンプレート（あとから差し替え可能）
NAMING_TEMPLATES = {
    "同意書":      "同意書_{patient}_{doctor}_{date}",
    "保険証":      "保険証_{patient}_{date}",
    "治療報告書":  "{patient}_{client}_{client_dept}_{ym}_{clinic}_治療報告書_{staff}",
    "患者リスト":  "患者リスト_{patient}_{doctor}_{date}",
    "請求書":      "請求書_{invoice_clinic}_{ym}",
    "実績":        "実績_{clinic}_{ym}",
    # 将来カテゴリを足す場合はここに追記
}

def build_filename(category: str,
                   patient: str,
                   doctor: str,
                   date_str: str,
                   ext: str,
                   text: str) -> str:
    toks = _tokens(text, patient, doctor, date_str)
    # テンプレートがなければ従来に近い汎用形式へ
    tmpl = NAMING_TEMPLATES.get(category, "{cat}_{patient}_{date}")
    name = tmpl.format_map({**toks, "cat": category})
    name = _compact(name)
    # 長すぎるとOneDriveの扱いが悪くなるので丸め（拡張子は維持）
    MAX_BASENAME = 80
    if len(name) > MAX_BASENAME:
        name = name[:MAX_BASENAME].rstrip("_ ").rstrip()
    return f"{name}{ext}"
