# classify_rules.py
# OCRテキストから分類・項目抽出・命名を行うロジック（既存インターフェース互換）

import re
from datetime import datetime
import unicodedata

# ▼ 既存の import 群の下あたりに追加
LABEL_STOPS = [
    r"生\s*年\s*月\s*日", r"生年月日", r"性別", r"住所", r"有効期限",
    r"保険者番号", r"被保険者番号", r"記号", r"番号",
    r"電話", r"TEL", r"FAX",
    r"保険医氏名", r"医師氏名", r"医師名", r"担当医",
    r"病院名", r"クリニック", r"医院", r"病院", r"施術者", r"担当者",
]
LABEL_STOP_RE = re.compile("|".join(LABEL_STOPS))

def _strip_after_labels(s: str) -> str:
    """「生年月日」「住所」などのラベル語が出たら、それ以降をばっさり捨てる。"""
    s = re.split(LABEL_STOP_RE, s)[0]
    return s.strip(" 　:：.-_/|,、。")

def _clean_name_token(tok: str) -> str:
    """氏名トークン末尾の敬称やラベル片を除去。"""
    tok = re.sub(r"(様|さま|殿|さん)$", "", tok)
    tok = re.sub(r"(生年月日|性別|住所|有効期限|記号|番号)$", "", tok)
    return tok.strip()

# ---------- 正規化 ----------
def normalize_text(t: str) -> str:
    if not t:
        return ""
    t = unicodedata.normalize("NFKC", t)
    t = t.replace("　", " ")
    t = re.sub(r"[ \t]+", " ", t)

    # ★ラベルの分割を結合（ここを追加）
    t = re.sub(r"患\s*者", "患者", t)
    t = re.sub(r"被\s*保\s*険\s*者", "被保険者", t)
    t = re.sub(r"保\s*険\s*医", "保険医", t)
    t = re.sub(r"医\s*師", "医師", t)

    # 住所・氏名ラベルの崩れ補正（既存）
    t = re.sub(r"住\s*所", "住所", t)
    t = re.sub(r"(患者|被保険者|保険医|医師)\s*氏\s*(?:所\s*)?名", r"\1氏名", t)
    t = re.sub(r"氏\s*(?:所\s*)?名", "氏名", t)

    # ★ 医療機関系の割れ表記を結合
    t = re.sub(r"鍼\s*灸\s*院", "鍼灸院", t)
    t = re.sub(r"針\s*灸\s*院", "針灸院", t)
    t = re.sub(r"整\s*骨\s*院", "整骨院", t)
    t = re.sub(r"接\s*骨\s*院", "接骨院", t)
    t = re.sub(r"ク\s*リ\s*ニ\s*ッ\s*ク", "クリニック", t)
    t = re.sub(r"訪\s*問\s*マ\s*ッ\s*サ\s*ー\s*ジ", "訪問マッサージ", t)

    # 相談支援/基本情報 系の割れ補正
    t = re.sub(r"相\s*談\s*支\s*援\s*事\s*業\s*所", "相談支援事業所", t)
    t = re.sub(r"計\s*画\s*作\s*成\s*担\s*当\s*者", "計画作成担当者", t)
    t = re.sub(r"申\s*請\s*者\s*の\s*現\s*状", "申請者の現状", t)
    t = re.sub(r"基\s*本\s*情\s*報", "基本情報", t)
    t = re.sub(r"送\s*付\s*状", "送付状", t)

    # 旧字体の統一（樣 → 様）
    t = t.replace("樣", "様")

    # ラベルの敬称つき表記を補正（患者様氏名→患者氏名 / スタッフ名→スタッフ）
    t = t.replace("患者様氏名", "患者氏名").replace("患者様名", "患者名")
    t = re.sub(r"(スタッフ)\s*名", r"\1", t)
    # 施設名ラベルのゆらぎ
    t = re.sub(r"(治療院|施術所|事業所|クリニック|医院|病院)\s*名", r"\1名", t)

    return t

# ---------- フルネーム判定 ----------
NAME_TOKEN = r"[ぁ-んァ-ンー一-龥々〆ヵヶA-Za-z]{1,15}"
# 区切りに \s（改行含む）/中黒を許容
FULLNAME_SEP    = rf"({NAME_TOKEN})[\s･・]+({NAME_TOKEN})"   # 例: 佐藤 太郎 / 佐藤･太郎 / 佐藤\n太郎
FULLNAME_CONTIG = rf"({NAME_TOKEN})({NAME_TOKEN})"           # 例: 佐藤太郎

# ▼ 既存の _join_fullname を置き換え
def _join_fullname(g1: str, g2: str) -> str:
    g1 = _clean_name_token(g1)
    g2 = _clean_name_token(g2)
    s = f"{g1}{g2}"
    # 末尾にくっついたラベル語は強制削除（後続に数字が続いても落とす）
    s = _strip_after_labels(s)
    # 数字や記号で終わっていたら落とす（住所・年月日の取り込み対策）
    s = re.sub(r"[\d\-/.]+$", "", s)
    return s.strip()

# ── 氏名バリデーション（住所語や項目ラベル、医療機関語を弾く） ──
# ▼ 既存の BAD_ANY_TOKEN を拡張（住所をより弾く）
BAD_ANY_TOKEN = r"(クリニック|病院|医院|医療法人|治療院|薬局|センター|大学|財団|協会|組合|科|御中|貴院|貴社|市|区|町|村|丁目|番地|荘|マンション|アパート|ビル)"

def _is_valid_person_tokens(g1: str, g2: str, role: str) -> bool:
    if re.search(r"\d", g1) or re.search(r"\d", g2):
        return False
    if re.search(BAD_ANY_TOKEN, g1) or re.search(BAD_ANY_TOKEN, g2):
        return False
    if g1 in BAD_ANY_EXACT or g2 in BAD_ANY_EXACT:
        return False
    if role == "patient" and (g1 in BAD_PATIENT_EXACT or g2 in BAD_PATIENT_EXACT):
        return False
    # 一文字×一文字（例：大 昭）は原則棄却（稀例は下流で様付きなどで拾える）
    if len(g1) == 1 and len(g2) == 1:
        return False
    return True

BAD_PATIENT_EXACT = {"生年月日", "住所", "電話番号", "電話", "郵便番号", "患者", "氏名",
                     "保険者番号", "記号", "番号"}
BAD_ANY_EXACT = {"氏名", "患者", "医師", "保険医"}

def _is_valid_person_tokens(g1: str, g2: str, role: str) -> bool:
    # 数字や記号だらけは不可
    if re.search(r"\d", g1) or re.search(r"\d", g2):
        return False
    # クリニック/病院など、明確に人名でない語を含む場合は除外
    if re.search(BAD_ANY_TOKEN, g1) or re.search(BAD_ANY_TOKEN, g2):
        return False
    # ラベルそのものや一般語をそのまま拾っていないか（完全一致で弾く）
    if g1 in BAD_ANY_EXACT or g2 in BAD_ANY_EXACT:
        return False
    if role == "patient" and (g1 in BAD_PATIENT_EXACT or g2 in BAD_PATIENT_EXACT):
        return False
    # 異常に短すぎる/長すぎる（姓名合計 2〜10 文字程度に制限、英字名も想定）
    total = len(g1) + len(g2)
    if total < 2 or total > 10:
        # 長めのカタカナ名などを許容したい場合は上限を緩めてもOK
        pass
    return True

# 住所っぽい候補の排除（氏名誤認防止：港区新茶屋 等）
ADDRESS_TOKENS = r"(都|道|府|県|市|区|町|村|丁目|番地|番|号|郡|荘|マンション|アパート|ビル)"
def _looks_addressy(s: str) -> bool:
    return bool(re.search(ADDRESS_TOKENS, s))

# ---------- カテゴリキーワード（単純スコア） ----------
# --- 既存 KEYWORDS を以下のように一部差し替え ---

KEYWORDS = {
    "同意書": [r"同意書", r"同意", r"承諾", r"署名", r"サイン", r"Consent"],
    "保険証": [
        r"健康保険証", r"\b保険証\b", r"保険者番号", r"記号[・\s]*番号",
        r"記号\s*[:：]?\s*\S+", r"番号\s*[:：]?\s*\S+", r"有効期限",
        r"交付日", r"発行者", r"保険者名",
    ],
    "治療報告書": [
        r"治療報告書", r"施術報告書", r"経過報告", r"報告対象年月",
        r"所感", r"目標", r"現状維持|現在の状態|当初からの状態の変化",
        r"初療日", r"往診日|訪問日|振替訪問日|休診日",
        r"マッサージ|施術|重点部位",
        r"疼痛|ROM|機能評価|再評価|治療計画|施術計画",
    ],
    "患者リスト": [
        r"患者リスト", r"患者一覧", r"Patient\s*List", r"患者台帳",
        r"フェイスシート", r"利用者情報", r"ご利用者様",
        r"介護状況", r"入居者情報", r"要介護", r"認定日",
        r"電話番号", r"住所", r"生年月日",
        r"リハビリテーション総合実施計画書", r"受入依頼票",
    ],
    "請求書": [
        r"\b請求書\b", r"\bINVOICE\b",
        r"請求日", r"請求書番号", r"請求金額|ご請求金額",
        r"振込先|お振込", r"銀行|支店|口座番号",
        r"件名",
        r"内訳", r"数量", r"単価", r"消費税", r"合計",
        r"合計金額|ご請求合計",
        # ★「御中」は請求書以外にも頻出するため除外（誤爆回避）
    ],
    "実績": [
        r"療養費支給申請書", r"あんま|マッサージ", r"施術内訳",
        r"施術日|施術年月日", r"往療", r"通院", r"施術回数|回数",
        r"単価", r"小計|合計|総計|総費用", r"施術者|施術管理者",
        r"審査", r"公費負担", r"受給者番号", r"摘要",
    ],
}

# ---------- 日付 ----------
DATE_PATTERNS = [
    r"(20\d{2})[./年-](\d{1,2})[./月-](\d{1,2})日?",
    r"(20\d{2})-(\d{1,2})-(\d{1,2})",
    r"(20\d{2})/(\d{1,2})/(\d{1,2})",
    r"令和\s*(\d{1,2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?",
]

def extract_date(text: str):
    """YYYYMMDD（文字列）を返す。見つからなければ None。令和対応。"""
    t = normalize_text(text)
    for p in DATE_PATTERNS:
        m = re.search(p, t)
        if m:
            if "令和" in p:
                y = 2018 + int(m.group(1))  # 令和1=2019
                mo, d = m.group(2), m.group(3)
            else:
                y, mo, d = m.groups()
            return f"{int(y):04d}{int(mo):02d}{int(d):02d}"
    return None

def detect_category(text: str) -> str:
    t = normalize_text(text)

    # 1) 患者リストの強シグナル（既存）
    STRONG_PATIENTLIST = [
        r"患者(一覧|台帳)", r"Patient\s*List", r"フェイスシート",
        r"利用者基本情報", r"基本情報", r"ご利用者様", r"申請者の現状",
        r"リハビリテーション総合実施計画書", r"受入依頼票",
    ]
    for pat in STRONG_PATIENTLIST:
        if re.search(pat, t, flags=re.IGNORECASE):
            return "患者リスト"

    # 2) ★治療報告書の強シグナル（あれば即決）
    STRONG_REPORT = [
        r"治療報告書", r"施術報告書", r"経過報告",
        r"(?:報告対象年月).*(?:目標|所感|現状)",  # 同ページ内に共起
    ]
    for pat in STRONG_REPORT:
        if re.search(pat, t, flags=re.IGNORECASE | re.DOTALL):
            return "治療報告書"

    # 3) スコアリング（既存）
    scores = {k: 0 for k in KEYWORDS.keys()}
    for cat, pats in KEYWORDS.items():
        for pat in pats:
            hits = re.findall(pat, t, flags=re.IGNORECASE)
            if hits:
                scores[cat] += len(hits)

    best = max(scores, key=lambda k: scores[k])

    # 4) 実績 vs 患者リストの既存タイブレーク
    if best == "実績":
        if (re.search(r"(相談支援|相談支援事業所|計画作成担当者|基本情報)", t)
            and not re.search(r"(療養費|療養費支給申請書)", t)):
            return "患者リスト"

    # 5) ★請求書の誤爆抑制：「請求固有語」がなければ請求書にしない
    if best == "請求書":
        invoice_core = re.search(
            r"(請求書|INVOICE|請求書番号|請求金額|ご請求金額|振込先|内訳|合計金額)",
            t, flags=re.IGNORECASE
        )
        if not invoice_core:
            # 報告書っぽい語が多ければ治療報告書へ倒す
            if re.search(r"(報告対象年月|所感|目標|初療日|往診日|施術|マッサージ)", t):
                return "治療報告書"

    # 6) ★報告書を優先する追加タイブレーク
    if best in {"実績", "患者リスト"}:
        if re.search(r"(報告対象年月|所感|目標|初療日|往診日)", t):
            return "治療報告書"

    return best if scores[best] > 0 else "その他"


# ---------- 項目抽出 ----------
# ▼ 既存 _fullname_on_same_line_after を置き換え（良い候補を後方優先で選別）
def _fullname_on_same_line_after(label: str, t: str):
    m = re.search(label + r"\s*[:：]?[^\n]*", t)
    if not m:
        return None
    line = _strip_after_labels(m.group(0))  # ラベル語以降は切り落とす
    cands = list(re.finditer(FULLNAME_SEP, line)) or list(re.finditer(FULLNAME_CONTIG, line))
    for g in reversed(cands):
        g1, g2 = g.group(1), g.group(2)
        if g1 == g2:
            # 「上野 上野 みどり」のような重複に強い：直後に3語目があれば差し替え
            tail = line[g.end():g.end()+20]
            nxt = re.match(r"\s*([ぁ-んァ-ンー一-龥々〆ヵヶA-Za-z]{1,15})", tail)
            if nxt and nxt.group(1) != g2:
                g2 = nxt.group(1)
        cand = _join_fullname(g1, g2)
        if cand and not _looks_addressy(cand) and _is_valid_person_tokens(g1, g2, "patient"):
            return cand
    return None

# ▼ 既存 _fullname_on_next_line_after も軽く強化
def _fullname_on_next_line_after(label: str, t: str):
    m = re.search(label + r"\s*[:：]?.*?\n", t)
    if not m:
        return None
    start = m.end()
    next_line = _strip_after_labels(t[start:start+120])
    g = re.search(FULLNAME_SEP, next_line) or re.search(FULLNAME_CONTIG, next_line)
    if g:
        g1, g2 = g.group(1), g.group(2)
        cand = _join_fullname(g1, g2)
        if cand and not _looks_addressy(cand) and _is_valid_person_tokens(g1, g2, "patient") and g1 != g2:
            return cand
    return None

def _name_after_label_window(lb: str, t: str):
    m = re.search(lb + r"\s*[:：]?\s*([^\n]{0,80})", t)
    if not m:
        return None
    win = _strip_after_labels(m.group(1))
    # 行内でフルネーム探索（区切り/連結の両対応）
    g = re.search(FULLNAME_SEP, win) or re.search(FULLNAME_CONTIG, win)
    if not g:
        return None
    g1, g2 = g.group(1), g.group(2)
    cand = _join_fullname(g1, g2)
    if cand and not _looks_addressy(cand) and _is_valid_person_tokens(g1, g2, "patient") and g1 != g2:
        return cand
    return None


def _fullname_after_broken_shimei(t: str):
    """
    同一行内で「氏 … 名」のように割れているケースを救済し、
    その行の末尾側にあるフルネームを返す。
    """
    for m in re.finditer(r"氏[^\n]{0,40}名[^\n]*", t):
        line = m.group(0)
        cands = list(re.finditer(FULLNAME_SEP, line)) or list(re.finditer(FULLNAME_CONTIG, line))
        if cands:
            g = cands[-1]
            return _join_fullname(g.group(1), g.group(2))
    return None

# ▼ 既存 extract_patient を置き換え
def extract_patient(text: str):
    t = normalize_text(text)
    TAIL_SAMA = r"(?:\s*(?:様|樣)(?:の|は|です|で|に)?|\s*さま)"

    def _accept(g1: str, g2: str):
        cand = _join_fullname(g1, g2)
        return cand if cand and not _looks_addressy(cand) and _is_valid_person_tokens(g1, g2, "patient") else None

    LABELS = [r"患者氏名", r"患者名", r"患者様氏名", r"患者様名", r"被保険者氏名", r"被保険者名"]

    # 0) 様付き（保険証など）最優先
    m = re.search(FULLNAME_SEP + TAIL_SAMA, t) or re.search(FULLNAME_CONTIG + TAIL_SAMA, t)
    if m:
        cand = _join_fullname(m.group(1), m.group(2))
        if cand and not _looks_addressy(cand):
            return cand

    # 1) ラベル直後の「窓取り」→ 最後に既存逐次探索
    for lb in LABELS:
        cand = _name_after_label_window(lb, t)
        if cand:
            return cand
        m = re.search(lb + r"\s*[:：]?\s*" + FULLNAME_SEP, t)
        if m:
            cand = _accept(m.group(1), m.group(2))
            if cand: return cand
        m2 = re.search(lb + r"\s*[:：]?\s*" + FULLNAME_CONTIG, t)
        if m2:
            cand = _accept(m2.group(1), m2.group(2))
            if cand: return cand
        fn = _fullname_on_same_line_after(lb, t)
        if fn: return fn
        fn2 = _fullname_on_next_line_after(lb, t)
        if fn2: return fn2

    # 2) ラベルなしでも「患者 …」近傍で拾う（窓取り）
    m = re.search(r"(患者[^\n]{0,60})", t)
    if m:
        cand = _name_after_label_window("患者", m.group(1))
        if cand:
            return cand

    # 3) 「氏…名」割れ救済
    fn_broken = _fullname_after_broken_shimei(t)
    if fn_broken and not _looks_addressy(fn_broken):
        return fn_broken

    return None

def extract_staff(text: str):
    t = normalize_text(text)
    # ★ スタッフ名/スタッフ氏名も拾う
    m = re.search(r"(スタッフ(?:氏名|名)?|担当者|施術者|作成者)\s*[:：]?\s*([^\n\r\t 　]{2,30})", t)
    return m.group(2).strip() if m else None

def extract_doctor(text: str):
    """医師/保険医のフルネームを抽出。改行区切りにも対応。"""
    t = normalize_text(text)
    for lb in [r"保険医氏名", r"医師氏名", r"医師名", r"担当医", r"先生", r"Dr", r"Doctor"]:
        m_sep = re.search(lb + r"\s*[:：]?\s*" + FULLNAME_SEP, t, flags=re.IGNORECASE)
        if m_sep: return _join_fullname(m_sep.group(1), m_sep.group(2))
        m_contig = re.search(lb + r"\s*[:：]?\s*" + FULLNAME_CONTIG, t, flags=re.IGNORECASE)
        if m_contig: return _join_fullname(m_contig.group(1), m_contig.group(2))
    return None  # 片方だけは未採用

def extract_client(text: str):
    m = re.search(r"(営業先|会社名|取引先)\s*[:：]?\s*([^\n\r\t 　]{2,50})", text)
    return m.group(2).strip() if m else None

def extract_client_dept(text: str):
    m = re.search(r"(担当|担当区|部署|部|課)\s*[:：]?\s*([^\n\r\t 　]{2,50})", text)
    return m.group(2).strip() if m else None

# 施設名の語尾（増強）
CLINIC_SUFFIX = r"(?:訪問マッサージ鍼灸院|鍼灸院|針灸院|はりきゅう院|鍼灸整骨院|整骨院|接骨院|整体院|治療院|クリニック|医院|病院|医科|歯科|施術所)"

def extract_clinic(text: str):
    t = normalize_text(text)

    # ★ ラベル明示のときは最優先で取得
    m = re.search(r"(治療院名|施術所名|事業所名|クリニック名|医院名|病院名)\s*[:：]?\s*([^\n\r]{2,60})", t)
    if m:
        return m.group(2).strip()

    # （以下は既存 A/B/C のロジックを継続）
    m = re.search(rf"([^\n\r]{{2,60}}?{CLINIC_SUFFIX})\s*(?:御中|様|殿|宛)", t)
    if m:
        return m.group(1).strip()
    m = re.search(rf"([^\s\n\r]{{1,60}}{CLINIC_SUFFIX})", t)
    if m:
        return m.group(1).strip()
    m = re.search(r"([^\n\r]{2,60}?)\s*(?:御中|様|殿|宛)", t)
    if m and re.search(CLINIC_SUFFIX, m.group(1)):
        return m.group(1).strip()
    return None

def extract_invoice_clinic(text: str):
    """請求書の宛先施設名を優先的に抽出。"""
    t = normalize_text(text)

    # 1) 宛先ラベル／敬称優先
    m = re.search(rf"([^\n\r]{{2,60}}?{CLINIC_SUFFIX})\s*(?:御中|様|殿|宛)", t)
    if m:
        return m.group(1).strip()

    # 2) 「請求先/宛先」っぽい行（あれば）
    m = re.search(rf"(?:請求先|宛先)\s*[:：]?\s*([^\n\r]{{2,60}}?{CLINIC_SUFFIX})", t)
    if m:
        return m.group(1).strip()

    # 3) 一般の施設名
    m = re.search(rf"([^\s\n\r]{{1,60}}{CLINIC_SUFFIX})", t)
    if m:
        return m.group(1).strip()

    return None


def extract_staff(text: str):
    m = re.search(r"(スタッフ|担当者|施術者|作成者)\s*[:：]?\s*([^\n\r\t 　]{2,30})", text)
    return m.group(2).strip() if m else None


# ---------- ファイル名生成 ----------
def _sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip() or "不明"

def _ym_from_dt(dt: str) -> str:
    return f"{dt[0:4]}年{dt[4:6]}月"

def _compact(s: str) -> str:
    """余分な空白/アンダースコアを整理。"""
    s = re.sub(r"[ \u3000]+", " ", s).strip()
    s = re.sub(r"_+", "_", s)
    s = re.sub(r"[ _]{2,}", " ", s)
    return s

def _tokens(text: str, patient: str, doctor: str, date_str: str) -> dict:
    dt = date_str or datetime.now().strftime("%Y%m%d")
    ym = _ym_from_dt(dt)
    return {
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

# カテゴリ別テンプレート（差し替え容易）
NAMING_TEMPLATES = {
    "同意書":      "同意書_{patient}_{doctor}_{date}",
    "保険証":      "保険証_{patient}_{date}",
    "治療報告書":  "{patient}_{client}_{client_dept}_{ym}_{clinic}_治療報告書_{staff}",
    "患者リスト":  "患者リスト_{patient}_{doctor}_{date}",
    "請求書":      "請求書_{invoice_clinic}_{ym}",
    "実績":        "実績_{clinic}_{ym}",
}

def build_filename(category: str,
                   patient: str,
                   doctor: str,
                   date_str: str,
                   ext: str,
                   text: str) -> str:
    """テンプレート駆動の命名（既存の引数/戻り値は不変）。"""
    toks = _tokens(text, patient, doctor, date_str)
    tmpl = NAMING_TEMPLATES.get(category, "{cat}_{patient}_{date}")
    name = tmpl.format_map({**toks, "cat": category})
    name = _compact(name)
    # OneDrive で扱いやすい長さに丸め（拡張子は維持）
    MAX_BASENAME = 80
    if len(name) > MAX_BASENAME:
        name = name[:MAX_BASENAME].rstrip("_ ").rstrip()
    return f"{name}{ext}"

