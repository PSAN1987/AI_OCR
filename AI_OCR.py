# AI_OCR.py
# -*- coding: utf-8 -*-
"""
LINE Bot with AI OCR (Google Vision) + OneDrive integration

A) 受信自動化
   - 画像（ImageMessage）/ PDF（FileMessage）を受信
   - Google VisionでOCR
     * 画像: images:annotate（APIキー or サービスアカウントのどちらでも可）
     * PDF : files:asyncBatchAnnotate（サービスアカウント + GCS必須）
   - OCRテキストからカテゴリ判定・氏名/先生名/日付を抽出
   - 命名規則に従って OneDrive に分類保存
   - 共有リンクを作成して LINE で保存結果と保存先を返信

B) 送信自動化（試験）
   - ユーザーがテキストで「名前」を送信
   - OneDrive 全階層を検索（ファイル名ヒット）→ 一致ファイルのリンクを返信
"""

# ★ classify_rules のみを使用（Main Code 側の重複実装は削除）
from classify_rules import (
    detect_category,
    extract_patient,
    extract_doctor,
    extract_date,
    build_filename,
)

import os
import re
import json
import time
import uuid
import base64
import hashlib
import unicodedata
from datetime import datetime
from urllib.parse import quote

import requests
from requests.exceptions import HTTPError
from flask import Flask, request, abort

# LINE SDK
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError, LineBotApiError
from linebot.models import (
    MessageEvent,
    ImageMessage,
    FileMessage,
    TextMessage,
    TextSendMessage,
)

# Google Auth / GCS
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import Request as GoogleAuthRequest
from google.cloud import storage

# ---------- Config ----------
CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET", "")

SERVICE_ACCOUNT_VALUE = os.environ.get("GCP_SERVICE_ACCOUNT_JSON", "").strip()
VISION_API_KEY = os.environ.get("GOOGLE_VISION_API_KEY", "").strip()
GCS_BUCKET = os.environ.get("GCS_BUCKET", "").strip()
SPREADSHEET_KEY = os.environ.get("SPREADSHEET_KEY", "").strip()
SPREADSHEET_NAME = "Logs"

MS_TENANT_ID = os.environ.get("MS_TENANT_ID", "").strip()
MS_CLIENT_ID = os.environ.get("MS_CLIENT_ID", "").strip()
MS_CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET", "").strip()
ONEDRIVE_DRIVE_ID = os.environ.get("ONEDRIVE_DRIVE_ID", "").strip()
ONEDRIVE_USER_ID = os.environ.get("ONEDRIVE_USER_ID", "").strip()
ONEDRIVE_BASE_FOLDER = os.environ.get("ONEDRIVE_BASE_FOLDER", "/").strip() or "/"
ONEDRIVE_LINK_SCOPE = os.environ.get("ONEDRIVE_LINK_SCOPE", "organization").strip()

VISION_PDF_POLL_TIMEOUT_SEC = int(os.environ.get("VISION_PDF_POLL_TIMEOUT_SEC", "90"))
VISION_PDF_POLL_INTERVAL_SEC = int(os.environ.get("VISION_PDF_POLL_INTERVAL_SEC", "3"))

if not (CHANNEL_ACCESS_TOKEN and CHANNEL_SECRET and SERVICE_ACCOUNT_VALUE and MS_TENANT_ID and MS_CLIENT_ID and MS_CLIENT_SECRET):
    missing = [k for k, v in {
        "LINE_CHANNEL_ACCESS_TOKEN": CHANNEL_ACCESS_TOKEN,
        "LINE_CHANNEL_SECRET": CHANNEL_SECRET,
        "GCP_SERVICE_ACCOUNT_JSON": SERVICE_ACCOUNT_VALUE,
        "MS_TENANT_ID": MS_TENANT_ID,
        "MS_CLIENT_ID": MS_CLIENT_ID,
        "MS_CLIENT_SECRET": MS_CLIENT_SECRET,
    }.items() if not v]
    raise RuntimeError(f"Missing environment variables: {', '.join(missing)}")

# ---------- Flask & LINE setup ----------
app = Flask(__name__)
line_bot_api = LineBotApi(CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)

# ---------- Utils ----------
def _load_service_account_info(value: str) -> dict:
    """JSON文字列 or JSONファイルパスの両対応でdictを返す"""
    try:
        return json.loads(value)
    except json.JSONDecodeError:
        path = os.path.expanduser(value)
        if not os.path.isfile(path):
            raise RuntimeError("GCP_SERVICE_ACCOUNT_JSON must be a JSON string or an existing file path.")
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

def _google_access_token(scopes=("https://www.googleapis.com/auth/cloud-platform",)):
    info = _load_service_account_info(SERVICE_ACCOUNT_VALUE)
    creds = Credentials.from_service_account_info(info, scopes=list(scopes))
    creds.refresh(GoogleAuthRequest())
    return creds.token

def _gcs_client():
    info = _load_service_account_info(SERVICE_ACCOUNT_VALUE)
    return storage.Client.from_service_account_info(info)

def _now_date_str():
    # 保存日付：YYYYMMDD
    return datetime.now().strftime("%Y%m%d")

def _sanitize_filename(name: str) -> str:
    # OneDrive禁止文字: \ / : * ? " < > | など
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    name = name.replace("\n", " ").replace("\r", " ")
    return name.strip() or "unnamed"

def gsheet_append_rows(rows: list[list[str]]):
    """
    rows: [["保存日時ISO", "保存日付YYYYMMDD", "種別", "分類", "患者", "先生", "抽出日付", "保存フォルダ", "ファイル名", "リンク", "OCR文字数", "OCR先頭100", "ステータス", "イベントID", "エラーメッセージ"]]
    """
    if not SPREADSHEET_KEY:
        # 設定が無ければ黙ってスキップ（本番運用ではログにWarnしてOK）
        return
    token = _google_access_token(scopes=("https://www.googleapis.com/auth/spreadsheets",))
    rng = f"{quote(SPREADSHEET_NAME)}!A1"
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_KEY}/values/{rng}:append"
    params = {"valueInputOption": "RAW", "insertDataOption": "INSERT_ROWS"}
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {"values": rows}
    r = requests.post(url, headers=headers, params=params, json=body, timeout=30)
    # 失敗してもメイン処理は継続させたいので raise はしない（必要ならここで例外化）
    if r.status_code not in (200, 201):
        try:
            detail = r.text[:300]
        except Exception:
            detail = ""
        print(f"[WARN] Sheets append failed: {r.status_code} {detail}")

# ---------- 治療報告書：OCRパーサ & 高精度命名 ----------
def _norm(s: str) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", s)
    # 空白（半角/全角/タブ）除去
    s = re.sub(r"[ \u3000\t]+", "", s)
    # 禁止文字を置換
    s = re.sub(r'[\\/:*?"<>|]', "_", s)
    # 連続 "_" を縮約
    s = re.sub(r"_+", "_", s).strip("_")
    return s

def _hash6(*parts) -> str:
    h = hashlib.sha1(("|".join(parts)).encode("utf-8")).hexdigest()
    return h[:6]

def build_filename_treatment_report(meta: dict, ext: str = ".pdf") -> str:
    """
    固定順命名:
      {患者名}_{YYYY-MM}_治療報告書_{治療院名}_{スタッフ名}_{市区}_{事業所名}_{作成日時YYYYMMDD-HHMMSS}{ext}
    """
    patient = _norm(meta.get("patient", ""))
    clinic  = _norm(meta.get("clinic", ""))
    staff   = _norm(meta.get("staff", ""))
    city    = _norm(meta.get("city", ""))
    office  = _norm(meta.get("office", ""))

    # YYYY-MM
    yyyymm = None
    for src in [meta.get("period_start"), meta.get("created_at")]:
        if src:
            m = re.search(r"(\d{4})[-/](\d{2})", src)
            if m:
                yyyymm = f"{m.group(1)}-{m.group(2)}"; break
    if not yyyymm:
        yyyymm = datetime.now().strftime("%Y-%m")

    # 作成日時スタンプ
    ts = meta.get("created_at") or ""
    m = re.search(r"(\d{4})[-/](\d{2})[-/](\d{2}).*?(\d{2}):(\d{2}):(\d{2})", ts)
    stamp = f"{m.group(1)}{m.group(2)}{m.group(3)}-{m.group(4)}{m.group(5)}{m.group(6)}" if m else ""

    parts = [patient, yyyymm, "治療報告書", clinic, staff, city, office, stamp]
    parts = [p for p in parts if p]
    name = "_".join(parts) + (ext if ext.startswith(".") else f".{ext}")

    # 長い場合、office を短縮＋ハッシュで衝突回避
    if len(name) > 120:
        short_office = (office[:12] + "…" if len(office) > 13 else office)
        hx = _hash6(patient, yyyymm, clinic, staff, city, office, stamp)
        parts2 = [patient, yyyymm, "治療報告書", clinic, staff, city, short_office + "_" + hx, stamp]
        parts2 = [p for p in parts2 if p]
        name = "_".join(parts2) + (ext if ext.startswith(".") else f".{ext}")
    return name

def parse_meta_from_tiryo_houkokusho(text: str) -> dict:
    """
    OCR結果から治療報告書の主要メタを抽出（全半角/改行/ラベル崩れにロバスト）
    """
    t = unicodedata.normalize("NFKC", text or "")
    t_space = re.sub(r"\s+", " ", t).strip()

    # 1) 期間
    period_start = period_end = ""
    m = re.search(r"(\d{4}[-/]\d{2}[-/]\d{2})\s*[-〜~]\s*(\d{4}[-/]\d{2}[-/]\d{2})", t_space)
    if m:
        period_start, period_end = m.group(1), m.group(2)

    # 2) 作成日時
    created_at = ""
    m = re.search(r"作成日時\s*[:：]?\s*(\d{4}[-/]\d{2}[-/]\d{2})\s+(\d{2}:\d{2}:\d{2})", t_space)
    if m:
        created_at = f"{m.group(1)} {m.group(2)}"

    # 3) 治療院名・スタッフ名・患者名（期間の直後～最初の「様」まで）
    clinic = staff = patient = ""
    tail = t_space
    if period_end:
        idx = t_space.find(period_end)
        if idx >= 0:
            tail = t_space[idx + len(period_end):].strip()
    idx_sama = tail.find("様")
    head = tail[:idx_sama] if idx_sama != -1 else tail
    toks = [x for x in head.split(" ") if x]
    if len(toks) >= 2:
        patient = toks[-1]
        staff   = toks[-2]
        clinic  = " ".join(toks[:-2]).strip()
        # 混入するラベル文言を除去
        clinic  = re.sub(r"(治療報告書|報告対象年月|治療院名|スタッフ名|患者様氏名|所|患者様住)+", "", clinic).strip()

    # 4) 市区・事業所名（「事業所名 … 御中」区間）
    city = ""
    office = ""
    idx_onchu = t_space.find("御中")
    idx_offlbl = t_space.find("事業所名")
    if idx_offlbl != -1 and idx_onchu != -1 and idx_offlbl < idx_onchu:
        seg = t_space[idx_offlbl + len("事業所名"): idx_onchu]
        seg = re.sub(r"(町村|配布先担当者|様)+", " ", seg).strip()
        seg_toks = [x for x in seg.split(" ") if x]
        for i, tok in enumerate(seg_toks):
            if tok.endswith(("市", "区", "町", "村")):
                city = tok
                office = " ".join(seg_toks[i+1:]).strip()
                break
        if not office and seg_toks:
            office = " ".join(seg_toks).strip()
    else:
        # フォールバック
        pre = t_space[:idx_onchu] if idx_onchu != -1 else t_space
        for tok in pre.split(" "):
            if tok.endswith(("市", "区", "町", "村")):
                city = tok
        m_off = re.search(r"([^\s]{2,50})\s*御中", t_space)
        if m_off:
            office = m_off.group(1)

    return {
        "patient": patient, "clinic": clinic, "staff": staff,
        "city": city, "office": office,
        "period_start": period_start, "period_end": period_end,
        "created_at": created_at,
    }

# ---------- OCR (Images via Vision) ----------
def ocr_image_bytes(image_bytes: bytes) -> str:
    """画像のOCR。APIキーがあればそれを、なければSA OAuthで呼び出し"""
    url = "https://vision.googleapis.com/v1/images:annotate"
    payload = {
        "requests": [{
            "image": {"content": base64.b64encode(image_bytes).decode("utf-8")},
            "features": [{"type": "DOCUMENT_TEXT_DETECTION"}],
            "imageContext": {"languageHints": ["ja", "en"]}
        }]
    }
    headers = {"Content-Type": "application/json"}
    params = {}
    if VISION_API_KEY:
        params["key"] = VISION_API_KEY
    else:
        token = _google_access_token()
        headers["Authorization"] = f"Bearer {token}"

    resp = requests.post(url, params=params, headers=headers, json=payload, timeout=60)
    try:
        resp.raise_for_status()
    except HTTPError as he:
        body = resp.text[:300] + "..." if resp is not None and resp.text else ""
        raise RuntimeError(f"Vision images:annotate error: HTTP {resp.status_code} {resp.reason} {body}") from he

    data = resp.json()
    try:
        return data["responses"][0]["fullTextAnnotation"]["text"]
    except Exception:
        ann = data["responses"][0].get("textAnnotations", [])
        return ann[0]["description"] if ann else ""

# ---------- OCR (PDF via Vision Async + GCS) ----------
def ocr_pdf_bytes_via_gcs(pdf_bytes: bytes, filename_hint: str = "input.pdf") -> str:
    """PDFを一時的にGCSへ置いて asyncBatchAnnotate → 結果JSONをGCSから取得"""
    if not GCS_BUCKET:
        raise RuntimeError("GCS_BUCKET is required for PDF OCR.")

    # 1) upload PDF to GCS
    gcs = _gcs_client()
    bucket = gcs.bucket(GCS_BUCKET)
    uid = uuid.uuid4().hex
    in_key = f"ocr_in/{uid}/{_sanitize_filename(filename_hint)}"
    out_prefix = f"ocr_out/{uid}/"

    blob = bucket.blob(in_key)
    blob.upload_from_string(pdf_bytes, content_type="application/pdf")

    # 2) call Vision files:asyncBatchAnnotate
    token = _google_access_token()
    url = "https://vision.googleapis.com/v1/files:asyncBatchAnnotate"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {
        "requests": [{
            "inputConfig": {
                "gcsSource": {"uri": f"gs://{GCS_BUCKET}/{in_key}"},
                "mimeType": "application/pdf"
            },
            "features": [{"type": "DOCUMENT_TEXT_DETECTION"}],
            "outputConfig": {
                "gcsDestination": {"uri": f"gs://{GCS_BUCKET}/{out_prefix}"},
                "batchSize": 20
            }
        }]}
    r = requests.post(url, headers=headers, json=body, timeout=60)
    r.raise_for_status()
    op = r.json().get("name")
    if not op:
        raise RuntimeError(f"Vision async operation name missing: {r.text}")

    # 3) poll operation
    op_url = f"https://vision.googleapis.com/v1/{op}"
    deadline = time.time() + VISION_PDF_POLL_TIMEOUT_SEC
    while time.time() < deadline:
        rr = requests.get(op_url, headers=headers, timeout=30)
        rr.raise_for_status()
        j = rr.json()
        if j.get("done"):
            break
        time.sleep(VISION_PDF_POLL_INTERVAL_SEC)
    else:
        raise RuntimeError("Vision PDF OCR timeout. Increase VISION_PDF_POLL_TIMEOUT_SEC.")

    # 4) read output JSON(s) from GCS
    texts = []
    for out_blob in gcs.list_blobs(GCS_BUCKET, prefix=out_prefix):
        if not out_blob.name.lower().endswith(".json"):
            continue
        content = out_blob.download_as_text(encoding="utf-8")
        try:
            data = json.loads(content)
            for resp in data.get("responses", []):
                full = resp.get("fullTextAnnotation", {}).get("text", "")
                if full:
                    texts.append(full)
                else:
                    ann = resp.get("textAnnotations", [])
                    if ann:
                        texts.append(ann[0].get("description", ""))
        except Exception:
            pass

    return "\n".join(t for t in texts if t).strip()

# ---------- 分類先フォルダ ----------
CATEGORY_TO_FOLDER = {
    "患者リスト": "01_患者リスト",
    "実績": "02_実績",
    "同意書": "03_同意書",
    "保険証": "04_保険証",
    "請求書": "05_請求書",
    "治療報告書": "06_治療報告書",
}

def category_folder(category: str) -> str:
    mapped = CATEGORY_TO_FOLDER.get(category)
    if mapped:
        return f"{ONEDRIVE_BASE_FOLDER.rstrip('/')}/{mapped}"
    return f"{ONEDRIVE_BASE_FOLDER.rstrip('/')}/その他"

# ---------- Microsoft Graph (OneDrive) ----------
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

def graph_token() -> str:
    url = f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": MS_CLIENT_ID,
        "client_secret": MS_CLIENT_SECRET,
        "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default",
    }
    r = requests.post(url, data=data, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]

def graph_headers():
    return {"Authorization": f"Bearer {graph_token()}"}

def _drive_base() -> str:
    if ONEDRIVE_DRIVE_ID:
        return f"/drives/{ONEDRIVE_DRIVE_ID}"
    elif ONEDRIVE_USER_ID:
        return f"/users/{quote(ONEDRIVE_USER_ID)}/drive"
    else:
        # app-only だと /me は使えない構成もある点に注意
        return "/me/drive"

def ensure_folder(path: str) -> dict:
    """
    '/A/B/C' のようなパスのフォルダを（存在しなければ）順に作成。最後のフォルダを返す。
    """
    headers = graph_headers()
    parts = [p for p in path.strip("/").split("/") if p]
    base = _drive_base()
    acc_path = ""
    for part in parts:
        acc_path += "/" + part
        # 存在確認
        url = f"{GRAPH_BASE}{base}/root:{quote(acc_path, safe='/')}"
        r = requests.get(url, headers=headers, timeout=30)
        if r.status_code == 200:
            continue
        # 親に作成
        parent_path = acc_path.rsplit("/", 1)[0] or "/"
        if parent_path == "/":
            create_url = f"{GRAPH_BASE}{base}/root/children"
            # ↑typoを避けるため修正
            create_url = f"{GRAPH_BASE}{base}/root/children"
        else:
            create_url = f"{GRAPH_BASE}{base}/root:{quote(parent_path, safe='/')}:/children"
        body = {"name": part, "folder": {}, "@microsoft.graph.conflictBehavior": "replace"}
        cr = requests.post(create_url, headers={**headers, "Content-Type": "application/json"}, json=body, timeout=30)
        if cr.status_code not in (200, 201):
            raise RuntimeError(f"Failed to create folder '{acc_path}': {cr.status_code} {cr.text}")

    final_url = f"{GRAPH_BASE}{base}/root:{quote(acc_path, safe='/')}"
    fr = requests.get(final_url, headers=headers, timeout=30)
    fr.raise_for_status()
    return fr.json()

# AI_OCR.py に追加（ensure_folder の近く。Graph の GET で存在確認）
def _file_exists(path_folder: str, filename: str) -> bool:
    base = _drive_base()
    url = f"{GRAPH_BASE}{base}/root:{quote(f'{path_folder.rstrip('/')}/{filename}', safe='/')}"
    r = requests.get(url, headers=graph_headers(), timeout=15)
    return r.status_code == 200

def uniquify_filename(path_folder: str, filename: str) -> str:
    if not _file_exists(path_folder, filename):
        return filename
    base, dot, ext = filename.rpartition(".")
    base = base if dot else filename  # 拡張子なしにも対応
    ext = f".{ext}" if dot else ""
    for i in range(2, 50):
        cand = f"{base}_v{i}{ext}"
        if not _file_exists(path_folder, cand):
            return cand
    return f"{base}_{uuid.uuid4().hex[:6]}{ext}"

def upload_small(path_folder: str, filename: str, data: bytes, content_type: str) -> dict:
    """単発アップロード（~4MB）。戻り値は driveItem。"""
    headers = graph_headers()
    headers["Content-Type"] = content_type
    base = _drive_base()
    target_path = f"{path_folder.rstrip('/')}/{filename}"
    url = f"{GRAPH_BASE}{base}/root:{quote(target_path, safe='/')}:/content"
    r = requests.put(url, headers=headers, data=data, timeout=120)
    r.raise_for_status()
    return r.json()

def upload_large(path_folder: str, filename: str, data: bytes, content_type: str, chunk_size=5*1024*1024) -> dict:
    """大容量アップロード（Upload Session）。戻り値は driveItem。"""
    headers = graph_headers()
    base = _drive_base()
    target_path = f"{path_folder.rstrip('/')}/{filename}"
    session_url = f"{GRAPH_BASE}{base}/root:{quote(target_path, safe='/')}:/createUploadSession"
    s = requests.post(session_url, headers=headers, json={"item": {"@microsoft.graph.conflictBehavior": "replace"}}, timeout=30)
    s.raise_for_status()
    upload_url = s.json()["uploadUrl"]

    total = len(data)
    offset = 0
    while offset < total:
        end = min(offset + chunk_size, total)
        chunk = data[offset:end]
        headers_chunk = {"Content-Length": str(len(chunk)), "Content-Range": f"bytes {offset}-{end-1}/{total}"}
        r = requests.put(upload_url, headers=headers_chunk, data=chunk, timeout=120)
        if r.status_code in (200, 201):
            return r.json()
        elif r.status_code == 202:
            offset = end
            continue
        else:
            raise RuntimeError(f"Upload session failed: {r.status_code} {r.text}")
    raise RuntimeError("Upload session ended unexpectedly.")

def create_share_link(item_id: str, scope: str = None, link_type: str = "view") -> str:
    headers = graph_headers()
    base = _drive_base()
    url = f"{GRAPH_BASE}{base}/items/{item_id}/createLink"
    body = {"type": link_type, "scope": scope or ONEDRIVE_LINK_SCOPE}
    r = requests.post(url, headers=headers, json=body, timeout=30)
    r.raise_for_status()
    return r.json()["link"]["webUrl"]

def onedrive_search(query: str, max_items=5, allowed_ext=(".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff")) -> list:
    headers = graph_headers()
    base = _drive_base()
    url = f"{GRAPH_BASE}{base}/root/search(q='{quote(query)}')"
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    items = r.json().get("value", [])
    results = []
    for it in items:
        if "file" not in it:
            continue
        name = it.get("name", "")
        if not name.lower().endswith(allowed_ext):
            continue
        results.append(it)
        if len(results) >= max_items:
            break
    return results

# --- 追加: ファイル名の一部完全一致マッチ用ユーティリティ＆検索 ---
def _normalize_person(s: str) -> str:
    """比較用に空白類を除去（半角/全角対応）。"""
    return re.sub(r"[ \u3000\t]", "", s or "")

def _filename_token_exact_match(file_name: str, person: str) -> bool:
    """
    ファイル名(拡張子除く)をトークン分割し、どれかのトークンが person と完全一致すれば True。
    トークン区切り: _, -, スペース(半/全角), ドット, 各種括弧, スラッシュなど
    """
    if not file_name or not person:
        return False
    person_n = _normalize_person(person)

    base, _ext = os.path.splitext(file_name)
    if _normalize_person(base) == person_n:
        return True

    tokens = re.split(r"[ _\-\.\(\)【】\[\]／/　]+", base)
    for t in tokens:
        if _normalize_person(t) == person_n and person_n != "":
            return True
    return False

def onedrive_search_by_filename_exact_token(person: str, max_items: int = 5, pool: int = 50):
    """
    Graph 検索で候補を取り、クライアント側で「ファイル名の一部完全一致」で絞り込む。
    pool: サーバ側検索で取ってくる候補数（フィルタで減る前提なので少し多めに）
    """
    # ※ Graph の /search はコンテンツヒットも混ざるため、取得後に「ファイル名だけ」でフィルタする
    candidates = onedrive_search(person, max_items=pool)
    results = []
    for it in candidates:
        name = it.get("name", "")
        if _filename_token_exact_match(name, person):
            results.append(it)
            if len(results) >= max_items:
                break
    return results

# ---------- LINE Routes ----------
@app.route("/", methods=["GET"])
def health():
    return "OK", 200

@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return "OK"

# ---------- Handlers ----------
@handler.add(MessageEvent, message=ImageMessage)
def handle_image(event: MessageEvent):
    # ---- 安全な初期値（例外時にも参照できる）----
    kind = "image"
    date_str = _now_date_str()
    category = "N/A"
    patient = ""
    doctor = ""
    folder = ""
    filename = ""
    link = ""
    text = ""
    # ----------------------------------------------
    try:
        # 1) 画像取得
        content = line_bot_api.get_message_content(event.message.id)
        image_bytes = b"".join(chunk for chunk in content.iter_content())

        # 2) OCR
        text = ocr_image_bytes(image_bytes)

        # 3) 分類・抽出
        category = detect_category(text)
        patient = extract_patient(text)
        doctor = extract_doctor(text)
        extracted = extract_date(text)
        if extracted:
            date_str = extracted

        # 4) 保存先・命名
        folder = category_folder(category)
        ensure_folder(folder)

        # ★ 治療報告書は専用パーサ＆命名で高精度化
        if category == "治療報告書":
            meta = parse_meta_from_tiryo_houkokusho(text)
            filename = build_filename_treatment_report(meta, ext=".jpg")
        else:
            filename = build_filename(category, patient, doctor, date_str, ext=".jpg", text=text)

        filename = uniquify_filename(folder, filename)

        # 5) OneDriveへ保存
        if len(image_bytes) <= 4 * 1024 * 1024:
            item = upload_small(folder, filename, image_bytes, "image/jpeg")
        else:
            item = upload_large(folder, filename, image_bytes, "image/jpeg")
        link = create_share_link(item["id"])

        # 6) ★成功ログをここで追記（OCR結果が入る）
        gsheet_append_rows([[
            datetime.now().isoformat(timespec="seconds"),  # 保存日時ISO
            date_str,                                      # 保存日付YYYYMMDD
            kind,                                          # 種別
            category,                                      # 分類
            patient or "",                                 # 患者
            doctor or "",                                  # 先生
            date_str,                                      # 抽出日付
            folder,                                        # 保存フォルダ
            filename,                                      # ファイル名
            link,                                          # リンク
            str(len(text or "")),                          # OCR文字数
            (text or "").replace("\n", " ")[:100],         # OCR先頭100
            "success",                                     # ステータス
            event.message.id,                              # イベントID
            ""                                             # エラーメッセージ
        ]])

        # 7) 返信
        msg = (f"分類: {category}\n"
               f"患者: {patient or '不明'} / 先生: {doctor or '不明'} / 日付: {date_str}\n"
               f"保存先: {folder}/{filename}\n"
               f"リンク: {link}")
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))

    except Exception as e:
        # ★失敗時ログ（初期値で安全に記録）
        try:
            gsheet_append_rows([[
                datetime.now().isoformat(timespec="seconds"),
                date_str, kind, category, patient, doctor, "",
                folder, filename, link, "0", "", "error",
                getattr(event.message, "id", ""), str(e)[:800]
            ]])
        except Exception:
            pass
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"処理失敗（画像）: {e}"))

@handler.add(MessageEvent, message=FileMessage)
def handle_file(event: MessageEvent):
    # ---- 安全な初期値 ----
    kind = "pdf"
    date_str = _now_date_str()
    category = "N/A"
    patient = ""
    doctor = ""
    folder = ""
    filename = ""
    link = ""
    text = ""
    # ----------------------
    try:
        if (event.message.file_name or "").lower().endswith(".pdf"):
            # 1) PDF取得
            content = line_bot_api.get_message_content(event.message.id)
            pdf_bytes = b"".join(chunk for chunk in content.iter_content())

            # 2) OCR（Vision async + GCS）
            text = ocr_pdf_bytes_via_gcs(pdf_bytes, filename_hint=event.message.file_name or "input.pdf")

            # 3) 分類・抽出
            category = detect_category(text)
            patient = extract_patient(text)
            doctor = extract_doctor(text)
            extracted = extract_date(text)
            if extracted:
                date_str = extracted

            # 4) 保存先・命名
            folder = category_folder(category)
            ensure_folder(folder)

            # ★ 治療報告書は専用パーサ＆命名で高精度化
            if category == "治療報告書":
                meta = parse_meta_from_tiryo_houkokusho(text)
                filename = build_filename_treatment_report(meta, ext=".pdf")
            else:
                filename = build_filename(category, patient, doctor, date_str, ext=".pdf", text=text)

            filename = uniquify_filename(folder, filename)

            # 5) OneDriveへ保存
            if len(pdf_bytes) <= 4 * 1024 * 1024:
                item = upload_small(folder, filename, pdf_bytes, "application/pdf")
            else:
                item = upload_large(folder, filename, pdf_bytes, "application/pdf")
            link = create_share_link(item["id"])

            # 6) ★成功ログ
            gsheet_append_rows([[
                datetime.now().isoformat(timespec="seconds"),
                date_str, kind, category, patient or "", doctor or "", date_str,
                folder, filename, link, str(len(text or "")),
                (text or "").replace("\n", " ")[:800],
                "success", event.message.id, ""
            ]])

            # 7) 返信
            msg = (f"分類: {category}\n"
                   f"患者: {patient or '不明'} / 先生: {doctor or '不明'} / 日付: {date_str}\n"
                   f"保存先: {folder}/{filename}\n"
                   f"リンク: {link}")
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
        else:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="PDF以外のファイルは未対応です。画像はそのまま送ってください。"))

    except Exception as e:
        try:
            gsheet_append_rows([[
                datetime.now().isoformat(timespec="seconds"),
                date_str, kind, category, patient, doctor, "",
                folder, filename, link, "0", "", "error",
                getattr(event.message, "id", ""), str(e)[:800]
            ]])
        except Exception:
            pass
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"処理失敗（PDF）: {e}"))

@handler.add(MessageEvent, message=TextMessage)
def handle_text(event: MessageEvent):
    try:
        query = (event.message.text or "").strip()

        # ★ 追加: 先頭が #/＃ でなければ完全に無反応
        if not query or not re.match(r"^[#＃]", query):
            return  # 返信しない

        # ★ 「#名前 ...」のみ反応（全角＃/コロン対応）
        m = re.match(r"^[#＃]名前\s*[:：]?\s*(.+)$", query)
        if m:
            person = m.group(1).strip()
            if not person:
                line_bot_api.reply_message(
                    event.reply_token, TextSendMessage(text="使い方: #名前 佐藤太郎")
                )
                return

            # ファイル名の一部完全一致（前回追加の関数を使用）
            items = onedrive_search_by_filename_exact_token(person, max_items=5, pool=50)
            if not items:
                line_bot_api.reply_message(
                    event.reply_token,
                    TextSendMessage(text=f"【名前検索】「{person}」に一致するファイル名は見つかりませんでした。"),
                )
                return

            lines = []
            for it in items:
                name = it.get("name", "(no name)")
                item_id = it.get("id")
                try:
                    link = create_share_link(item_id)
                except Exception:
                    link = it.get("webUrl", "")
                lines.append(f"• {name}\n  {link}")

            reply = "【名前検索】ファイル名の一部完全一致（最大5件）\n" + "\n".join(lines)
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))
            return

        # ★ それ以外の #コマンドは無視（返信しない）
        return

    except Exception as e:
        # #コマンド処理中のみエラーを返す（#なしはそもそも無反応）
        try:
            if re.match(r"^[#＃]", (event.message.text or "")):
                line_bot_api.reply_message(
                    event.reply_token, TextSendMessage(text=f"検索処理失敗: {e}")
                )
        except Exception:
            pass

# ---------- Local run ----------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "3000"))
    app.run(host="0.0.0.0", port=port)
