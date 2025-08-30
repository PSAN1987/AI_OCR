# LINE_BOT_OCR_OneDrive.py
# -*- coding: utf-8 -*-
"""
LINE Bot with AI OCR (Google Vision) + OneDrive integration

A) 受信自動化
   - 画像（ImageMessage）/ PDF（FileMessage）を受信
   - Google VisionでOCR
     * 画像: images:annotate（APIキー or サービスアカウントのどちらでも可）
     * PDF : files:asyncBatchAnnotate（サービスアカウント + GCS必須）
   - OCRテキストからカテゴリ判定・氏名抽出・先生名抽出・保存日付決定
   - 命名規則に従って OneDrive に分類保存（/同意書, /保険証, /治療報告書, /実績, /請求書）
   - 共有リンクを作成して LINE で保存結果と保存先を返信

B) 送信自動化（試験）
   - ユーザーがテキストで「名前」を送信
   - OneDrive 全階層を検索（ファイル名ヒット）→ 一致ファイルのリンクを返信

環境変数（必須/推奨）:
  LINE_CHANNEL_ACCESS_TOKEN
  LINE_CHANNEL_SECRET

  # Google / Vision
  GCP_SERVICE_ACCOUNT_JSON         # サービスアカウントのJSONそのもの or JSONファイルパス
  GOOGLE_VISION_API_KEY            # 任意（画像OCRで利用。未設定ならSA OAuthで実行）
  GCS_BUCKET                       # PDF OCRで使用する一時バケット名（gs://<bucket> ではなく <bucket> 名）

  # Microsoft Graph / OneDrive (アプリケーション権限を推奨)
  MS_TENANT_ID
  MS_CLIENT_ID
  MS_CLIENT_SECRET
  ONEDRIVE_DRIVE_ID                # 推奨: /drives/{drive-id}
  # もしくは
  # ONEDRIVE_USER_ID               # /users/{user-or-upn}/drive（DRIVE_ID未設定時に使用）

  # 任意
  ONEDRIVE_BASE_FOLDER=/           # 保存ルート（既定 "/"）
  ONEDRIVE_LINK_SCOPE=organization # createLink の scope（organization / anonymous / users）
  VISION_PDF_POLL_TIMEOUT_SEC=90   # PDFの非同期OCR待機秒
  VISION_PDF_POLL_INTERVAL_SEC=3   # PDFの非同期OCRポーリング間隔（秒）
"""

import os
import re
import io
import json
import time
import uuid
import base64
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
    return datetime.now().strftime("%Y%m%d")

def _sanitize_filename(name: str) -> str:
    # OneDrive禁止文字: \ / : * ? " < > | など
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    name = name.replace("\n", " ").replace("\r", " ")
    return name.strip() or "unnamed"

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
            # data['responses'] is an array of page responses
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

# ---------- Classification / Extraction ----------
CAT_RULES = [
    ("同意書",   r"同意|承諾|Consent|署名|サイン"),
    ("保険証",   r"保険証|被保険者|保険者番号|記号|番号|有効期限"),
    ("治療報告書", r"治療報告書|報告書|診療報酬|所見|診断|経過"),
    ("実績",     r"治療実績|実績|処置|施術|点数|算定"),
    ("請求書",   r"請求書|請求|請求金額|振込先|お振込|INVOICE|請求合計"),
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
    # どれにも該当しない場合の素朴なフォールバック
    return ("請求書" if re.search(r"請求", t) else
            "実績" if re.search(r"実績", t) else
            "治療報告書" if re.search(r"報告", t) else
            "同意書" if re.search(r"同意", t) else
            "保険証" if re.search(r"保険証", t) else "その他")

def extract_date(text: str):
    for p in DATE_PATTERNS:
        m = re.search(p, text)
        if m:
            y, mo, d = m.groups()
            return f"{int(y):04d}{int(mo):02d}{int(d):02d}"
    return None

def extract_patient(text: str):
    for label in ["患者氏名", "患者名", "氏名", "お名前", "患者", "被保険者氏名"]:
        m = re.search(label + r"\s*[:：]?\s*([^\n\r\t 　]{2,20})", text)
        if m:
            return m.group(1).strip()
    m = re.search(r"([^\s\n\r]{2,20})\s*様", text)
    return m.group(1).strip() if m else None

def extract_doctor(text: str):
    for label in ["医師名", "担当医", "担当歯科医師", "歯科医師", "先生", "Dr", "Doctor"]:
        m = re.search(label + r"\s*[:：]?\s*([^\n\r\t 　]{2,20})", text, flags=re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return None

def build_filename(category: str, patient, doctor, date_str, ext: str) -> str:
    p = _sanitize_filename(patient or "不明")
    d = _sanitize_filename(doctor or "不明")
    dt = date_str or _now_date_str()
    if category == "同意書":
        return f"同意書_{p}_{d}_{dt}{ext}"
    if category == "保険証":
        return f"保険証_{p}_{dt}{ext}"
    if category == "治療報告書":
        return f"治療報告書_{p}_{d}_{dt}{ext}"
    if category == "実績":
        return f"実績_{p}_{d}_{dt}{ext}"
    if category == "請求書":
        return f"請求書_{p}_{dt}{ext}"
    return f"{category}_{p}_{dt}{ext}"

def category_folder(category: str) -> str:
    if category in ["同意書", "保険証", "治療報告書", "実績", "請求書"]:
        return f"{ONEDRIVE_BASE_FOLDER.rstrip('/')}/{category}"
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
    try:
        # 1) 画像バイト取得
        content = line_bot_api.get_message_content(event.message.id)
        image_bytes = b"".join(chunk for chunk in content.iter_content())

        # 2) OCR
        text = ocr_image_bytes(image_bytes)

        # 3) 分類＆命名
        category = detect_category(text)
        patient = extract_patient(text)
        doctor = extract_doctor(text)
        date_str = extract_date(text) or _now_date_str()
        folder = category_folder(category)
        ensure_folder(folder)
        filename = build_filename(category, patient, doctor, date_str, ext=".jpg")

        # 4) OneDriveへ保存
        if len(image_bytes) <= 4 * 1024 * 1024:
            item = upload_small(folder, filename, image_bytes, "image/jpeg")
        else:
            item = upload_large(folder, filename, image_bytes, "image/jpeg")
        link = create_share_link(item["id"])

        # 5) 返信
        msg = (f"分類: {category}\n"
               f"患者: {patient or '不明'} / 先生: {doctor or '不明'} / 日付: {date_str}\n"
               f"保存先: {folder}/{filename}\n"
               f"リンク: {link}")
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))

    except Exception as e:
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"処理失敗（画像）: {e}"))

@handler.add(MessageEvent, message=FileMessage)
def handle_file(event: MessageEvent):
    try:
        if (event.message.file_name or "").lower().endswith(".pdf"):
            # 1) PDFバイト取得
            content = line_bot_api.get_message_content(event.message.id)
            pdf_bytes = b"".join(chunk for chunk in content.iter_content())

            # 2) OCR（Vision async + GCS）
            text = ocr_pdf_bytes_via_gcs(pdf_bytes, filename_hint=event.message.file_name or "input.pdf")

            # 3) 分類＆命名
            category = detect_category(text)
            patient = extract_patient(text)
            doctor = extract_doctor(text)
            date_str = extract_date(text) or _now_date_str()
            folder = category_folder(category)
            ensure_folder(folder)
            filename = build_filename(category, patient, doctor, date_str, ext=".pdf")

            # 4) OneDriveへ保存（PDF）
            if len(pdf_bytes) <= 4 * 1024 * 1024:
                item = upload_small(folder, filename, pdf_bytes, "application/pdf")
            else:
                item = upload_large(folder, filename, pdf_bytes, "application/pdf")
            link = create_share_link(item["id"])

            # 5) 返信
            msg = (f"分類: {category}\n"
                   f"患者: {patient or '不明'} / 先生: {doctor or '不明'} / 日付: {date_str}\n"
                   f"保存先: {folder}/{filename}\n"
                   f"リンク: {link}")
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
        else:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="PDF以外のファイルは未対応です。画像はそのまま送ってください。"))
    except Exception as e:
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"処理失敗（PDF）: {e}"))

@handler.add(MessageEvent, message=TextMessage)
def handle_text(event: MessageEvent):
    try:
        query = (event.message.text or "").strip()
        if not query:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="検索キーワード（氏名など）を入力してください。"))
            return

        # OneDrive 検索（ファイル名ヒット）
        items = onedrive_search(query, max_items=5)
        if not items:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"「{query}」に一致するファイルは見つかりませんでした。"))
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
        reply = "検索結果（最大5件）:\n" + "\n".join(lines)
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

    except Exception as e:
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"検索処理失敗: {e}"))

# ---------- Local run ----------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "3000"))
    app.run(host="0.0.0.0", port=port)
