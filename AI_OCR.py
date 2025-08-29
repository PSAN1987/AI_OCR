import os
import io
import json
import base64
from datetime import datetime

import requests
from flask import Flask, request, abort

from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, ImageMessage, TextSendMessage

# Google Sheets (gspread)
import gspread
from google.oauth2.service_account import Credentials

# ---------- Config ----------
CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET", "")
VISION_API_KEY = os.environ.get("GOOGLE_VISION_API_KEY", "")
SPREADSHEET_KEY = os.environ.get("SPREADSHEET_KEY", "")
SERVICE_ACCOUNT_JSON = os.environ.get("GCP_SERVICE_ACCOUNT_JSON", "")
WORKSHEET_NAME = os.environ.get("WORKSHEET_NAME", "OCR")

if not (CHANNEL_ACCESS_TOKEN and CHANNEL_SECRET and VISION_API_KEY and SPREADSHEET_KEY and SERVICE_ACCOUNT_JSON):
    missing = [k for k, v in {
        "LINE_CHANNEL_ACCESS_TOKEN": CHANNEL_ACCESS_TOKEN,
        "LINE_CHANNEL_SECRET": CHANNEL_SECRET,
        "GOOGLE_VISION_API_KEY": VISION_API_KEY,
        "SPREADSHEET_KEY": SPREADSHEET_KEY,
        "GCP_SERVICE_ACCOUNT_JSON": SERVICE_ACCOUNT_JSON,
    }.items() if not v]
    raise RuntimeError(f"Missing environment variables: {', '.join(missing)}")

# ---------- App & LINE setup ----------
app = Flask(__name__)
line_bot_api = LineBotApi(CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)

# ---------- Google Sheets client ----------
def get_gspread_client():
    try:
        info = json.loads(SERVICE_ACCOUNT_JSON)
    except json.JSONDecodeError:
        raise RuntimeError("GCP_SERVICE_ACCOUNT_JSON must be a valid JSON string.")
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds)

def append_ocr_to_sheet(user_id: str, ocr_text: str, image_id: str = "") -> None:
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_KEY)
    try:
        ws = sh.worksheet(WORKSHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=WORKSHEET_NAME, rows=100, cols=4)
        ws.append_row(["timestamp", "user_id", "line_image_message_id", "ocr_text"], value_input_option="RAW")
    ts = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
    ws.append_row([ts, user_id, image_id, ocr_text], value_input_option="RAW")

# ---------- OCR (Google Vision API) ----------
def ocr_image_by_vision(image_bytes: bytes) -> str:
    url = f"https://vision.googleapis.com/v1/images:annotate?key={VISION_API_KEY}"
    payload = {
        "requests": [{
            "image": {"content": base64.b64encode(image_bytes).decode("utf-8")},
            "features": [{"type": "DOCUMENT_TEXT_DETECTION"}],
            "imageContext": {"languageHints": ["ja", "en"]}
        }]
    }
    resp = requests.post(url, json=payload, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    # Prefer fullTextAnnotation, fallback to textAnnotations[0].description
    try:
        return data["responses"][0]["fullTextAnnotation"]["text"]
    except Exception:
        ann = data["responses"][0].get("textAnnotations", [])
        return ann[0]["description"] if ann else ""

# ---------- Routes ----------
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
        # 1) Get image bytes from LINE
        content = line_bot_api.get_message_content(event.message.id)
        image_bytes = b"".join(chunk for chunk in content.iter_content())

        # 2) OCR
        text = ocr_image_by_vision(image_bytes)

        # 3) Save to Google Sheets
        user_id = getattr(event.source, "user_id", "")
        append_ocr_to_sheet(user_id=user_id, ocr_text=text, image_id=event.message.id)

        # 4) Reply to user
        if not text.strip():
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="画像から文字が読み取れませんでした。もう一度お試しください。")
            )
        else:
            preview = text.strip()
            if len(preview) > 900:
                preview = preview[:900] + "…"
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="OCRで読み取ってスプレッドシートに保存しました。\n---\n" + preview)
            )
    except Exception as e:
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=f"OCR処理でエラーが発生しました: {e}")
        )

# ---------- Local run ----------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "3000"))
    app.run(host="0.0.0.0", port=port)

