import os
import json
import base64
from datetime import datetime

import requests
from requests.exceptions import HTTPError
from flask import Flask, request, abort

from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError, LineBotApiError
from linebot.models import MessageEvent, ImageMessage, TextSendMessage

# Google Sheets (gspread)
import gspread
from google.oauth2.service_account import Credentials

import re


# ---------- Config ----------
CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET", "")
VISION_API_KEY = os.environ.get("GOOGLE_VISION_API_KEY", "")
SPREADSHEET_KEY = os.environ.get("SPREADSHEET_KEY", "")
SERVICE_ACCOUNT_VALUE = os.environ.get("GCP_SERVICE_ACCOUNT_JSON", "").strip()
WORKSHEET_NAME = os.environ.get("WORKSHEET_NAME", "OCR")

if not (CHANNEL_ACCESS_TOKEN and CHANNEL_SECRET and VISION_API_KEY and SPREADSHEET_KEY and SERVICE_ACCOUNT_VALUE):
    missing = [k for k, v in {
        "LINE_CHANNEL_ACCESS_TOKEN": CHANNEL_ACCESS_TOKEN,
        "LINE_CHANNEL_SECRET": CHANNEL_SECRET,
        "GOOGLE_VISION_API_KEY": VISION_API_KEY,
        "SPREADSHEET_KEY": SPREADSHEET_KEY,
        "GCP_SERVICE_ACCOUNT_JSON": SERVICE_ACCOUNT_VALUE,
    }.items() if not v]
    raise RuntimeError(f"Missing environment variables: {', '.join(missing)}")

# ---------- App & LINE setup ----------
app = Flask(__name__)
line_bot_api = LineBotApi(CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)

# ---------- Helpers ----------
def _load_service_account_info(value: str) -> dict:
    """
    Accepts either:
      - JSON string of service account
      - File path to a JSON file
    Returns parsed dict.
    """
    # Try as JSON string
    try:
        return json.loads(value)
    except json.JSONDecodeError:
        # Treat as file path
        path = os.path.expanduser(value)
        if not os.path.isfile(path):
            raise RuntimeError(
                "GCP_SERVICE_ACCOUNT_JSON must be a JSON string or an existing file path."
            )
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

# ---------- Google Sheets client ----------
def get_gspread_client():
    info = _load_service_account_info(SERVICE_ACCOUNT_VALUE)
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
        ws = sh.add_worksheet(title=WORKSHEET_NAME, rows=200, cols=10)
        ws.append_row(
            ["timestamp", "user_id", "line_image_message_id", "ocr_text",
             "amount", "tx_date", "vendor", "category", "account_code", "confidence"],
            value_input_option="RAW"
        )

    # 追加: OCRテキストの後処理
    amount = extract_amount(ocr_text)
    tx_date = extract_date(ocr_text)
    vendor = extract_vendor(ocr_text)
    category, account_code, confidence = classify_category(ocr_text)

    ts = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
    ws.append_row(
        [
            ts, user_id, image_id, ocr_text,
            amount if amount is not None else "",
            tx_date if tx_date else "",
            vendor if vendor else "",
            category, account_code, round(confidence, 2)
        ],
        value_input_option="RAW"
    )


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
    # ここで Vision 側のレスポンスエラーを明確化
    try:
        resp.raise_for_status()
    except HTTPError as he:
        body = resp.text[:300] + "..." if resp is not None and resp.text else ""
        raise RuntimeError(f"Vision API error: HTTP {resp.status_code} {resp.reason} {body}") from he

    data = resp.json()
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
        # 1) LINEから画像取得（ここでの404を見分ける）
        try:
            content = line_bot_api.get_message_content(event.message.id)
            image_bytes = b"".join(chunk for chunk in content.iter_content())
        except LineBotApiError as le:
            status = getattr(le, "status_code", "unknown")
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(
                    text=f"画像取得に失敗しました（LINE側）: status={status}. "
                         f"チャネルのアクセストークン/シークレットがWebhookのチャネルと一致しているか確認してください。"
                )
            )
            return
        except HTTPError as he:
            resp = getattr(he, "response", None)
            detail = f"HTTP {resp.status_code}" if resp is not None else str(he)
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=f"画像取得に失敗しました（HTTP層）: {detail}")
            )
            return

        # 2) VisionでOCR
        try:
            text = ocr_image_by_vision(image_bytes)
        except Exception as e:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=f"OCR実行に失敗しました（Vision）: {e}")
            )
            return

        # 3) シート保存（サービスアカウント権限エラーもここで拾える）
        try:
            user_id = getattr(event.source, "user_id", "")
            append_ocr_to_sheet(user_id=user_id, ocr_text=text, image_id=event.message.id)
        except Exception as e:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(
                    text=f"スプレッドシート保存に失敗しました: {e}\n"
                         f"・対象シートがサービスアカウント（client_email）に共有されているか\n"
                         f"・SPREADSHEET_KEY が正しいか を確認してください。"
                )
            )
            return

        # 4) 成功応答
        preview = (text or "").strip()
        if not preview:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="画像から文字が読み取れませんでした。もう一度お試しください。")
            )
        else:
            if len(preview) > 900:
                preview = preview[:900] + "…"
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="OCRで読み取り、スプレッドシートに保存しました。\n---\n" + preview)
            )

    except Exception as e:
        # 想定外の例外
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=f"想定外のエラー: {e}")
        )


# OCRテキストの抽出・分類ユーティリティ
AMOUNT_PAT = re.compile(r"(¥|￥)?\s*([0-9]{1,3}(?:[,，][0-9]{3})+|[0-9]+)\s*(円)?")
DATE_PAT = re.compile(r"(?:(20\d{2})[./年-](\d{1,2})[./月-](\d{1,2})日?)")
VENDOR_PAT = re.compile(r"(株式会社|有限会社|カブシキガイシャ)[^\s\n\r]{1,20}|[A-Za-z0-9&'().\-]{3,}\s?(?:Store|SHOP|Cafe|CO|INC|LTD)\b", re.IGNORECASE)

# キーワード規則（必要に応じて追加/調整）
CATEGORY_RULES = [
    (r"交通|電車|乗車|切符|運賃|タクシ|駐車|駐輪|高速|ETC",          ("旅費交通費",      "6611")),
    (r"通信|モバイル|携帯|インターネット|Wi-?Fi|電話|回線|SIM",        ("通信費",          "6213")),
    (r"交際|懇親|接待|会食|贈答|お土産|祝電",                         ("交際費",          "8121")),
    (r"消耗品|文具|事務用品|インク|トナー|備品|テプラ|コピー紙",       ("消耗品費",        "6261")),
    (r"会議|打合|打ち合わせ|ミーティング|会場費|レンタルスペース",     ("会議費",          "6222")),
    (r"広告|宣伝|プロモ|チラシ|フライヤ|SNS広告|リスティング|出稿",     ("広告宣伝費",      "6111")),
    (r"水道|電気|ガス|光熱|電力|検針票",                               ("水道光熱費",      "6221")),
    (r"配送料|宅配|運送|クール便|郵送|ゆうパック|ネコポス|レターパック", ("荷造運賃",        "6241")),
]

def extract_amount(text: str) -> float | None:
    amounts = []
    for m in AMOUNT_PAT.finditer(text.replace(",", "").replace("，", "")):
        try:
            amounts.append(float(m.group(2)))
        except Exception:
            pass
    return max(amounts) if amounts else None

def extract_date(text: str) -> str | None:
    m = DATE_PAT.search(text)
    if not m:
        return None
    y, mo, d = m.group(1), m.group(2), m.group(3)
    return f"{int(y):04d}-{int(mo):02d}-{int(d):02d}"

def extract_vendor(text: str) -> str | None:
    m = VENDOR_PAT.search(text)
    return m.group(0).strip() if m else None

def classify_category(text: str) -> tuple[str, str, float]:
    """
    ルールで大分類と科目コードを決める。
    return: (category, account_code, confidence)
    """
    text_norm = text.replace("　", " ")
    for pat, (cat, code) in CATEGORY_RULES:
        if re.search(pat, text_norm):
            # マッチ数で簡易スコア
            hits = len(re.findall(pat, text_norm))
            conf = min(0.6 + 0.2 * hits, 0.95)
            return cat, code, conf
    return "雑費", "8899", 0.3



# ---------- Local run ----------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "3000"))
    app.run(host="0.0.0.0", port=port)
