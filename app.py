"""
app.py - Webhook server cho LINE bot báo cáo doanh thu.

Luồng hoạt động:
1. Người dùng gửi file Excel (.xlsx) vào LINE bot.
2. Bot tải file về, đọc dữ liệu, tạo báo cáo gọn (report_generator.py).
3. Bot lưu file báo cáo vào thư mục /reports, trả lời bằng:
   - Tin nhắn text tóm tắt
   - Link tải file Excel báo cáo

CẤU HÌNH BẮT BUỘC (đặt trong Environment Variables khi deploy):
- LINE_CHANNEL_ACCESS_TOKEN : lấy từ LINE Developers Console
- LINE_CHANNEL_SECRET       : lấy từ LINE Developers Console
- PUBLIC_BASE_URL           : URL công khai của server sau khi deploy
                               (vd: https://ten-app-cua-anh.up.railway.app)
"""

import os
import uuid
import traceback

from flask import Flask, request, abort, send_from_directory

from linebot.v3 import WebhookHandler
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.messaging import (
    Configuration,
    ApiClient,
    MessagingApi,
    MessagingApiBlob,
    ReplyMessageRequest,
    TextMessage,
)
from linebot.v3.webhooks import MessageEvent, FileMessageContent

from report_generator import read_raw_data, build_report, build_text_summary

CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET", "")
PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "").rstrip("/")

REPORTS_DIR = os.path.join(os.path.dirname(__file__), "reports")
os.makedirs(REPORTS_DIR, exist_ok=True)

app = Flask(__name__)
configuration = Configuration(access_token=CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)


@app.route("/", methods=["GET"])
def health():
    return "LINE bot báo cáo doanh thu đang chạy.", 200


@app.route("/reports/<path:filename>", methods=["GET"])
def download_report(filename):
    return send_from_directory(REPORTS_DIR, filename, as_attachment=True)


@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return "OK"


@handler.add(MessageEvent, message=FileMessageContent)
def handle_file_message(event):
    message_id = event.message.id
    original_name = event.message.file_name or "input.xlsx"

    with ApiClient(configuration) as api_client:
        blob_api = MessagingApiBlob(api_client)
        messaging_api = MessagingApi(api_client)

        if not original_name.lower().endswith((".xlsx", ".xlsm")):
            reply_text(
                messaging_api,
                event.reply_token,
                "Bot chỉ đọc được file Excel (.xlsx). Anh gửi lại file định dạng .xlsx giúp em nhé.",
            )
            return

        try:
            content = blob_api.get_message_content(message_id)

            tmp_input = os.path.join(REPORTS_DIR, f"input_{uuid.uuid4().hex}.xlsx")
            with open(tmp_input, "wb") as f:
                f.write(content)

            data = read_raw_data(tmp_input)

            report_filename = f"bao_cao_{uuid.uuid4().hex[:8]}.xlsx"
            report_path = os.path.join(REPORTS_DIR, report_filename)
            summary = build_report(data, report_path)

            os.remove(tmp_input)

            text_summary = build_text_summary(summary)
            download_url = f"{PUBLIC_BASE_URL}/reports/{report_filename}"
            reply = f"{text_summary}\n\n📎 Tải file Excel: {download_url}"

            reply_text(messaging_api, event.reply_token, reply)

        except Exception as e:
            traceback.print_exc()
            reply_text(
                messaging_api,
                event.reply_token,
                f"Có lỗi khi xử lý file: {e}\nAnh kiểm tra lại file có đúng định dạng doanh thu theo siêu thị không.",
            )


def reply_text(messaging_api, reply_token, text):
    # LINE giới hạn tin nhắn text tối đa 5000 ký tự
    text = text[:4900]
    messaging_api.reply_message(
        ReplyMessageRequest(
            reply_token=reply_token,
            messages=[TextMessage(text=text)],
        )
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
