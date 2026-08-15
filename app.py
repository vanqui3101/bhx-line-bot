"""
app.py - Webhook server LINE bot báo cáo doanh thu (bản 2)

Luồng hoạt động:
1. Người dùng gửi file Excel doanh thu (nhiều dòng, mỗi dòng 1 siêu thị)
   -> Bot đọc, lưu vào SQLite theo ngày, trả lời xác nhận.
2. Người dùng gõ lệnh "báo cáo doanh thu"
   -> Bot lấy dữ liệu mới nhất + kỳ trước, trả về Flex Message (thẻ đẹp)
      kèm link tải file Excel chi tiết đầy đủ tất cả siêu thị.

CẤU HÌNH (Environment Variables):
- LINE_CHANNEL_ACCESS_TOKEN
- LINE_CHANNEL_SECRET
"""

import os
import re
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
    FlexMessage,
    FlexContainer,
)
from linebot.v3.webhooks import MessageEvent, FileMessageContent, TextMessageContent

from excel_reader import read_all_rows
from flex_builder import build_flex_message
from excel_report import build_detail_excel
import storage

CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET", "")

TMP_DIR = os.path.join(os.path.dirname(__file__), "tmp")
os.makedirs(TMP_DIR, exist_ok=True)

REPORTS_DIR = os.path.join(os.path.dirname(__file__), "reports")
os.makedirs(REPORTS_DIR, exist_ok=True)

app = Flask(__name__)
configuration = Configuration(access_token=CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)

REPORT_COMMAND_PATTERN = re.compile(r"b[aá]o\s*c[aá]o\s*doanh\s*thu", re.IGNORECASE)


@app.route("/", methods=["GET"])
def health():
    return "LINE bot báo cáo doanh thu (bản 2) đang chạy.", 200


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
            reply_text(messaging_api, event.reply_token, "Bot chỉ đọc được file Excel (.xlsx). Anh gửi lại file định dạng .xlsx giúp em nhé.")
            return

        try:
            content = blob_api.get_message_content(message_id)
            tmp_path = os.path.join(TMP_DIR, f"input_{uuid.uuid4().hex}.xlsx")
            with open(tmp_path, "wb") as f:
                f.write(content)

            date_str, rows = read_all_rows(tmp_path)
            os.remove(tmp_path)

            storage.save_records(rows)

            from datetime import datetime as _dt
            try:
                from zoneinfo import ZoneInfo
                now_vn = _dt.now(ZoneInfo("Asia/Ho_Chi_Minh"))
            except Exception:
                now_vn = _dt.now()
            gio_str = now_vn.strftime("%H:%M")
            storage.save_snapshot_time(date_str, gio_str)

            date_display = _dt.strptime(date_str, "%Y-%m-%d").strftime("%d/%m/%Y")
            total = sum((r["dt_offline"] or 0) + (r["dt_online"] or 0) for r in rows)
            total_str = f"{total:,.0f}".replace(",", ".")
            reply = f"✅ Đã lưu dữ liệu ngày {date_display} cho {len(rows)} siêu thị.\nTổng doanh thu: {total_str} đ\n\nGõ \"Báo cáo doanh thu\" để xem báo cáo chi tiết."
            reply_text(messaging_api, event.reply_token, reply)

        except Exception as e:
            traceback.print_exc()
            reply_text(messaging_api, event.reply_token, f"Có lỗi khi xử lý file: {e}")


@handler.add(MessageEvent, message=TextMessageContent)
def handle_text_message(event):
    text = event.message.text or ""
    if not REPORT_COMMAND_PATTERN.search(text):
        return

    with ApiClient(configuration) as api_client:
        messaging_api = MessagingApi(api_client)
        try:
            latest_date, latest_records, prev_date, prev_records = storage.get_latest_and_previous()

            if latest_date is None:
                reply_text(
                    messaging_api,
                    event.reply_token,
                    "Chưa có dữ liệu nào được lưu. Anh gửi file Excel doanh thu trước nhé.",
                )
                return

            gio_now = storage.get_snapshot_time(latest_date)
            gio_prev = storage.get_snapshot_time(prev_date)

            bubble = build_flex_message(latest_date, latest_records, prev_date, prev_records, gio_now, gio_prev)
            flex_message = FlexMessage(
                alt_text=f"Báo cáo doanh thu {latest_date}",
                contents=FlexContainer.from_dict(bubble),
            )

            report_filename = f"chi_tiet_{uuid.uuid4().hex[:8]}.xlsx"
            report_path = os.path.join(REPORTS_DIR, report_filename)
            build_detail_excel(latest_date, latest_records, prev_date, prev_records, report_path)
            download_url = f"{request.host_url.rstrip('/')}/reports/{report_filename}"
            link_text = f"📎 File Excel chi tiết đầy đủ (offline/online/tháng trước/chênh lệch từng siêu thị):\n{download_url}"

            messaging_api.reply_message(
                ReplyMessageRequest(
                    reply_token=event.reply_token,
                    messages=[flex_message, TextMessage(text=link_text)],
                )
            )
        except Exception as e:
            traceback.print_exc()
            reply_text(messaging_api, event.reply_token, f"Có lỗi khi tạo báo cáo: {e}")


def reply_text(messaging_api, reply_token, text):
    text = text[:4900]
    messaging_api.reply_message(
        ReplyMessageRequest(reply_token=reply_token, messages=[TextMessage(text=text)])
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
