"""
app.py - Webhook server LINE bot báo cáo doanh thu (bản 3)

Luồng hoạt động:
1. Người dùng gửi file Excel doanh thu (nhiều dòng, mỗi dòng 1 siêu thị)
   -> Bot đọc, lưu vào SQLite theo ngày, trả lời xác nhận
   -> Nếu đã cấu hình GROUP_ID: tự động gửi luôn báo cáo (thẻ đẹp + file Excel) vào nhóm.
2. Người dùng gõ lệnh "báo cáo doanh thu" (ở bất kỳ đâu, kể cả chat riêng)
   -> Bot lấy dữ liệu mới nhất, trả về báo cáo.
3. Gõ "id nhóm" trong 1 nhóm -> bot trả về Group ID của nhóm đó (để cấu hình GROUP_ID).
4. Mỗi ngày lúc giờ cấu hình (mặc định 08:00, giờ VN) -> tự động gửi báo cáo vào nhóm.

CẤU HÌNH (Environment Variables):
- LINE_CHANNEL_ACCESS_TOKEN  (bắt buộc)
- LINE_CHANNEL_SECRET        (bắt buộc)
- GROUP_ID                   (ID nhóm để tự động gửi báo cáo - lấy bằng lệnh "id nhóm")
- DAILY_REPORT_HOUR          (giờ gửi tự động mỗi ngày, mặc định 8)
- DAILY_REPORT_MINUTE        (phút gửi tự động mỗi ngày, mặc định 0)
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
    PushMessageRequest,
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
GROUP_ID = os.environ.get("GROUP_ID", "").strip()
DAILY_HOUR = int(os.environ.get("DAILY_REPORT_HOUR", "8"))
DAILY_MINUTE = int(os.environ.get("DAILY_REPORT_MINUTE", "0"))

TMP_DIR = os.path.join(os.path.dirname(__file__), "tmp")
os.makedirs(TMP_DIR, exist_ok=True)

REPORTS_DIR = os.path.join(os.path.dirname(__file__), "reports")
os.makedirs(REPORTS_DIR, exist_ok=True)

app = Flask(__name__)
configuration = Configuration(access_token=CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)

REPORT_COMMAND_PATTERN = re.compile(r"b[aá]o\s*c[aá]o\s*doanh\s*thu", re.IGNORECASE)
GROUP_ID_COMMAND_PATTERN = re.compile(r"^id\s*nh[oó]m$", re.IGNORECASE)

PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "").rstrip("/")


@app.route("/", methods=["GET"])
def health():
    return "LINE bot báo cáo doanh thu (bản 3) đang chạy.", 200


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


def _base_url():
    if PUBLIC_BASE_URL:
        return PUBLIC_BASE_URL
    return request.host_url.rstrip("/")


def build_report_messages(base_url):
    """Tạo (flex_message, text_message) báo cáo mới nhất, hoặc None nếu chưa có dữ liệu."""
    latest_date, latest_records, prev_date, prev_records = storage.get_latest_and_previous()
    if latest_date is None:
        return None

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
    download_url = f"{base_url}/reports/{report_filename}"
    link_text = f"📎 File Excel chi tiết đầy đủ (offline/online/tháng trước/chênh lệch từng siêu thị):\n{download_url}"

    return flex_message, TextMessage(text=link_text)


@handler.add(MessageEvent, message=FileMessageContent)
def handle_file_message(event):
    message_id = event.message.id
    original_name = event.message.file_name or "input.xlsx"
    source_type = event.source.type
    group_id = event.source.group_id if source_type == "group" else None

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
            reply = f"✅ Đã lưu dữ liệu ngày {date_display} cho {len(rows)} siêu thị.\nTổng doanh thu: {total_str} đ"
            reply_text(messaging_api, event.reply_token, reply)

            # Tự động gửi báo cáo đầy đủ vào nhóm đã cấu hình (nếu có)
            if GROUP_ID:
                try:
                    result = build_report_messages(_base_url())
                    if result:
                        flex_message, link_message = result
                        messaging_api.push_message(
                            PushMessageRequest(to=GROUP_ID, messages=[flex_message, link_message])
                        )
                except Exception:
                    traceback.print_exc()

        except Exception as e:
            traceback.print_exc()
            reply_text(messaging_api, event.reply_token, f"Có lỗi khi xử lý file: {e}")


@handler.add(MessageEvent, message=TextMessageContent)
def handle_text_message(event):
    text = event.message.text or ""

    with ApiClient(configuration) as api_client:
        messaging_api = MessagingApi(api_client)

        # Lệnh lấy Group ID
        if GROUP_ID_COMMAND_PATTERN.match(text.strip()):
            if event.source.type == "group":
                reply_text(messaging_api, event.reply_token, f"Group ID của nhóm này:\n{event.source.group_id}")
            else:
                reply_text(messaging_api, event.reply_token, "Lệnh này chỉ dùng được trong nhóm (group) nhé anh.")
            return

        if not REPORT_COMMAND_PATTERN.search(text):
            return

        user_id = getattr(event.source, "user_id", None)
        group_id = event.source.group_id if event.source.type == "group" else None
        target_id = group_id or user_id

        # Trả lời ngay lập tức để không bị hết hạn "reply token" của LINE.
        try:
            reply_text(messaging_api, event.reply_token, "⏳ Đang tổng hợp báo cáo, đợi anh vài giây nhé...")
        except Exception:
            traceback.print_exc()

        try:
            result = build_report_messages(_base_url())
            if result is None:
                push_text(messaging_api, target_id, "Chưa có dữ liệu nào được lưu. Anh gửi file Excel doanh thu trước nhé.")
                return
            flex_message, link_message = result
            messaging_api.push_message(
                PushMessageRequest(to=target_id, messages=[flex_message, link_message])
            )
        except Exception as e:
            traceback.print_exc()
            try:
                push_text(messaging_api, target_id, f"Có lỗi khi tạo báo cáo: {e}")
            except Exception:
                traceback.print_exc()


def push_text(messaging_api, target_id, text):
    text = text[:4900]
    messaging_api.push_message(
        PushMessageRequest(to=target_id, messages=[TextMessage(text=text)])
    )


def reply_text(messaging_api, reply_token, text):
    text = text[:4900]
    messaging_api.reply_message(
        ReplyMessageRequest(reply_token=reply_token, messages=[TextMessage(text=text)])
    )


def send_scheduled_report():
    """Được gọi tự động mỗi ngày theo lịch để gửi báo cáo vào nhóm."""
    if not GROUP_ID:
        return
    try:
        with ApiClient(configuration) as api_client:
            messaging_api = MessagingApi(api_client)
            base_url = PUBLIC_BASE_URL or "https://web-production-1fd9b8.up.railway.app"
            result = build_report_messages(base_url)
            if result:
                flex_message, link_message = result
                messaging_api.push_message(
                    PushMessageRequest(to=GROUP_ID, messages=[flex_message, link_message])
                )
    except Exception:
        traceback.print_exc()


def start_scheduler():
    if not GROUP_ID:
        return
    try:
        from apscheduler.schedulers.background import BackgroundScheduler
        from apscheduler.triggers.cron import CronTrigger
        try:
            from zoneinfo import ZoneInfo
            tz = ZoneInfo("Asia/Ho_Chi_Minh")
        except Exception:
            tz = None

        scheduler = BackgroundScheduler(timezone=tz)
        scheduler.add_job(
            send_scheduled_report,
            CronTrigger(hour=DAILY_HOUR, minute=DAILY_MINUTE, timezone=tz),
            id="daily_report",
            replace_existing=True,
        )
        scheduler.start()
    except Exception:
        traceback.print_exc()


start_scheduler()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
