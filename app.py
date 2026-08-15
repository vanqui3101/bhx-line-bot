"""
app.py - Webhook server LINE bot báo cáo doanh thu + ngành hàng (bản 4)

Luồng hoạt động (MỚI):
1. Người dùng gửi file Excel vào bot -> bot TỰ NHẬN DIỆN loại file:
   - File "doanh thu theo siêu thị" -> lưu vào SQLite, CHỈ trả lời xác nhận.
   - File "chi tiết ngành hàng" (Nấm/Bánh trung thu/C2) -> lưu vào SQLite,
     CHỈ trả lời xác nhận.
   => Bot KHÔNG còn tự động gửi thẻ báo cáo vào nhóm khi nhận file nữa.
      Chỉ gửi báo cáo khi có người gõ đúng lệnh (bên dưới).

2. Gõ lệnh, có thể gõ TRỰC TIẾP trong nhóm LINE (không cần chat riêng với bot):
   - "DT" hoặc "báo cáo doanh thu"          -> trả về thẻ Báo cáo doanh thu
   - "MỤC TIÊU KHUYẾN MÃI" hoặc "MTKM"      -> trả về thẻ Báo cáo ngành hàng
   - "id nhóm" (gõ trong nhóm)              -> trả về Group ID của nhóm đó

CẤU HÌNH (Environment Variables):
- LINE_CHANNEL_ACCESS_TOKEN  (bắt buộc)
- LINE_CHANNEL_SECRET        (bắt buộc)
- GROUP_ID                   (tuỳ chọn — nếu gõ lệnh trong chat riêng với bot,
                               bot sẽ gửi báo cáo vào nhóm này thay vì chat riêng)
"""

import os
import re
import uuid
import traceback
from datetime import datetime as _dt

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

from excel_reader import read_all_rows, read_category_rows, detect_file_type
from flex_builder import build_flex_message, build_category_flex_message
from excel_report import build_detail_excel
import storage

CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET", "")
GROUP_ID = os.environ.get("GROUP_ID", "").strip()
PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "").rstrip("/")

TMP_DIR = os.path.join(os.path.dirname(__file__), "tmp")
os.makedirs(TMP_DIR, exist_ok=True)

REPORTS_DIR = os.path.join(os.path.dirname(__file__), "reports")
os.makedirs(REPORTS_DIR, exist_ok=True)

app = Flask(__name__)
configuration = Configuration(access_token=CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)

# ---- Lệnh nhận diện (khớp cả khi gõ trong nhóm, không phân biệt hoa/thường) ----
DT_COMMAND_PATTERN = re.compile(
    r"^\s*(dt|b[aá]o\s*c[aá]o\s*doanh\s*thu)\s*$", re.IGNORECASE
)
MTKM_COMMAND_PATTERN = re.compile(
    r"^\s*(mtkm|m[uụ]c\s*ti[eê]u\s*khuy[eế]n\s*m[aã]i)\s*$", re.IGNORECASE
)
GROUP_ID_COMMAND_PATTERN = re.compile(r"^\s*id\s*nh[oó]m\s*$", re.IGNORECASE)


@app.route("/", methods=["GET"])
def health():
    return "LINE bot báo cáo doanh thu + ngành hàng (bản 4) đang chạy.", 200


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


def _now_vn_time_str():
    try:
        from zoneinfo import ZoneInfo
        now_vn = _dt.now(ZoneInfo("Asia/Ho_Chi_Minh"))
    except Exception:
        now_vn = _dt.now()
    return now_vn.strftime("%H:%M")


# ---------------------------------------------------------------------------
# Dựng nội dung báo cáo (dùng chung cho lệnh DT)
# ---------------------------------------------------------------------------

def build_revenue_report_messages(base_url):
    """Tạo flex_message báo cáo doanh thu mới nhất, hoặc None nếu chưa có dữ liệu."""
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

    return flex_message


def build_category_report_message():
    """Tạo flex_message báo cáo ngành hàng mới nhất, hoặc None nếu chưa có dữ liệu."""
    ngay, ten_st, payload, _gio = storage.get_latest_category_report()
    if ngay is None:
        return None

    bubble = build_category_flex_message(ngay, ten_st, payload)
    flex_message = FlexMessage(
        alt_text=f"Báo cáo ngành hàng {ngay}",
        contents=FlexContainer.from_dict(bubble),
    )
    return flex_message


# ---------------------------------------------------------------------------
# Nhận file Excel: chỉ LƯU DỮ LIỆU + xác nhận, KHÔNG tự động gửi báo cáo
# ---------------------------------------------------------------------------

@handler.add(MessageEvent, message=FileMessageContent)
def handle_file_message(event):
    message_id = event.message.id
    original_name = event.message.file_name or "input.xlsx"

    with ApiClient(configuration) as api_client:
        blob_api = MessagingApiBlob(api_client)
        messaging_api = MessagingApi(api_client)

        if not original_name.lower().endswith((".xlsx", ".xlsm")):
            reply_text(messaging_api, event.reply_token,
                       "Bot chỉ đọc được file Excel (.xlsx). Anh gửi lại file định dạng .xlsx giúp em nhé.")
            return

        tmp_path = os.path.join(TMP_DIR, f"input_{uuid.uuid4().hex}.xlsx")
        try:
            content = blob_api.get_message_content(message_id)
            with open(tmp_path, "wb") as f:
                f.write(content)

            file_type = detect_file_type(tmp_path)

            if file_type == "revenue":
                date_str, rows = read_all_rows(tmp_path)
                storage.save_records(rows)
                storage.save_snapshot_time(date_str, _now_vn_time_str())

                date_display = _dt.strptime(date_str, "%Y-%m-%d").strftime("%d/%m/%Y")
                total = sum((r["dt_offline"] or 0) + (r["dt_online"] or 0) for r in rows)
                total_str = f"{total:,.0f}".replace(",", ".")
                reply = (
                    f"✅ Đã lưu dữ liệu DOANH THU ngày {date_display} cho {len(rows)} dòng.\n"
                    f"Tổng doanh thu: {total_str} đ\n\n"
                    f"Gõ \"DT\" để xem báo cáo."
                )
                reply_text(messaging_api, event.reply_token, reply)

            elif file_type == "category":
                payload = read_category_rows(tmp_path)
                storage.save_category_report(
                    payload["ngay"], payload["ten_st"], payload, _now_vn_time_str()
                )
                date_display = _dt.strptime(payload["ngay"], "%Y-%m-%d").strftime("%d/%m/%Y")
                reply = (
                    f"✅ Đã lưu dữ liệu NGÀNH HÀNG ngày {date_display}.\n"
                    f"Nấm: {payload['nam']['doanh_thu']:,.0f} đ | "
                    f"Bánh trung thu: {payload['banh_trung_thu']['tong_sl']:.0f} cái | "
                    f"C2: {payload['c2']['tong_chai']:.0f} chai\n\n"
                    f"Gõ \"MỤC TIÊU KHUYẾN MÃI\" để xem báo cáo."
                ).replace(",", ".")
                reply_text(messaging_api, event.reply_token, reply)

            else:
                reply_text(
                    messaging_api, event.reply_token,
                    "Không nhận diện được loại file này. Anh kiểm tra lại đúng file "
                    "\"Doanh thu theo siêu thị\" hoặc file \"Doanh thu chi tiết\" (ngành hàng) nhé."
                )

        except Exception as e:
            traceback.print_exc()
            reply_text(messaging_api, event.reply_token, f"Có lỗi khi xử lý file: {e}")
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)


# ---------------------------------------------------------------------------
# Nhận lệnh text — hoạt động cả khi gõ TRỰC TIẾP trong nhóm LINE
# ---------------------------------------------------------------------------

@handler.add(MessageEvent, message=TextMessageContent)
def handle_text_message(event):
    text = (event.message.text or "").strip()

    with ApiClient(configuration) as api_client:
        messaging_api = MessagingApi(api_client)

        # Nguồn gửi lệnh: nếu gõ trong nhóm -> group_id của chính nhóm đó.
        # Nếu gõ chat riêng với bot -> ưu tiên gửi vào GROUP_ID đã cấu hình (nếu có),
        # để báo cáo vẫn lên đúng nhóm chung.
        source_type = event.source.type
        if source_type == "group":
            target_id = event.source.group_id
        else:
            target_id = GROUP_ID or getattr(event.source, "user_id", None)

        # Lệnh lấy Group ID
        if GROUP_ID_COMMAND_PATTERN.match(text):
            if source_type == "group":
                reply_text(messaging_api, event.reply_token, f"Group ID của nhóm này:\n{event.source.group_id}")
            else:
                reply_text(messaging_api, event.reply_token, "Lệnh này chỉ dùng được trong nhóm (group) nhé anh.")
            return

        # Lệnh DT — báo cáo doanh thu
        if DT_COMMAND_PATTERN.match(text):
            try:
                reply_text(messaging_api, event.reply_token, "⏳ Đang tổng hợp báo cáo doanh thu, đợi anh vài giây...")
            except Exception:
                traceback.print_exc()
            try:
                result = build_revenue_report_messages(_base_url())
                if result is None:
                    push_text(messaging_api, target_id, "Chưa có dữ liệu doanh thu nào được lưu. Anh gửi file Excel doanh thu trước nhé.")
                    return
                flex_message = result
                messaging_api.push_message(
                    PushMessageRequest(to=target_id, messages=[flex_message])
                )
            except Exception as e:
                traceback.print_exc()
                try:
                    push_text(messaging_api, target_id, f"Có lỗi khi tạo báo cáo doanh thu: {e}")
                except Exception:
                    traceback.print_exc()
            return

        # Lệnh MỤC TIÊU KHUYẾN MÃI — báo cáo ngành hàng
        if MTKM_COMMAND_PATTERN.match(text):
            try:
                reply_text(messaging_api, event.reply_token, "⏳ Đang tổng hợp báo cáo ngành hàng, đợi anh vài giây...")
            except Exception:
                traceback.print_exc()
            try:
                flex_message = build_category_report_message()
                if flex_message is None:
                    push_text(messaging_api, target_id, "Chưa có dữ liệu ngành hàng nào được lưu. Anh gửi file Excel chi tiết ngành hàng trước nhé.")
                    return
                messaging_api.push_message(
                    PushMessageRequest(to=target_id, messages=[flex_message])
                )
            except Exception as e:
                traceback.print_exc()
                try:
                    push_text(messaging_api, target_id, f"Có lỗi khi tạo báo cáo ngành hàng: {e}")
                except Exception:
                    traceback.print_exc()
            return

        # Không khớp lệnh nào -> bỏ qua, không phản hồi.


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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
