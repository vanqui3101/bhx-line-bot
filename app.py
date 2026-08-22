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
- GROUP_ID                   (không còn dùng để định tuyến tin nhắn — bot luôn
                               trả lời đúng nơi gõ lệnh, giữ biến này chỉ để
                               tương thích ngược nếu cần dùng lại sau này)
"""
import os
import re
import uuid
import traceback
import requests
import json
from datetime import datetime as _dt
from flask import Flask, request, abort, send_from_directory
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
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
from excel_reader import (
    read_all_rows, read_category_rows, read_stock_rows, attach_stock_percentage,
    read_thuong_period_rows, count_distinct_dates, detect_file_type,
    is_schedule_file, read_schedule_rows,
    is_ca_schedule_file, read_ca_schedule, phan_line_assign,
    TEN_NGAN_TO_MA_NV, NOI_DUNG_CA_MAC_DINH,
)
from flex_builder import build_flex_message, build_category_flex_message, build_thuong_flex_message
from excel_report import build_detail_excel
import storage

CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET", "")
GROUP_ID = os.environ.get("GROUP_ID", "").strip()
PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "").rstrip("/")

# Lịch hỗ trợ siêu thị khác mặc định (nạp sẵn nếu chưa có file nào được gửi lên).
# Anh có thể gửi file Excel mới (cột Ngày | Tên | Ca làm) bất cứ lúc nào để thay lịch này.
DEFAULT_SUPPORT_SCHEDULE = [
    {"ngay": "2026-08-20", "ten": "QUYÊN", "ca": "123"},
    {"ngay": "2026-08-21", "ten": "ÁNH", "ca": "456"},
    {"ngay": "2026-08-23", "ten": "THI", "ca": "123"},
    {"ngay": "2026-08-24", "ten": "MI", "ca": "456"},
    {"ngay": "2026-08-25", "ten": "LINH", "ca": "123"},
    {"ngay": "2026-08-26", "ten": "SANG", "ca": "456"},
    {"ngay": "2026-08-27", "ten": "ÁNH", "ca": "123"},
    {"ngay": "2026-08-28", "ten": "QUYÊN", "ca": "456"},
]
storage.ensure_default_schedule(DEFAULT_SUPPORT_SCHEDULE)

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
TD_COMMAND_PATTERN = re.compile(
    r"^\s*(td|th[uư][oở]ng)\s*$", re.IGNORECASE
)
GROUP_ID_COMMAND_PATTERN = re.compile(r"^\s*id\s*nh[oó]m\s*$", re.IGNORECASE)
DANG_KY_COMMAND_PATTERN = re.compile(r"^\s*dk\s+(.+?)\s*$", re.IGNORECASE)

# ---- [TẠM THỜI - TEST] Lệnh so sánh tag "@Tất cả" vs tag 1 người cụ thể ----
# Gõ trong nhóm: "TEST TAG ALL" hoặc "TEST TAG <Tên>" (VD: "TEST TAG Mi").
# Dùng để xác định nguyên nhân lỗi "not a member of the group" là do ID người
# bị lệch (chỉ tag riêng lỗi) hay do cả tài khoản OA (tag kiểu nào cũng lỗi).
# Xoá cả khối này (đánh dấu TẠM THỜI - TEST) sau khi đã xác định xong nguyên nhân.
TEST_TAG_ALL_PATTERN = re.compile(r"^\s*test\s*tag\s*all\s*$", re.IGNORECASE)
TEST_TAG_ME_PATTERN = re.compile(r"^\s*test\s*tag\s*me\s*$", re.IGNORECASE)
TEST_TAG_ONE_PATTERN = re.compile(r"^\s*test\s*tag\s+([a-zA-ZÀ-ỹ]+)\s*$", re.IGNORECASE)

TEN_NGAN_HOP_LE = {"mi", "quyên", "quyen", "sang", "thi", "ánh", "anh", "linh"}
TEN_NGAN_CHUAN_HOA = {
    "mi": "Mi", "quyên": "Quyên", "quyen": "Quyên", "sang": "Sang",
    "thi": "Thi", "ánh": "Ánh", "anh": "Ánh", "linh": "Linh",
}


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
    nh_ngay, nh_ten_st, nh_payload, _nh_gio = storage.get_latest_category_report()
    nganh_hang_breakdown = nh_payload["nganh_hang"]["items"] if nh_payload else None
    bubble = build_flex_message(
        latest_date, latest_records, prev_date, prev_records, gio_now, gio_prev,
        nganh_hang_breakdown=nganh_hang_breakdown, nganh_hang_ngay=nh_ngay,
    )
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
    _stock_ten_st, stock_payload, _stock_gio = storage.get_latest_stock_snapshot()
    if stock_payload:
        payload = attach_stock_percentage(payload, stock_payload)
    bubble = build_category_flex_message(ngay, ten_st, payload)
    flex_message = FlexMessage(
        alt_text=f"Báo cáo ngành hàng {ngay}",
        contents=FlexContainer.from_dict(bubble),
    )
    return flex_message


def build_thuong_report_message():
    """Tạo flex_message báo cáo thưởng (FRESH + FMCG) mới nhất, hoặc None nếu chưa có dữ liệu."""
    ten_st, payload, _gio = storage.get_latest_thuong_report()
    if ten_st is None:
        return None
    bubble = build_thuong_flex_message(ten_st, payload)
    flex_message = FlexMessage(
        alt_text="Báo cáo thưởng FRESH + FMCG",
        contents=FlexContainer.from_dict(bubble),
    )
    return flex_message


# ---------------------------------------------------------------------------
# NHẮC LỊCH HỖ TRỢ SIÊU THỊ KHÁC — tự động tag đúng người lúc 20h & 21h
# ---------------------------------------------------------------------------
def refresh_group_members(group_id):
    """[KHÔNG CÒN DÙNG] API lấy danh sách thành viên nhóm bị LINE chặn
    (ForbiddenException: 'Access to this API is not available for your
    account'). Giữ hàm rỗng để không phải sửa các chỗ gọi tới nó."""
    return


def _find_user_id_by_name(ten):
    """Tìm user_id đã ĐĂNG KÝ (không còn gọi API bị chặn) — mỗi bạn chỉ cần
    gõ 1 lần lệnh 'DK <Tên>' trong nhóm để bot ghi nhớ ID thật."""
    ten_norm = ten.strip()
    user_id = storage.get_user_id_da_dang_ky(ten_norm)
    if user_id:
        return user_id, ten_norm
    return None, None


def _push_mention_message(group_id, display_name, user_id, ngay_hien_thi, ca):
    """Gửi tin nhắn có TAG THẬT (@tên) vào nhóm — dùng "textV2" +
    "substitution" (đúng định dạng LINE yêu cầu cho tin nhắn GỬI ĐI có tag,
    khác với định dạng "mention.mentionees" chỉ dùng cho tin nhắn NHẬN VÀO)."""
    ca_text = f" ca {ca}" if ca else ""
    full_text = f"Nhắc {{u1}} ngày {ngay_hien_thi} có lịch đi hỗ trợ siêu thị khác{ca_text}."
    body = {
        "to": group_id,
        "messages": [
            {
                "type": "textV2",
                "text": full_text,
                "substitution": {
                    "u1": {"type": "mention", "mentionee": {"type": "user", "userId": user_id}}
                },
            }
        ],
    }
    resp = requests.post(
        "https://api.line.me/v2/bot/message/push",
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {CHANNEL_ACCESS_TOKEN}",
        },
        json=body,
        timeout=15,
    )
    if resp.status_code >= 300:
        print("Loi gui tin nhac lich ho tro:", resp.status_code, resp.text)


def send_support_reminder(gio_nhac):
    """Kiểm tra lịch hỗ trợ của NGÀY MAI (nhắc trước 1 ngày, tối hôm trước),
    nếu có người thì tag nhắc vào nhóm.
    gio_nhac: '20h' hoặc '21h' — chỉ để tránh gửi trùng trong cùng khung giờ."""
    if not GROUP_ID:
        return
    try:
        from zoneinfo import ZoneInfo
        now_vn = _dt.now(ZoneInfo("Asia/Ho_Chi_Minh"))
    except Exception:
        now_vn = _dt.now()
    from datetime import timedelta
    ngay_mai = now_vn + timedelta(days=1)
    ngay_str = ngay_mai.strftime("%Y-%m-%d")
    log_key = ngay_str  # log theo ngày được nhắc (ngày mai), tránh gửi trùng
    if storage.da_nhac_chua(log_key, gio_nhac):
        return
    row = storage.get_schedule_for_date(ngay_str)
    if not row:
        return
    ten, ca = row
    refresh_group_members(GROUP_ID)
    user_id, display_name = _find_user_id_by_name(ten)
    if not user_id:
        print(f"Khong tim thay '{ten}' trong danh ba nhom de tag.")
        return
    try:
        ngay_hien_thi = ngay_mai.strftime("%d/%m")
        _push_mention_message(GROUP_ID, display_name, user_id, ngay_hien_thi, ca)
        storage.danh_dau_da_nhac(log_key, gio_nhac)
    except Exception:
        traceback.print_exc()


# ---------------------------------------------------------------------------
# [TẠM THỜI - TEST] So sánh tag "@Tất cả" và tag 1 người cụ thể qua textV2,
# để xác định nguyên nhân thật của lỗi "not a member of the group":
#   - Nếu "TEST TAG ALL" gửi được nhưng "TEST TAG <Tên>" lỗi
#     -> do ID của người đó bị lệch/cũ, KHÔNG liên quan tài khoản OA.
#   - Nếu cả 2 đều lỗi giống nhau
#     -> nghiêng về khả năng tài khoản OA thật sự bị giới hạn tính năng mention.
# Xoá cả khối "TẠM THỜI - TEST" này (kể cả 2 regex pattern ở trên) sau khi
# đã xác định xong nguyên nhân.
# ---------------------------------------------------------------------------
def _test_push_tag(mode, ten_ngan=None, direct_user_id=None):
    if not GROUP_ID:
        print("[TEST-TAG-DEBUG] khong co GROUP_ID")
        return
    if mode == "all":
        text = "{u1} — test tag toàn bộ nhóm"
        substitution = {"u1": {"type": "mention", "mentionee": {"type": "all"}}}
    elif direct_user_id:
        # Dung thang user_id cua nguoi vua go lenh (chac chan dung 100%,
        # khong phu thuoc du lieu da hoc/dang ky truoc do).
        user_id = direct_user_id
        text = "{u1} — test tag chính người vừa gõ lệnh"
        substitution = {"u1": {"type": "mention", "mentionee": {"type": "user", "userId": user_id}}}
    else:
        user_id, ten_chuan = _find_user_id_by_name(ten_ngan)
        if not user_id:
            print(f"[TEST-TAG-DEBUG] chua co ID cua '{ten_ngan}' "
                  f"(chua duoc hoc tu bai phan line hoac chua go lenh DK)")
            return
        text = "{u1} — test tag 1 người"
        substitution = {"u1": {"type": "mention", "mentionee": {"type": "user", "userId": user_id}}}
    body = {
        "to": GROUP_ID,
        "messages": [{"type": "textV2", "text": text, "substitution": substitution}],
    }
    resp = requests.post(
        "https://api.line.me/v2/bot/message/push",
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {CHANNEL_ACCESS_TOKEN}",
        },
        json=body,
        timeout=15,
    )
    print(f"[TEST-TAG-DEBUG] mode={mode} ten={ten_ngan} direct_user_id={direct_user_id} "
          f"-> status={resp.status_code} body={resp.text}")


# ---------------------------------------------------------------------------
# NHẮC THEO BÀI PHÂN LINE HÀNG NGÀY (THU NGÂN / FRESH / FMCG)
# ---------------------------------------------------------------------------
def _noi_dung_thu_ngan_fresh(ca):
    cho = "sáng" if ca == "sang" else "chiều"
    return (
        "Team bám sát mục tiêu tập trung tư vấn khuyến mãi giúp em\n"
        f"Tận dụng từng lượt khách chợ {cho}\n"
        "Cảm ơn Anh/Chị"
    )


NOI_DUNG_FMCG_CHIEU = "Xử lí nhanh hàng kho trung tâm, chỉnh chu kệ, dọn kho"


def _is_phan_line_message(text):
    tu = text.upper()
    return "THU NGÂN" in tu and "FMCG" in tu


TEN_TU_DONG_HOC = {"Mi": "Mi", "Quyên": "Quyên", "Sang": "Sang", "Thi": "Thi", "Ánh": "Ánh", "Linh": "Linh"}


def _tu_dong_dang_ky_tu_mention(text, mentionees):
    """Tự động học ID thật của từng bạn ngay từ bài phân line anh gửi —
    không cần ai gõ lệnh DK gì cả. Cách làm: gộp cả phần TRONG tag lẫn đoạn
    text NGAY SAU tag (tới hết dòng) rồi tìm đúng 1 trong 6 tên ngắn đã
    biết (không phân biệt hoa/thường), vì có bạn tên nằm trong tag
    (VD "...198434-SANG"), có bạn tên nằm sau tag (VD "112006 Mi")."""
    for m in mentionees:
        user_id = getattr(m, "user_id", None) or getattr(m, "userId", None)
        idx = getattr(m, "index", None)
        length = getattr(m, "length", None)
        if user_id is None or idx is None or length is None:
            continue
        end = idx + length
        span_text = text[idx:end]
        newline_pos = text.find("\n", end)
        window_end = newline_pos if newline_pos != -1 else min(end + 40, len(text))
        after_text = text[end:window_end]
        combined = span_text + " " + after_text
        for token in re.split(r"[\s\(\),\-_]+", combined):
            token = token.strip()
            if not token:
                continue
            token_upper = token.upper()
            matched = None
            for ten_ngan, ten_chuan in TEN_TU_DONG_HOC.items():
                if token_upper == ten_ngan.upper():
                    matched = ten_chuan
                    break
            if matched:
                storage.dang_ky_thanh_vien(matched, user_id)
                print(f"[PHANLINE-DEBUG] Tu dong hoc: {matched} = {user_id}")
                break


def _parse_phan_line(text, mentionees):
    """Tách bài phân line thành dữ liệu theo ca (sáng/chiều) x nhóm
    (thu_ngan_fresh / fmcg). mentionees: list các object có .index, .length,
    .user_id (lấy từ event.message.mention.mentionees, tin nhắn LINE gốc)."""
    lines = text.split("\n")
    # tinh vi tri (offset ky tu) bat dau cua tung dong trong text goc
    offsets = []
    pos = 0
    for line in lines:
        offsets.append(pos)
        pos += len(line) + 1  # +1 cho ky tu xuong dong
    data = {
        "sang": {"thu_ngan_fresh_users": [], "fmcg_users": [], "fmcg_text_lines": []},
        "chieu": {"thu_ngan_fresh_users": [], "fmcg_users": [], "fmcg_text_lines": []},
    }
    current_ca = "sang"
    current_group = "thu_ngan_fresh"
    for i, line in enumerate(lines):
        line_upper = line.strip().upper()
        line_start = offsets[i]
        line_end = line_start + len(line)
        is_header = False
        if "THU NGÂN" in line_upper:
            is_header = True
            current_group = "thu_ngan_fresh"
            if "CHIỀU" in line_upper or "CHIEU" in line_upper:
                current_ca = "chieu"
            elif "SÁNG" in line_upper or "SANG" in line_upper:
                current_ca = "sang"
        elif line_upper.startswith("FRESH"):
            is_header = True
            current_group = "thu_ngan_fresh"
        elif line_upper.startswith("FMCG"):
            is_header = True
            current_group = "fmcg"
        # tim mention nam trong dong nay
        line_has_mention = False
        for m in mentionees:
            m_index = getattr(m, "index", None)
            m_uid = getattr(m, "user_id", None) or getattr(m, "userId", None)
            if m_index is None or m_uid is None:
                continue
            if line_start <= m_index < line_end:
                line_has_mention = True
                bucket = data[current_ca][f"{current_group}_users"]
                if m_uid not in bucket:
                    bucket.append(m_uid)
        if is_header or line_has_mention:
            continue
        stripped = line.strip()
        if not stripped:
            continue
        if current_group == "fmcg":
            data[current_ca]["fmcg_text_lines"].append(stripped)
    for ca in ("sang", "chieu"):
        data[ca]["fmcg_text"] = "\n".join(data[ca]["fmcg_text_lines"])
        del data[ca]["fmcg_text_lines"]
    return data


def _get_display_name(group_id, user_id):
    """[KHÔNG CÒN GỌI API] Tra tên ngắn từ danh sách đã ĐĂNG KÝ (DK <Tên>),
    không còn gọi API get_group_member_profile bị LINE chặn."""
    ten_ngan = storage.get_ten_ngan_tu_user_id(user_id)
    if ten_ngan:
        return ten_ngan
    print(f"[PHANLINE-DEBUG] user_id={user_id} chua duoc dang ky (DK <Ten>) -> khong co ten de tag")
    return None


def _push_mention_many(group_id, content, user_ids):
    """Gửi 1 tin nhắn có nội dung + tag thật NHIỀU người cùng lúc."""
    if not user_ids:
        return
    names = []
    for uid in user_ids:
        name = _get_display_name(group_id, uid)
        if name:
            names.append((uid, name))
    if not names:
        return
    text = content + "\n"
    substitution = {}
    for i, (uid, name) in enumerate(names):
        placeholder = f"u{i}"
        text += f"{{{placeholder}}}\n"
        substitution[placeholder] = {"type": "mention", "mentionee": {"type": "user", "userId": uid}}
    final_text = text.rstrip("\n")
    body = {
        "to": group_id,
        "messages": [{"type": "textV2", "text": final_text, "substitution": substitution}],
    }
    resp = requests.post(
        "https://api.line.me/v2/bot/message/push",
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {CHANNEL_ACCESS_TOKEN}",
        },
        json=body,
        timeout=15,
    )
    if resp.status_code >= 300:
        print("Loi gui tin nhac phan line:", resp.status_code, resp.text)


def _rotation_picker(candidates):
    """Xoay vòng công bằng: chọn 1 người trong danh sách, lưu trạng thái để
    lần sau (cùng danh sách này) sẽ chọn người tiếp theo."""
    key = "|".join(sorted(candidates))
    idx = storage.get_rotation_index(key)
    pick = candidates[idx % len(candidates)]
    storage.advance_rotation_index(key, len(candidates))
    return pick


def _build_phan_line_text_and_data(ngay_str, roster_for_date):
    """Từ danh sách người có mặt (sáng/chiều), áp quy tắc phân line, dựng
    sẵn văn bản (để lưu lịch sử) + substitution thật (để tag qua textV2) +
    dữ liệu user_id (để dùng cho lịch nhắc). Trả về (noi_dung_full_text,
    substitution_dict, phan_line_data)."""
    phan_line_data = {
        "sang": {"thu_ngan_fresh_users": [], "fmcg_users": [], "fmcg_text": ""},
        "chieu": {"thu_ngan_fresh_users": [], "fmcg_users": [], "fmcg_text": ""},
    }
    lines = []
    substitution = {}
    _placeholder_counter = [0]

    def them_dong(text_line):
        lines.append(text_line)

    def them_tag(ten_ngan):
        user_id, display_name = _find_user_id_by_name(ten_ngan)
        if user_id:
            placeholder = f"u{_placeholder_counter[0]}"
            _placeholder_counter[0] += 1
            substitution[placeholder] = {"type": "mention", "mentionee": {"type": "user", "userId": user_id}}
            them_dong("{" + placeholder + "}")
        else:
            them_dong(f"@{ten_ngan}")
        return user_id

    for ca_key, ca_label_sang_chieu, ten_hien in [("sang", "sáng", "( sáng )"), ("chieu", "chieu", "(chiều)")]:
        ten_ca_key = "sang" if ca_key == "sang" else "chieu"
        names = roster_for_date.get(ca_key, [])
        if not names:
            continue
        ket_qua = phan_line_assign(names, _rotation_picker)
        them_dong(f"THU NGÂN {ten_hien}" if ca_key == "sang" else f"THU NGÂN {ten_hien}")
        them_dong("")
        thu_ngan_ids = []
        for ten in ket_qua["THU_NGAN"]:
            uid = them_tag(ten)
            if uid:
                thu_ngan_ids.append(uid)
            for cv in NOI_DUNG_CA_MAC_DINH["THU_NGAN"][ten_ca_key]:
                them_dong(f"- {cv}")
            them_dong("")
        if ten_ca_key == "sang":
            them_dong("=> BC LỰA TRÁI CÂY TRƯỚC 8h")
            them_dong("")
        them_dong("FRESH")
        them_dong("")
        fresh_ids = []
        for ten in ket_qua["FRESH"]:
            uid = them_tag(ten)
            if uid:
                fresh_ids.append(uid)
            them_dong("( hỗ trợ thu ngân)")
            if ten_ca_key == "sang":
                them_dong("=> BC LỰA RAU CỦ TRƯỚC 8h")
            else:
                for cv in NOI_DUNG_CA_MAC_DINH["FRESH"]["chieu"]:
                    them_dong(f"- {cv}")
            them_dong("")
        them_dong("FMCG")
        them_dong("")
        fmcg_ids = []
        fmcg_text_lines = []
        for ten in ket_qua["FMCG"]:
            uid = them_tag(ten)
            if uid:
                fmcg_ids.append(uid)
            for cv in NOI_DUNG_CA_MAC_DINH["FMCG"][ten_ca_key]:
                them_dong(f"- {cv}")
                fmcg_text_lines.append(f"- {cv}")
            them_dong("")
        phan_line_data[ca_key]["thu_ngan_fresh_users"] = thu_ngan_ids + fresh_ids
        phan_line_data[ca_key]["fmcg_users"] = fmcg_ids
        phan_line_data[ca_key]["fmcg_text"] = "\n".join(fmcg_text_lines)
    them_dong("==> MỤC TIÊU CỤ THỂ TỪNG ANH/CHỊ báo cáo trước 22h")
    full_text = "\n".join(lines).rstrip()
    return full_text, substitution, phan_line_data


def auto_generate_and_post_phan_line(target_date_str, test_mode=False):
    """Tự tạo bài phân line cho 1 ngày (từ file lịch phân ca đã lưu), đăng
    lên nhóm với tag thật, và lưu vào kho dữ liệu phân line (dùng chung cho
    lịch nhắc THU NGÂN/FRESH/FMCG đã có)."""
    if not GROUP_ID:
        print("[CAVIEC3-DEBUG] khong co GROUP_ID")
        return
    roster = storage.get_ca_schedule(target_date_str)
    if not roster:
        print(f"[CAVIEC3-DEBUG] khong co du lieu lich ca cho ngay {target_date_str}")
        return
    full_text, substitution, phan_line_data = _build_phan_line_text_and_data(target_date_str, roster)
    print(f"[CAVIEC3-DEBUG] da tao bai phan line cho {target_date_str}, {len(substitution)} mentions")
    body = {
        "to": GROUP_ID,
        "messages": [{"type": "textV2", "text": full_text, "substitution": substitution}],
    }
    try:
        resp = requests.post(
            "https://api.line.me/v2/bot/message/push",
            headers={"Content-Type": "application/json", "Authorization": f"Bearer {CHANNEL_ACCESS_TOKEN}"},
            json=body, timeout=15,
        )
        if resp.status_code >= 300:
            print("[CAVIEC3-DEBUG] Loi gui bai phan line:", resp.status_code, resp.text)
    except Exception:
        traceback.print_exc()
    storage.save_phan_line(target_date_str, phan_line_data)
    print(f"[CAVIEC3-DEBUG] da luu phan_line cho {target_date_str}")


def send_phanline_reminder(ca, group, slot, noi_dung_co_dinh=None):
    """Gửi nhắc theo bài phân line hôm nay cho đúng ca/nhóm.
    Nếu noi_dung_co_dinh=None thì dùng nội dung anh viết trong bài (dành cho FMCG sáng)."""
    print(f"[PHANLINE-DEBUG] scheduler chay slot={slot} ca={ca} group={group}")
    if not GROUP_ID:
        print("[PHANLINE-DEBUG] khong co GROUP_ID -> bo qua")
        return
    try:
        from zoneinfo import ZoneInfo
        now_vn = _dt.now(ZoneInfo("Asia/Ho_Chi_Minh"))
    except Exception:
        now_vn = _dt.now()
    ngay_str = now_vn.strftime("%Y-%m-%d")
    if storage.da_nhac_phanline(ngay_str, slot):
        print(f"[PHANLINE-DEBUG] slot {slot} ngay {ngay_str} da gui roi -> bo qua")
        return
    data = storage.get_phan_line(ngay_str)
    if not data:
        print(f"[PHANLINE-DEBUG] khong co du lieu phan line cho ngay {ngay_str} -> bo qua")
        return
    ca_data = data.get(ca, {})
    user_ids = ca_data.get(f"{group}_users", [])
    print(f"[PHANLINE-DEBUG] user_ids cho ca={ca} group={group}: {user_ids}")
    if not user_ids:
        print("[PHANLINE-DEBUG] danh sach user_ids rong -> bo qua, khong gui gi")
        return
    if noi_dung_co_dinh is not None:
        noi_dung = noi_dung_co_dinh(ca) if callable(noi_dung_co_dinh) else noi_dung_co_dinh
    else:
        noi_dung = ca_data.get("fmcg_text") or "Xử lí công việc FMCG hôm nay giúp em."
    try:
        _push_mention_many(GROUP_ID, noi_dung, user_ids)
        storage.danh_dau_da_nhac_phanline(ngay_str, slot)
        print(f"[PHANLINE-DEBUG] da gui xong slot={slot}")
    except Exception:
        traceback.print_exc()


scheduler = BackgroundScheduler(timezone="Asia/Ho_Chi_Minh")
scheduler.add_job(lambda: send_support_reminder("20h"), CronTrigger(hour=20, minute=0))
scheduler.add_job(lambda: send_support_reminder("21h"), CronTrigger(hour=21, minute=0))

# Ca sáng: THU NGÂN + FRESH -> 8h, 11h
scheduler.add_job(lambda: send_phanline_reminder("sang", "thu_ngan_fresh", "sang_8h", _noi_dung_thu_ngan_fresh),
                   CronTrigger(hour=8, minute=0))
scheduler.add_job(lambda: send_phanline_reminder("sang", "thu_ngan_fresh", "sang_11h", _noi_dung_thu_ngan_fresh),
                   CronTrigger(hour=11, minute=0))
# Ca chiều: THU NGÂN + FRESH -> 15h, 17h, 19h
scheduler.add_job(lambda: send_phanline_reminder("chieu", "thu_ngan_fresh", "chieu_15h", _noi_dung_thu_ngan_fresh),
                   CronTrigger(hour=15, minute=0))
scheduler.add_job(lambda: send_phanline_reminder("chieu", "thu_ngan_fresh", "chieu_17h", _noi_dung_thu_ngan_fresh),
                   CronTrigger(hour=17, minute=0))
scheduler.add_job(lambda: send_phanline_reminder("chieu", "thu_ngan_fresh", "chieu_19h", _noi_dung_thu_ngan_fresh),
                   CronTrigger(hour=19, minute=0))
# FMCG sáng -> 10h (dùng đúng nội dung anh viết trong bài)
scheduler.add_job(lambda: send_phanline_reminder("sang", "fmcg", "fmcg_sang_10h", None),
                   CronTrigger(hour=10, minute=0))
# FMCG chiều -> 19h (nội dung cố định)
scheduler.add_job(lambda: send_phanline_reminder("chieu", "fmcg", "fmcg_chieu_19h", NOI_DUNG_FMCG_CHIEU),
                   CronTrigger(hour=19, minute=0))

# ---- LỊCH TEST KHẨN CẤP (tạm thời, để debug ngay hôm nay) ----
# Bắt đầu 15h40, lặp lại mỗi 35 phút — dùng đúng nội dung/dữ liệu ca chiều
# THU NGÂN+FRESH thật, chỉ khác là dùng slot-key riêng mỗi lần nên không bị
# chặn gửi trùng như bản chính thức. Có thể xoá khối này sau khi hết debug.
def _test_phanline_reminder():
    try:
        from zoneinfo import ZoneInfo
        now_vn = _dt.now(ZoneInfo("Asia/Ho_Chi_Minh"))
    except Exception:
        now_vn = _dt.now()
    slot_test = f"test_{now_vn.strftime('%H%M')}"
    print(f"[PHANLINE-DEBUG] === CHAY LICH TEST KHAN CAP, slot={slot_test} ===")
    send_phanline_reminder("chieu", "thu_ngan_fresh", slot_test, _noi_dung_thu_ngan_fresh)


# ---- LỊCH TEST CÔNG VIỆC 3 (tạm thời, hôm nay 21/08) ----
# 16h30: tự tạo + đăng bài phân line cho ngày MAI (22/08) từ file lịch phân ca.
# 16h40: gửi thử 1 lượt nhắc THU NGÂN+FRESH ca sáng của 22/08 để test tag thật.
# Sau khi chạy thật ổn định, đổi lại giờ cho đúng quy trình chính thức
# (ví dụ đăng bài tối hôm trước, tương tự lịch hỗ trợ siêu thị khác).
def _test_tao_phan_line_224():
    print("[CAVIEC3-DEBUG] === TEST: tao bai phan line cho 22/08 ===")
    auto_generate_and_post_phan_line("2026-08-22", test_mode=True)


def _test_nhac_phan_line_224():
    print("[CAVIEC3-DEBUG] === TEST: gui thu nhac cho 22/08 ===")
    try:
        from zoneinfo import ZoneInfo
        now_vn = _dt.now(ZoneInfo("Asia/Ho_Chi_Minh"))
    except Exception:
        now_vn = _dt.now()
    data = storage.get_phan_line("2026-08-22")
    if not data:
        print("[CAVIEC3-DEBUG] chua co du lieu phan_line cho 22/08 -> bo qua nhac thu")
        return
    user_ids = data.get("sang", {}).get("thu_ngan_fresh_users", [])
    if not user_ids:
        print("[CAVIEC3-DEBUG] khong co user de nhac thu")
        return
    _push_mention_many(GROUP_ID, _noi_dung_thu_ngan_fresh("sang"), user_ids)
    print("[CAVIEC3-DEBUG] === da gui thu nhac ===")


from apscheduler.triggers.date import DateTrigger
_gio_tao_bai = _dt(2026, 8, 22, 12, 23, 0)
_gio_nhac_thu = _dt(2026, 8, 21, 15, 30, 0)
try:
    from zoneinfo import ZoneInfo as _ZI2
    _tz_vn = _ZI2("Asia/Ho_Chi_Minh")
    _gio_tao_bai = _gio_tao_bai.replace(tzinfo=_tz_vn)
    _gio_nhac_thu = _gio_nhac_thu.replace(tzinfo=_tz_vn)
except Exception:
    pass
scheduler.add_job(_test_tao_phan_line_224, DateTrigger(run_date=_gio_tao_bai))
scheduler.add_job(_test_nhac_phan_line_224, DateTrigger(run_date=_gio_nhac_thu))

scheduler.start()


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
            if is_ca_schedule_file(tmp_path):
                ca_data = read_ca_schedule(tmp_path)
                for ngay, roster in ca_data.items():
                    storage.save_ca_schedule(ngay, roster)
                print(f"[CAVIEC3-DEBUG] Da luu LICH PHAN CA ({len(ca_data)} ngay), khong tra loi theo yeu cau anh Qui.")
                return
            if is_schedule_file(tmp_path):
                schedule_rows = read_schedule_rows(tmp_path)
                storage.save_support_schedule_rows(schedule_rows)
                so_dong = len(schedule_rows)
                reply = f"✅ Đã cập nhật LỊCH HỖ TRỢ SIÊU THỊ KHÁC ({so_dong} dòng). Bot sẽ tự nhắc đúng người vào 20h và 21h mỗi ngày."
                reply_text(messaging_api, event.reply_token, reply)
                return
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
                so_ngay = count_distinct_dates(tmp_path)
                if so_ngay >= 2:
                    # File trai nhieu ngay (vd 01/08 -> hien tai) -> bao cao THUONG
                    # + tu dong cap nhat luon MTKM bang du lieu cua NGAY GAN NHAT trong file
                    thuong_payload = read_thuong_period_rows(tmp_path)
                    storage.save_thuong_report(
                        thuong_payload["ten_st"], thuong_payload, _now_vn_time_str()
                    )
                    mtkm_payload = read_category_rows(tmp_path, filter_date=thuong_payload["ngay_ket_thuc"])
                    storage.save_category_report(
                        mtkm_payload["ngay"], mtkm_payload["ten_st"], mtkm_payload, _now_vn_time_str()
                    )
                    ngay_bd_disp = _dt.strptime(thuong_payload["ngay_bat_dau"], "%Y-%m-%d").strftime("%d/%m")
                    ngay_kt_disp = _dt.strptime(thuong_payload["ngay_ket_thuc"], "%Y-%m-%d").strftime("%d/%m")
                    reply = (
                        f"✅ Đã lưu dữ liệu THƯỞNG ({thuong_payload['so_ngay_da_qua']} ngày, "
                        f"{ngay_bd_disp} - {ngay_kt_disp}) và cập nhật MTKM cho ngày {ngay_kt_disp}.\n"
                        f"Tổng thưởng dự kiến: {thuong_payload['tong_thuong_du_kien']:,.0f} đ\n\n"
                        f"Gõ \"TD\"/\"THƯỞNG\" hoặc \"MTKM\" để xem báo cáo tương ứng."
                    ).replace(",", ".")
                    reply_text(messaging_api, event.reply_token, reply)
                else:
                    # File 1 ngay -> bao cao nganh hang (MTKM) nhu cu
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
            elif file_type == "stock":
                stock_payload = read_stock_rows(tmp_path)
                storage.save_stock_snapshot(
                    stock_payload["ten_st"], stock_payload, _now_vn_time_str()
                )
                so_sp = len(stock_payload["ton_kho_map"])
                reply = (
                    f"✅ Đã lưu dữ liệu TỒN KHO cho {so_sp} sản phẩm.\n\n"
                    f"Gõ \"MỤC TIÊU KHUYẾN MÃI\" để xem báo cáo có kèm % bán trên tồn."
                )
                reply_text(messaging_api, event.reply_token, reply)
            else:
                reply_text(
                    messaging_api, event.reply_token,
                    "Không nhận diện được loại file này. Anh kiểm tra lại đúng file "
                    "\"Doanh thu theo siêu thị\", \"Doanh thu chi tiết\" (ngành hàng), "
                    "hoặc \"BC tồn theo model\" (tồn kho) nhé."
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
        # Bot luôn trả lời đúng nơi người dùng gõ lệnh:
        # - Gõ trong nhóm -> trả lời trong chính nhóm đó.
        # - Gõ chat riêng với bot -> trả lời lại trong chat riêng đó.
        source_type = event.source.type
        if source_type == "group":
            target_id = event.source.group_id
        else:
            target_id = getattr(event.source, "user_id", None)

        # Bài PHÂN LINE hàng ngày (THU NGÂN / FRESH / FMCG) — tự nhận diện,
        # KHÔNG cần lệnh gì cả. Chỉ xử lý khi gõ trong nhóm (cần group_id để
        # tra tên hiển thị + tag lại sau này).
        if source_type == "group" and _is_phan_line_message(text):
            try:
                mention_obj = getattr(event.message, "mention", None)
                print(f"[PHANLINE-DEBUG] raw event.message = {event.message}")
                print(f"[PHANLINE-DEBUG] mention_obj = {mention_obj!r} (type={type(mention_obj)})")
                mentionees = getattr(mention_obj, "mentionees", None) if mention_obj else None
                if mentionees is None:
                    mentionees = []
                print(f"[PHANLINE-DEBUG] so mentionees tim thay: {len(mentionees)}")
                for i, m in enumerate(mentionees):
                    print(f"[PHANLINE-DEBUG] mentionee[{i}] = {m!r} "
                          f"(index={getattr(m,'index',None)}, "
                          f"user_id={getattr(m,'user_id',None)}, "
                          f"userId={getattr(m,'userId',None)}, "
                          f"type_field={getattr(m,'type',None)})")
                data = _parse_phan_line(text, mentionees)
                print(f"[PHANLINE-DEBUG] ket qua tach: {json.dumps(data, ensure_ascii=False)}")
                _tu_dong_dang_ky_tu_mention(text, mentionees)
                try:
                    from zoneinfo import ZoneInfo
                    now_vn = _dt.now(ZoneInfo("Asia/Ho_Chi_Minh"))
                except Exception:
                    now_vn = _dt.now()
                ngay_str = now_vn.strftime("%Y-%m-%d")
                storage.save_phan_line(ngay_str, data)
                reply_text(messaging_api, event.reply_token,
                           "Em nhận line này rồi để em nhắc mấy anh/chị bám sát mục tiêu ngày để hoàn tất tốt mục tiêu anh ạ")
            except Exception:
                traceback.print_exc()
            return

        # ---- [TẠM THỜI - TEST] So sánh tag "@Tất cả" và tag 1 người cụ thể ----
        # Gõ trong nhóm: "TEST TAG ALL" hoặc "TEST TAG <Tên>" (VD: "TEST TAG Mi").
        # Xem log ở Railway để đọc status code + nội dung LINE trả về.
        if source_type == "group" and TEST_TAG_ALL_PATTERN.match(text):
            _test_push_tag("all")
            return
        if source_type == "group" and TEST_TAG_ME_PATTERN.match(text):
            my_user_id = getattr(event.source, "user_id", None)
            _test_push_tag("user", direct_user_id=my_user_id)
            return
        if source_type == "group" and TEST_TAG_ONE_PATTERN.match(text):
            ten_test = TEST_TAG_ONE_PATTERN.match(text).group(1)
            _test_push_tag("user", ten_ngan=ten_test)
            return

        # Lệnh ĐĂNG KÝ (thay thế API bị LINE chặn) — mỗi bạn gõ "DK <Tên>" 1 lần
        dk_match = DANG_KY_COMMAND_PATTERN.match(text)
        if dk_match:
            ten_raw = dk_match.group(1).strip()
            ten_key = ten_raw.lower()
            if ten_key in TEN_NGAN_HOP_LE:
                ten_chuan = TEN_NGAN_CHUAN_HOA[ten_key]
                user_id = getattr(event.source, "user_id", None)
                if user_id:
                    storage.dang_ky_thanh_vien(ten_chuan, user_id)
                    reply_text(messaging_api, event.reply_token, f"Đã ghi nhận {ten_chuan}, cảm ơn bạn!")
                else:
                    reply_text(messaging_api, event.reply_token, "Không lấy được ID của bạn, thử lại giúp em nhé.")
            else:
                reply_text(messaging_api, event.reply_token,
                           "Tên chưa đúng — gõ đúng 1 trong: Mi, Quyên, Sang, Thi, Ánh, Linh. VD: DK Mi")
            return

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
                reply_text(messaging_api, event.reply_token, "Em gửi Anh và Team 8363 luôn ạ")
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
                reply_text(messaging_api, event.reply_token, "Em gửi Anh và Team 8363 luôn ạ")
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

        # Lệnh TD / THƯỞNG — báo cáo thưởng FRESH + FMCG
        if TD_COMMAND_PATTERN.match(text):
            try:
                reply_text(messaging_api, event.reply_token, "Em gửi Anh và Team 8363 luôn ạ")
            except Exception:
                traceback.print_exc()
            try:
                flex_message = build_thuong_report_message()
                if flex_message is None:
                    push_text(messaging_api, target_id, "Chưa có dữ liệu thưởng nào được lưu. Anh gửi file Excel doanh thu chi tiết (từ đầu tháng đến hiện tại) trước nhé.")
                    return
                messaging_api.push_message(
                    PushMessageRequest(to=target_id, messages=[flex_message])
                )
            except Exception as e:
                traceback.print_exc()
                try:
                    push_text(messaging_api, target_id, f"Có lỗi khi tạo báo cáo thưởng: {e}")
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
