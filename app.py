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
   - Đăng bài PHÂN LINE (nhóm hoặc chat riêng) -> thẻ MỤC TIÊU NGÀY TỪNG BẠN (01/11/2026)
   - Các bạn nhắn số bán cuối ngày (tự do)   -> bot ghi nhận, nhắc lại số đã hiểu
   - "CHỐT SỐ" (hoặc "chốt số 01/11")        -> thẻ chốt số cuối ngày: mục tiêu ngày cửa hàng + từng bạn
   - "BC <tên> [dd/mm]: <số>" / "XÓA BC <tên> [dd/mm]" -> anh nhập hộ / xoá báo cáo của 1 bạn
CẤU HÌNH (Environment Variables):
- LINE_CHANNEL_ACCESS_TOKEN  (bắt buộc)
- LINE_CHANNEL_SECRET        (bắt buộc)
- GROUP_ID                   (không dùng để định tuyến tin nhắn TRẢ LỜI theo
                               lệnh — bot luôn trả lời đúng nơi gõ lệnh. CHỈ
                               dùng cho các tin TỰ ĐỘNG gửi/nhắc (phân line,
                               lịch hỗ trợ...). Có thể để 1 ID hoặc NHIỀU ID
                               cách nhau bằng dấu phẩy để bot tự động gửi/
                               nhắc vào TẤT CẢ các nhóm đó cùng lúc, VD:
                               "Cxxxx1,Cxxxx2")
- ANTHROPIC_API_KEY          (bắt buộc cho tính năng TROLY trả lời tự do VÀ
                               tính năng đọc ảnh, lấy trong Anthropic Console
                               -> API Keys)

TÍNH NĂNG MỚI - ĐỌC ẢNH + PHÂN TÍCH SỐ LIỆU TỪ ẢNH:
- Anh gửi ảnh (chụp rõ, đủ thông tin số liệu) vào ĐÚNG đoạn chat (nhóm hoặc
  riêng) mà anh sẽ tag bot phân tích -> bot đọc và ghi nhớ TẠM (trong RAM,
  còn hiệu lực 2 tiếng) làm "ảnh gần nhất" của đúng đoạn chat đó.
- Sau đó tag bot ("Quí 227216 - BOT") + gõ có cụm "phân tích số liệu" (trong
  nhóm) -> bot đọc lại đúng ảnh vừa gửi bằng Claude Vision rồi phân tích.
- Nếu KHÔNG có ảnh gần nhất còn hiệu lực -> quay lại luồng cũ (phân tích dựa
  trên báo cáo "hủy mmkk" đã gửi bằng lệnh trước đó).
"""
import os
import re
import uuid
import traceback
import requests
import json
from datetime import datetime as _dt, timedelta as _td
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
from linebot.v3.webhooks import MessageEvent, FileMessageContent, TextMessageContent, ImageMessageContent, PostbackEvent
from excel_reader import (
    read_all_rows, read_category_rows, read_category_rows_by_date, read_stock_rows, attach_stock_percentage,
    read_thuong_period_rows, count_distinct_dates, detect_file_type,
    is_schedule_file, read_schedule_rows,
    is_ca_schedule_file, read_ca_schedule, phan_line_assign,
    TEN_NGAN_TO_MA_NV, NOI_DUNG_CA_MAC_DINH,
    read_fresh_rows,
)
from flex_builder import build_flex_message, build_category_flex_message, build_thuong_flex_message
from excel_report import build_detail_excel
import fresh_report
import dtdk_report
import storage
import ai_assistant
import thuong_freshfmcg
import mtkm_tracker
import muc_tieu_ngay
import muc_tieu_flex
import gdrive_reader
import menu_bot
CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET", "")
GROUP_ID = os.environ.get("GROUP_ID", "").strip()
# MỚI (16/09/2026): thư mục Google Drive cố định để bot tự kiểm tra file mới
# theo lịch (xem check_google_drive_for_new_files phía dưới). Để trống thì
# tính năng này tự tắt, không ảnh hưởng gì tới phần còn lại của bot.
GOOGLE_DRIVE_FOLDER_ID = os.environ.get("GOOGLE_DRIVE_FOLDER_ID", "").strip()
# Hỗ trợ NHIỀU nhóm cùng lúc cho các tính năng tự động gửi/nhắc (phân line,
# lịch hỗ trợ...): trên Railway, biến GROUP_ID có thể để 1 ID (như cũ) hoặc
# nhiều ID cách nhau bằng dấu phẩy, VD: "Cxxxx1,Cxxxx2". Mọi nơi tự động gửi
# tin (không phải trả lời theo lệnh) sẽ gửi tới TẤT CẢ nhóm trong danh sách.
GROUP_IDS = [g.strip() for g in GROUP_ID.split(",") if g.strip()]
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
# ---------------------------------------------------------------------------
# Cache TẠM (trong RAM, không lưu SQLite) - lưu ảnh gần nhất mỗi đoạn chat để
# lệnh "phân tích số liệu" đọc lại đúng ảnh anh vừa gửi. Mất khi bot khởi
# động lại (restart) - nếu mất thì anh gửi lại ảnh là được. Chỉ dùng được
# trong vòng ANH_HET_HAN_GIAY kể từ lúc gửi, tránh phân tích nhầm ảnh cũ.
# ---------------------------------------------------------------------------
_ANH_GAN_NHAT = {}  # target_id (group_id hoặc user_id) -> {"data": bytes, "luc": datetime}
ANH_HET_HAN_GIAY = 2 * 60 * 60  # 2 tiếng
def _luu_anh_gan_nhat(target_id, image_bytes):
    _ANH_GAN_NHAT[target_id] = {"data": image_bytes, "luc": _dt.now()}
def _lay_anh_gan_nhat(target_id):
    info = _ANH_GAN_NHAT.get(target_id)
    if not info:
        return None
    if (_dt.now() - info["luc"]).total_seconds() > ANH_HET_HAN_GIAY:
        return None
    return info["data"]
def _xoa_anh_gan_nhat(target_id):
    """Xoá ảnh đang lưu tạm của đoạn chat này — gọi khi có lệnh báo cáo mới
    (DT/MTKM) chạy, để lệnh "phân tích số liệu" sau đó không lỡ dùng nhầm
    ảnh cũ còn sót lại (ưu tiên đúng hành động gần nhất anh vừa làm)."""
    _ANH_GAN_NHAT.pop(target_id, None)
app = Flask(__name__)
configuration = Configuration(access_token=CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)
# (11/11/2026) GHI LOG KẾT QUẢ GỬI TIN LÊN LINE — để biết chắc LINE nhận hay từ chối
# (VD hết hạn mức tin tháng: HTTP 429). Mỗi lần trả lời/gửi in 1 dòng [LINE-REPLY]/[LINE-PUSH].
def _boc_log_gui_tin(ten_ham, nhan):
    goc = getattr(MessagingApi, ten_ham)
    def _goi(self, *args, **kwargs):
        try:
            kq = goc(self, *args, **kwargs)
            print(f"[{nhan}] OK")
            return kq
        except Exception as e:
            print(f"[{nhan}] LOI {getattr(e, 'status', '')} {str(getattr(e, 'body', '') or e)[:300]}")
            raise
    setattr(MessagingApi, ten_ham, _goi)
_boc_log_gui_tin("reply_message", "LINE-REPLY")
_boc_log_gui_tin("push_message", "LINE-PUSH")
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
DTDK_COMMAND_PATTERN = re.compile(
    r"^\s*dtdk\s*$", re.IGNORECASE
)
# TTFF — Tiến độ THƯỞNG Fresh/FMCG theo Base cố định (MỚI 12/09/2026, TÁCH
# RIÊNG khỏi lệnh TD/THƯỞNG cũ). Gõ "TTFF" hoặc "tiến độ thưởng".
TTFF_COMMAND_PATTERN = re.compile(
    r"^\s*(ttff|ti[eế]n\s*đ[ộo]\s*th[uư][oở]ng)\s*$", re.IGNORECASE
)
GROUP_ID_COMMAND_PATTERN = re.compile(r"^\s*id\s*nh[oó]m\s*$", re.IGNORECASE)
# (01/11/2026) "CHỐT SỐ" — thẻ chốt số cuối ngày (mục tiêu ngày cửa hàng + từng bạn). Dò trên bản không dấu.
CHOT_SO_PATTERN = re.compile(r"(?<![a-z])chot\s*so(?![a-z])")
DANG_KY_COMMAND_PATTERN = re.compile(r"^\s*dk\s+(.+?)\s*$", re.IGNORECASE)
# ---- Lệnh mới (HỦY TỒN + MMKK / PHÂN TÍCH / DOANH THU THỦY HẢI SẢN) ----
# Khác với DT/MTKM/TD: gõ TỰ DO miễn có đúng cụm từ khóa trong câu, VÀ phải
# TAG TÊN BOT trong câu thì bot mới trả lời (bot tên hiển thị "Quí 227216 - BOT").
#
# Tiếng Việt có nhiều cách gõ dấu cho cùng 1 từ (vd "hủy" và "huỷ" là 1 từ,
# chỉ khác chỗ đặt dấu) -> so khớp trên bản KHÔNG DẤU để không bị bỏ sót.
import unicodedata
def _bo_dau(text):
    """Bỏ dấu tiếng Việt (và hạ chữ thường) để so khớp từ khóa không phụ
    thuộc cách gõ dấu, vd "hủy" và "huỷ" đều thành "huy"."""
    text = text or ""
    nfkd = unicodedata.normalize("NFKD", text)
    khong_dau = "".join(c for c in nfkd if not unicodedata.combining(c))
    khong_dau = khong_dau.replace("đ", "d").replace("Đ", "D")
    return khong_dau.lower()
BOT_TAG_TEXT_KD = "qui 227216"
# 08/09/2026: nới lỏng để bắt được cả cách gõ tự nhiên có chữ "tồn" xen giữa
# (vd "hủy tồn mmkk", "hủy tồn + mmkk") — trước đây chỉ khớp đúng "hủy mmkk"
# 2 từ liền nhau, các cách gõ khác bị rớt xuống luồng AI trả lời tự do (ra
# bảng chữ thô thay vì thẻ card).
HUY_MMKK_TRIGGER = re.compile(r"huy.{0,12}mmkk|mmkk.{0,12}huy")
PHAN_TICH_TRIGGER = re.compile(r"phan\s*tich(\s*so\s*lieu)?")
SEAFOOD_TRIGGER = re.compile(r"doanh\s*thu\s*thuy\s*hai\s*san")
# MỚI (08/09/2026): lệnh "nhận xét" — so mục tiêu (bài phân line) với báo
# cáo thực tế (nhật ký tin nhắn) của 1 bạn nhân viên. Khớp trên text_kd
# (không dấu). Tên bạn thì PHẢI dò trên text GỐC (còn dấu, đúng hoa/thường)
# để tránh nhầm "Ánh" (tên) với "anh" (đại từ "anh" Quí hay dùng) — 2 từ
# này chỉ phân biệt được nhờ dấu, bỏ dấu đi thì giống hệt nhau.
NHAN_XET_TRIGGER = re.compile(r"nhan\s*xet")
TEN_NHAN_XET_RE = re.compile(r"\b(Mi|Quyên|Sang|Thi|Ánh|Linh|Son|Tiền)\b")
# MỚI (16/09/2026, theo yêu cầu anh Quí): "nhận xét <loại data> <ngày>" —
# NHẬN XÉT SỐ LIỆU trực tiếp, không cần gõ lệnh lấy báo cáo trước như "phân
# tích số liệu" cũ. Phân biệt với "nhận xét mục tiêu" (nhân viên) ở trên nhờ
# có kèm 1 trong các từ khoá loại data bên dưới — có từ khoá thì xử lý ở
# đây, KHÔNG có thì rớt xuống nhánh "nhận xét mục tiêu" cũ như bình thường.
NHAN_XET_MMKK_RE = re.compile(r"mmkk|huy\s*ton")
NHAN_XET_DOANH_THU_RE = re.compile(r"doanh\s*thu")
NHAN_XET_KM_RE = re.compile(r"khuyen\s*mai|nganh\s*hang|mtkm")
def _co_tag_bot(text):
    return BOT_TAG_TEXT_KD in _bo_dau(text)
_BOT_USER_ID_CACHE = {"id": None}
def _lay_bot_user_id():
    """Lấy user_id THẬT của chính bot (TROLY) trên LINE, để nhận diện khi anh
    dùng tính năng @ CHỌN TÊN BOT thật trong nhóm (tag thật), thay vì phải gõ
    tay cụm "Quí 227216". Gọi 1 lần rồi cache lại (user_id của bot không đổi)."""
    if _BOT_USER_ID_CACHE["id"]:
        return _BOT_USER_ID_CACHE["id"]
    try:
        resp = requests.get(
            "https://api.line.me/v2/bot/info",
            headers={"Authorization": f"Bearer {CHANNEL_ACCESS_TOKEN}"},
            timeout=5,
        )
        uid = (resp.json() or {}).get("userId")
        if uid:
            _BOT_USER_ID_CACHE["id"] = uid
        return uid
    except Exception:
        traceback.print_exc()
        return None
def _da_tag_bot_thuc_su(event, text):
    """"Tag bot" = ĐÚNG 1 trong 2 cách sau (MỚI 13/09/2026 — trước đây CHỈ
    nhận cách 1, anh Quí phản hồi không muốn gõ tay tên/mã của anh, chỉ muốn
    @ chọn tên bot thật cho gọn):
    1. Gõ tay cụm "Quí 227216" (cách cũ, vẫn giữ, không phá lệnh cũ nếu anh
       đã quen gõ kiểu này).
    2. Dùng tính năng @ chọn tên bot (TROLY) THẬT trong LINE — LINE trả về
       user_id của bot trong event.message.mention.mentionees, so khớp với
       user_id thật của bot lấy qua API."""
    if _co_tag_bot(text):
        return True
    mention_obj = getattr(event.message, "mention", None)
    mentionees = getattr(mention_obj, "mentionees", None) if mention_obj else None
    if not mentionees:
        return False
    bot_uid = _lay_bot_user_id()
    if not bot_uid:
        return False
    for m in mentionees:
        uid = getattr(m, "user_id", None) or getattr(m, "userId", None)
        if uid == bot_uid:
            return True
    return False
# ---- TROLY TRẢ LỜI TỰ DO (MỚI) ----
# Dùng để cắt bỏ đoạn tag tên bot ra khỏi câu hỏi trước khi gửi cho AI, cho
# câu hỏi sạch sẽ hơn (không bắt buộc, chỉ để câu hỏi gọn hơn khi đưa vào AI).
BOT_NAME_STRIP_PATTERN = re.compile(r"qu[ií]\s*227216\s*-?\s*bot", re.IGNORECASE)
# ---- [TẠM THỜI - TEST] Lệnh so sánh tag "@Tất cả" vs tag 1 người cụ thể ----
# Gõ trong nhóm: "TEST TAG ALL" hoặc "TEST TAG <Tên>" (VD: "TEST TAG Mi").
# Dùng để xác định nguyên nhân lỗi "not a member of the group" là do ID người
# bị lệch (chỉ tag riêng lỗi) hay do cả tài khoản OA (tag kiểu nào cũng lỗi).
# Xoá cả khối này (đánh dấu TẠM THỜI - TEST) sau khi đã xác định xong nguyên nhân.
TEST_TAG_ALL_PATTERN = re.compile(r"^\s*test\s*tag\s*all\s*$", re.IGNORECASE)
TEST_TAG_ME_PATTERN = re.compile(r"^\s*test\s*tag\s*me\s*$", re.IGNORECASE)
TEST_TAG_ONE_PATTERN = re.compile(r"^\s*test\s*tag\s+([a-zA-ZÀ-ỹ]+)\s*$", re.IGNORECASE)
TEN_NGAN_HOP_LE = {"mi", "quyên", "quyen", "sang", "thi", "ánh", "anh", "linh", "tiền", "tien", "son"}
TEN_NGAN_CHUAN_HOA = {
    "mi": "Mi", "quyên": "Quyên", "quyen": "Quyên", "sang": "Sang",
    "thi": "Thi", "ánh": "Ánh", "anh": "Ánh", "linh": "Linh",
    "tiền": "Tiền", "tien": "Tiền", "son": "Son",
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
def _now_vn_datetime():
    try:
        from zoneinfo import ZoneInfo
        return _dt.now(ZoneInfo("Asia/Ho_Chi_Minh"))
    except Exception:
        return _dt.now()
def _ngay_hien_tai_str():
    return _now_vn_datetime().strftime("%Y-%m-%d")
def _ngay_hom_qua_str():
    return (_now_vn_datetime() - _td(days=1)).strftime("%Y-%m-%d")
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
    ngay, ten_st, payload, gio = storage.get_latest_category_report()
    if ngay is None:
        return None
    # (05/10/2026) Thẻ MTKM mới: tính lũy kế + mục tiêu ngày từ dữ liệu đã lưu
    try:
        kq = mtkm_tracker.tinh_mtkm(ngay)
    except Exception:
        traceback.print_exc()
        kq = None
    bubble = build_category_flex_message(ngay, ten_st, payload, kq=kq, gio=gio)
    flex_message = FlexMessage(
        alt_text=f"Báo cáo ngành hàng {ngay}",
        contents=FlexContainer.from_dict(bubble),
    )
    return flex_message
def _luu_mtkm_tat_ca_ngay(tmp_path):
    """(05/10/2026) Lưu số liệu MTKM cho TỪNG NGÀY có trong file POS 76 (file
    1 ngày hay nhiều ngày đều được). Nạp lại ngày đã có -> ghi đè, không cộng
    trùng. Trả về (danh sách ngày đã lưu, payload của ngày mới nhất)."""
    theo_ngay = read_category_rows_by_date(tmp_path)
    gio = _now_vn_time_str()
    for ngay in sorted(theo_ngay):
        p = theo_ngay[ngay]
        storage.save_category_report(ngay, p["ten_st"], p, gio)
    ds = sorted(theo_ngay)
    return ds, (theo_ngay[ds[-1]] if ds else None)


def _tin_xac_nhan_mtkm(ds_ngay):
    """Tin nhắn xác nhận ngắn sau khi nạp file: ngày đã lưu + tóm tắt MTKM."""
    if not ds_ngay:
        return "Không tìm thấy dữ liệu ngày nào trong file."
    if len(ds_ngay) == 1:
        nhan = _dt.strptime(ds_ngay[0], "%Y-%m-%d").strftime("%d/%m/%Y")
    else:
        nhan = (f"{len(ds_ngay)} ngày ({_dt.strptime(ds_ngay[0], '%Y-%m-%d').strftime('%d/%m')} - "
                f"{_dt.strptime(ds_ngay[-1], '%Y-%m-%d').strftime('%d/%m')})")
    try:
        tom_tat = mtkm_tracker.tom_tat_text(mtkm_tracker.tinh_mtkm(ds_ngay[-1]))
    except Exception:
        traceback.print_exc()
        tom_tat = ""
    return f"Đã cập nhật MTKM {nhan}." + (f"\n{tom_tat}" if tom_tat else "")


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
    if not GROUP_IDS:
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
    user_id, display_name = _find_user_id_by_name(ten)
    if not user_id:
        print(f"Khong tim thay '{ten}' trong danh ba nhom de tag.")
        return
    try:
        ngay_hien_thi = ngay_mai.strftime("%d/%m")
        for gid in GROUP_IDS:
            _push_mention_message(gid, display_name, user_id, ngay_hien_thi, ca)
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
    if not GROUP_IDS:
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
    for gid in GROUP_IDS:
        body = {
            "to": gid,
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
              f"nhom={gid} -> status={resp.status_code} body={resp.text}")
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
    tu = _bo_dau_upper(text)
    return "THU NGAN" in tu and "FMCG" in tu
def _bo_dau_upper(s):
    """Bỏ dấu tiếng Việt + viết hoa (chỉ để NHẬN DIỆN chữ, không dùng để cắt text)."""
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D").upper()
def _mention_index_py(text, idx):
    """LINE có thể tính vị trí tag theo UTF-16 (1 emoji = 2 ký tự) trong khi
    Python tính 1 — bài có emoji trước tag sẽ bị lệch dòng. Ưu tiên vị trí
    nào rơi đúng vào ký tự '@'."""
    if idx is None:
        return None
    if 0 <= idx < len(text) and text[idx] == "@":
        return idx
    dem = 0
    for i, ch in enumerate(text):
        if dem == idx:
            if ch == "@":
                return i
            break
        if dem > idx:
            break
        dem += 2 if ord(ch) > 0xFFFF else 1
    return idx
_DAU_DONG_NOI_DUNG = ("-", "(", "=", ">", "+", "•", "*-")
def _tieu_de_phan_line(doan_text):
    """Nhận diện 1 dòng (hoặc phần chữ TRƯỚC tag đầu tiên của dòng) có phải TIÊU ĐỀ.
    Trả về (vai_tro, ca) — vai_tro: thu_ngan / fresh / fmcg / None (chỉ đổi ca) —
    hoặc None nếu là dòng nội dung. Chỉ tính là tiêu đề khi dòng BẮT ĐẦU bằng
    THU NGÂN / FRESH / FMCG / CA SÁNG / CA CHIỀU (bỏ qua emoji, dấu * # trang trí).
    Dòng bắt đầu bằng - ( = > + là nội dung công việc (VD "( hỗ trợ thu ngân)")."""
    s = doan_text.strip()
    if not s or s.startswith(_DAU_DONG_NOI_DUNG):
        return None
    s = _bo_dau_upper(s)
    s = re.sub(r"^[^A-Z0-9(\-=>+]+", "", s)
    if not s or s.startswith(("-", "(", "=", ">", "+")):
        return None
    ca = None
    if re.search(r"\bCHIEU\b", s):
        ca = "chieu"
    elif re.search(r"\bSANG\b", s):
        ca = "sang"
    if re.match(r"THU\s*NGAN\b", s):
        return ("thu_ngan", ca)
    if s.startswith("FRESH"):
        return ("fresh", ca)
    if s.startswith("FMCG"):
        return ("fmcg", ca)
    if re.match(r"CA\s+(SANG|CHIEU)\b", s):
        return (None, ca)
    return None
TEN_TU_DONG_HOC = {"Mi": "Mi", "Quyên": "Quyên", "Sang": "Sang", "Thi": "Thi", "Ánh": "Ánh", "Linh": "Linh",
                   "Tiền": "Tiền", "Son": "Son"}
def _tu_dong_dang_ky_tu_mention(text, mentionees):
    """Tự động học ID thật của từng bạn ngay từ bài phân line anh gửi —
    không cần ai gõ lệnh DK gì cả. Cách làm: gộp cả phần TRONG tag lẫn đoạn
    text NGAY SAU tag (tới hết dòng) rồi tìm đúng 1 trong 6 tên ngắn đã
    biết (không phân biệt hoa/thường), vì có bạn tên nằm trong tag
    (VD "...198434-SANG"), có bạn tên nằm sau tag (VD "112006 Mi")."""
    for m in mentionees:
        user_id = getattr(m, "user_id", None) or getattr(m, "userId", None)
        idx = _mention_index_py(text, getattr(m, "index", None))
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
    .user_id (lấy từ event.message.mention.mentionees, tin nhắn LINE gốc).

    MỚI (08/09/2026): còn tách thêm MỤC TIÊU riêng từng bạn — các dòng ngay
    SAU dòng có tag tên bạn đó (tới khi gặp dòng tag khác/dòng tiêu đề mới/
    dòng trống) được gom lại thành text mục tiêu của đúng bạn đó, lưu vào
    data[ca]["muc_tieu"][user_id] — dùng cho tính năng "nhận xét mục tiêu"
    (so với báo cáo thực tế bạn đó tự nhắn cuối ca/cuối ngày)."""
    lines = text.split("\n")
    # tinh vi tri (offset ky tu) bat dau cua tung dong trong text goc
    offsets = []
    pos = 0
    for line in lines:
        offsets.append(pos)
        pos += len(line) + 1  # +1 cho ky tu xuong dong
    data = {
        "sang": {"thu_ngan_fresh_users": [], "thu_ngan_users": [], "fresh_users": [],
                 "fmcg_users": [], "fmcg_text_lines": [], "muc_tieu": {},
                 "thu_ngan_names": [], "fresh_names": [], "fmcg_names": []},
        "chieu": {"thu_ngan_fresh_users": [], "thu_ngan_users": [], "fresh_users": [],
                  "fmcg_users": [], "fmcg_text_lines": [], "muc_tieu": {},
                  "thu_ngan_names": [], "fresh_names": [], "fmcg_names": []},
    }

    def _them_ten(ten):
        """(01/11/2026) Ghi tên bạn vào đúng vai trò + ca hiện tại (dùng chia mục tiêu từng bạn)."""
        vt = "fmcg" if current_group == "fmcg" else current_role
        if ten and vt:
            ds = data[current_ca][f"{vt}_names"]
            if ten not in ds:
                ds.append(ten)
    current_ca = "sang"
    current_group = "thu_ngan_fresh"
    # (01/11/2026) Tách riêng vai trò THU NGÂN và FRESH để chia mục tiêu MTKM cho từng bạn
    current_role = "thu_ngan"
    current_target_users = []  # user_id của (các) bạn vừa được tag ở dòng gần nhất
    # vị trí tag (đã đổi về index Python) + user_id
    tag_list = []
    for m in mentionees:
        m_uid = getattr(m, "user_id", None) or getattr(m, "userId", None)
        m_index = _mention_index_py(text, getattr(m, "index", None))
        if m_index is None or m_uid is None:
            continue
        tag_list.append((m_index, m_uid))
    tag_list.sort()
    for i, line in enumerate(lines):
        line_start = offsets[i]
        line_end = line_start + len(line)
        tag_trong_dong = [(ix, uid) for ix, uid in tag_list if line_start <= ix < line_end]
        # (01/11/2026) Chỉ xét TIÊU ĐỀ ở phần chữ TRƯỚC tag đầu tiên của dòng, và dòng phải
        # BẮT ĐẦU bằng THU NGÂN/FRESH/FMCG — tránh dòng "( hỗ trợ thu ngân)" hay
        # "@Sang hỗ trợ thu ngân" bị hiểu nhầm là tiêu đề THU NGÂN (lỗi Fresh mất hết)
        doan_truoc_tag = line[: tag_trong_dong[0][0] - line_start] if tag_trong_dong else line
        tieu_de = _tieu_de_phan_line(doan_truoc_tag)
        is_header = tieu_de is not None
        if tieu_de:
            vai_tro, ca_moi = tieu_de
            if ca_moi:
                current_ca = ca_moi
            if vai_tro == "fmcg":
                current_group = "fmcg"
                current_role = None
            elif vai_tro in ("thu_ngan", "fresh"):
                current_group = "thu_ngan_fresh"
                current_role = vai_tro
        line_has_mention = bool(tag_trong_dong)
        line_users = []
        for _ix, m_uid in tag_trong_dong:
            line_users.append(m_uid)
            try:
                _them_ten(storage.get_ten_ngan_tu_user_id(m_uid))
            except Exception:
                traceback.print_exc()
            bucket = data[current_ca][f"{current_group}_users"]
            if m_uid not in bucket:
                bucket.append(m_uid)
            if current_group == "thu_ngan_fresh" and current_role:
                role_bucket = data[current_ca][f"{current_role}_users"]
                if m_uid not in role_bucket:
                    role_bucket.append(m_uid)
        if is_header and line_has_mention:
            current_target_users = line_users
            continue
        if is_header:
            current_target_users = []
            continue
        if line_has_mention:
            current_target_users = line_users
            continue
        stripped = line.strip()
        if not stripped:
            current_target_users = []
            continue
        # (01/11/2026) Bài phân line kiểu gọn: tên gõ thường, không tag (VD "* mi", "* quyên")
        ten_dong = muc_tieu_ngay.nhan_nhieu_ten(stripped)
        if ten_dong:
            for ten in ten_dong:
                _them_ten(ten)
            current_target_users = []
            continue
        if current_group == "fmcg":
            data[current_ca]["fmcg_text_lines"].append(stripped)
        if current_target_users:
            for uid in current_target_users:
                muc_tieu_bucket = data[current_ca]["muc_tieu"]
                existing = muc_tieu_bucket.get(uid, "")
                muc_tieu_bucket[uid] = f"{existing}\n{stripped}".strip() if existing else stripped
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
        phan_line_data[ca_key]["thu_ngan_users"] = thu_ngan_ids
        phan_line_data[ca_key]["fresh_users"] = fresh_ids
        phan_line_data[ca_key]["fmcg_users"] = fmcg_ids
        phan_line_data[ca_key]["fmcg_text"] = "\n".join(fmcg_text_lines)
    them_dong("==> MỤC TIÊU CỤ THỂ TỪNG ANH/CHỊ báo cáo trước 22h")
    full_text = "\n".join(lines).rstrip()
    return full_text, substitution, phan_line_data
def auto_generate_and_post_phan_line(target_date_str, test_mode=False):
    """Tự tạo bài phân line cho 1 ngày (từ file lịch phân ca đã lưu), đăng
    lên nhóm với tag thật, và lưu vào kho dữ liệu phân line (dùng chung cho
    lịch nhắc THU NGÂN/FRESH/FMCG đã có)."""
    if not GROUP_IDS:
        print("[CAVIEC3-DEBUG] khong co GROUP_ID")
        return
    roster = storage.get_ca_schedule(target_date_str)
    if not roster:
        print(f"[CAVIEC3-DEBUG] khong co du lieu lich ca cho ngay {target_date_str}")
        return
    full_text, substitution, phan_line_data = _build_phan_line_text_and_data(target_date_str, roster)
    print(f"[CAVIEC3-DEBUG] da tao bai phan line cho {target_date_str}, {len(substitution)} mentions")
    try:
        for gid in GROUP_IDS:
            body = {
                "to": gid,
                "messages": [{"type": "textV2", "text": full_text, "substitution": substitution}],
            }
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
    if not GROUP_IDS:
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
        for gid in GROUP_IDS:
            _push_mention_many(gid, noi_dung, user_ids)
        storage.danh_dau_da_nhac_phanline(ngay_str, slot)
        print(f"[PHANLINE-DEBUG] da gui xong slot={slot}")
    except Exception:
        traceback.print_exc()
# QUAN TRỌNG — LỖI GIỜ UTC (đã bị lặp lại 2 lần, 07/09 và 08/09/2026):
# Chỉ truyền timezone cho BackgroundScheduler(timezone=...) KHÔNG đủ — mỗi
# CronTrigger(hour=..., minute=...) được TẠO SẴN rồi mới add_job() như bên
# dưới sẽ TỰ LẤY GIỜ HỆ MÁY CHỦ (UTC trên Railway) làm giờ tính, không tự
# động ăn theo timezone đã set cho scheduler (đây là hành vi đã được chính
# nhà phát triển thư viện APScheduler xác nhận — xem issue #346 trên GitHub
# repo agronholm/apscheduler: cơ chế "ăn theo giờ scheduler" chỉ áp dụng khi
# khai lịch qua dạng chuỗi rút gọn add_job(func, 'cron', hour=8), KHÔNG áp
# dụng khi truyền thẳng 1 object CronTrigger đã tạo sẵn như cách file này
# đang viết). Kết quả thực tế đã kiểm chứng qua log Railway: job hẹn 8h sáng
# VN chạy lúc 8h UTC = 15h VN, job hẹn 11h chạy lúc 18h VN — lệch đúng 7
# tiếng mỗi lần.
# SỬA TẬN GỐC: gắn thẳng timezone=_TZ_VN_SCHEDULER vào TỪNG CronTrigger bên
# dưới (không chỉ set 1 chỗ chung ở BackgroundScheduler nữa), để không còn
# phụ thuộc vào cơ chế "ăn theo mặc định" nữa.
try:
    from zoneinfo import ZoneInfo as _ZoneInfoLichChay
    _TZ_VN_SCHEDULER = _ZoneInfoLichChay("Asia/Ho_Chi_Minh")
except Exception:
    from datetime import timezone as _TzFixed, timedelta as _TdFixed
    _TZ_VN_SCHEDULER = _TzFixed(_TdFixed(hours=7))
scheduler = BackgroundScheduler(timezone=_TZ_VN_SCHEDULER)
# (31/10/2026) Anh Quí yêu cầu: TỪ THÁNG 11 HỦY TOÀN BỘ NHẮC TỰ ĐỘNG
# (nhắc lịch hỗ trợ 20h/21h + nhắc phân line 8h/10h/11h/15h/17h/19h).
# Bài phân line vẫn được lưu bình thường (dùng cho "nhận xét mục tiêu").
# Muốn bật lại: đổi BAT_NHAC_TU_DONG = True.
BAT_NHAC_TU_DONG = False
if BAT_NHAC_TU_DONG:
    scheduler.add_job(lambda: send_support_reminder("20h"), CronTrigger(hour=20, minute=0, timezone=_TZ_VN_SCHEDULER))
    scheduler.add_job(lambda: send_support_reminder("21h"), CronTrigger(hour=21, minute=0, timezone=_TZ_VN_SCHEDULER))
    # Ca sáng: THU NGÂN + FRESH -> 8h, 11h
    scheduler.add_job(lambda: send_phanline_reminder("sang", "thu_ngan_fresh", "sang_8h", _noi_dung_thu_ngan_fresh),
                       CronTrigger(hour=8, minute=0, timezone=_TZ_VN_SCHEDULER))
    scheduler.add_job(lambda: send_phanline_reminder("sang", "thu_ngan_fresh", "sang_11h", _noi_dung_thu_ngan_fresh),
                       CronTrigger(hour=11, minute=0, timezone=_TZ_VN_SCHEDULER))
    # Ca chiều: THU NGÂN + FRESH -> 15h, 17h, 19h
    scheduler.add_job(lambda: send_phanline_reminder("chieu", "thu_ngan_fresh", "chieu_15h", _noi_dung_thu_ngan_fresh),
                       CronTrigger(hour=15, minute=0, timezone=_TZ_VN_SCHEDULER))
    scheduler.add_job(lambda: send_phanline_reminder("chieu", "thu_ngan_fresh", "chieu_17h", _noi_dung_thu_ngan_fresh),
                       CronTrigger(hour=17, minute=0, timezone=_TZ_VN_SCHEDULER))
    scheduler.add_job(lambda: send_phanline_reminder("chieu", "thu_ngan_fresh", "chieu_19h", _noi_dung_thu_ngan_fresh),
                       CronTrigger(hour=19, minute=0, timezone=_TZ_VN_SCHEDULER))
    # FMCG sáng -> 10h (dùng đúng nội dung anh viết trong bài)
    scheduler.add_job(lambda: send_phanline_reminder("sang", "fmcg", "fmcg_sang_10h", None),
                       CronTrigger(hour=10, minute=0, timezone=_TZ_VN_SCHEDULER))
    # FMCG chiều -> 19h (nội dung cố định)
    scheduler.add_job(lambda: send_phanline_reminder("chieu", "fmcg", "fmcg_chieu_19h", NOI_DUNG_FMCG_CHIEU),
                       CronTrigger(hour=19, minute=0, timezone=_TZ_VN_SCHEDULER))
# TỰ ĐỘNG KIỂM TRA GOOGLE DRIVE (MỚI 16/09/2026) — mỗi 4 tiếng 1 lần
# (2h/6h/10h/14h/18h/22h giờ VN), tự tắt nếu chưa cấu hình GOOGLE_DRIVE_FOLDER_ID.
scheduler.add_job(lambda: check_google_drive_for_new_files(),
                   CronTrigger(hour="2,6,10,14,18,22", minute=5, timezone=_TZ_VN_SCHEDULER))
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
    for gid in GROUP_IDS:
        _push_mention_many(gid, _noi_dung_thu_ngan_fresh("sang"), user_ids)
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
# (31/10/2026) Bỏ 2 lịch test 21-22/08 (đã quá hạn, không còn dùng)
scheduler.start()
# XÁC NHẬN GIỜ SCHEDULER LÚC KHỞI ĐỘNG (để đối chiếu ngay trên log Railway,
# không cần đợi tới đúng khung giờ mới biết đúng/sai) — in giờ hệ thống
# (thường là UTC trên Railway) VÀ giờ VN mà từng CronTrigger đang thực sự
# dùng để tính lịch, phải LỆCH NHAU đúng 7 tiếng nếu máy chủ chạy UTC.
try:
    _gio_he_thong_luc_khoi_dong = _dt.now()
    _gio_vn_luc_khoi_dong = _dt.now(_TZ_VN_SCHEDULER)
    print(f"[SCHEDULER-CHECK] gio he thong (server) = {_gio_he_thong_luc_khoi_dong.strftime('%Y-%m-%d %H:%M:%S')} "
          f"| gio VN cac CronTrigger dang dung = {_gio_vn_luc_khoi_dong.strftime('%Y-%m-%d %H:%M:%S %Z')}")
except Exception:
    traceback.print_exc()
# ---------------------------------------------------------------------------
# Nhận file Excel: chỉ LƯU DỮ LIỆU + xác nhận, KHÔNG tự động gửi báo cáo
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# TỰ ĐỘNG NẠP FILE TỪ GOOGLE DRIVE (MỚI 16/09/2026)
# Dùng lại ĐÚNG logic đọc/lưu như khi anh gửi file tay vào bot
# (xem handle_file_message bên dưới), chỉ khác nguồn file (Drive thay vì
# LINE) và cách báo kết quả (nhắn chủ động vào nhóm thay vì reply tin nhắn).
# Viết thành hàm riêng (không sửa handle_file_message) để không ảnh hưởng
# tới tính năng gửi file tay đang chạy ổn định.
# ---------------------------------------------------------------------------
def _luu_file_ln_fresh(tmp_path, file_name):
    """(09/11/2026) Đọc + lưu file LN Fresh (POS 3933). Trả về câu xác nhận."""
    p = menu_bot.doc_ln_fresh(tmp_path, file_name or "")
    moi = menu_bot.luu_ln_fresh(p, _now_vn_time_str())
    gio = p.get("gio_xuat")
    gio_txt = f" (xuất {gio[11:]} ngày {gio[8:10]}/{gio[5:7]})" if gio else ""
    if not moi:
        return f"Bot đã có file LN Fresh mới hơn, bỏ qua file này{gio_txt}."
    t = p["tong"]
    return (f"✅ Đã lưu LN FRESH{gio_txt}: LN {menu_bot._tr(t['ln'], True)}tr, chênh so TB 3 tháng "
            f"{menu_bot._tr(t['chenh'], True)}tr.\nGõ \"LN FRESH\" để xem bảng.")
def _process_drive_file(tmp_path, file_name):
    """Đọc + lưu 1 file Excel đã tải về từ Drive. Trả về (thanh_cong: bool, noi_dung_thong_bao: str)."""
    try:
        if menu_bot.la_file_ln_fresh(tmp_path):  # (09/11/2026) file LN Fresh POS 3933
            return True, "[Tự động từ Drive] " + _luu_file_ln_fresh(tmp_path, file_name)
        if is_ca_schedule_file(tmp_path):
            ca_data = read_ca_schedule(tmp_path)
            for ngay, roster in ca_data.items():
                storage.save_ca_schedule(ngay, roster)
            return True, f"✅ [Tự động từ Drive] Đã lưu LỊCH PHÂN CA từ file \"{file_name}\" ({len(ca_data)} ngày)."
        if is_schedule_file(tmp_path):
            schedule_rows = read_schedule_rows(tmp_path)
            storage.save_support_schedule_rows(schedule_rows)
            return True, (f"✅ [Tự động từ Drive] Đã cập nhật LỊCH HỖ TRỢ SIÊU THỊ KHÁC từ file "
                          f"\"{file_name}\" ({len(schedule_rows)} dòng).")
        file_type = detect_file_type(tmp_path)
        if file_type == "revenue":
            date_str, rows = read_all_rows(tmp_path)
            storage.save_records(rows)
            storage.save_snapshot_time(date_str, _now_vn_time_str())
            date_display = _dt.strptime(date_str, "%Y-%m-%d").strftime("%d/%m/%Y")
            total = sum((r["dt_offline"] or 0) + (r["dt_online"] or 0) for r in rows)
            total_str = f"{total:,.0f}".replace(",", ".")
            return True, (f"✅ [Tự động từ Drive] Đã lưu DOANH THU ngày {date_display} ({len(rows)} dòng).\n"
                          f"Tổng doanh thu: {total_str} đ")
        elif file_type == "category":
            try:
                so_ngay_ttff = thuong_freshfmcg.extract_and_save(tmp_path)
                print(f"[DRIVE-DEBUG] da luu {so_ngay_ttff} ngay DT Fresh/FMCG")
            except Exception:
                traceback.print_exc()
            so_ngay = count_distinct_dates(tmp_path)
            if so_ngay >= 2:
                thuong_payload = read_thuong_period_rows(tmp_path)
                storage.save_thuong_report(thuong_payload["ten_st"], thuong_payload, _now_vn_time_str())
                ds_ngay, _p = _luu_mtkm_tat_ca_ngay(tmp_path)
                ngay_bd_disp = _dt.strptime(thuong_payload["ngay_bat_dau"], "%Y-%m-%d").strftime("%d/%m")
                ngay_kt_disp = _dt.strptime(thuong_payload["ngay_ket_thuc"], "%Y-%m-%d").strftime("%d/%m")
                return True, (f"✅ [Tự động từ Drive] Đã lưu THƯỞNG ({thuong_payload['so_ngay_da_qua']} ngày, "
                              f"{ngay_bd_disp} - {ngay_kt_disp}).\n" + _tin_xac_nhan_mtkm(ds_ngay))
            else:
                ds_ngay, _p = _luu_mtkm_tat_ca_ngay(tmp_path)
                return True, "✅ [Tự động từ Drive] " + _tin_xac_nhan_mtkm(ds_ngay)
        elif file_type == "stock":
            stock_payload = read_stock_rows(tmp_path)
            storage.save_stock_snapshot(
                stock_payload["ten_st"], stock_payload, _now_vn_time_str()
            )
            so_sp = len(stock_payload["ton_kho_map"])
            return True, f"✅ [Tự động từ Drive] Đã lưu dữ liệu TỒN KHO cho {so_sp} sản phẩm."
        elif file_type == "fresh":
            fresh_rows = read_fresh_rows(tmp_path)
            storage.save_fresh_records(fresh_rows)
            ngay_set = sorted({r["ngay"] for r in fresh_rows})
            ngay_bd_disp = _dt.strptime(ngay_set[0], "%Y-%m-%d").strftime("%d/%m/%Y")
            ngay_kt_disp = _dt.strptime(ngay_set[-1], "%Y-%m-%d").strftime("%d/%m/%Y")
            return True, (f"✅ [Tự động từ Drive] Đã lưu HỦY TỒN + MMKK ({len(ngay_set)} ngày, "
                          f"{ngay_bd_disp} - {ngay_kt_disp}), {len(fresh_rows)} dòng, từ file \"{file_name}\".")
        else:
            return False, f"⚠️ [Drive] Không nhận diện được loại file \"{file_name}\", bot bỏ qua file này."
    except Exception as e:
        traceback.print_exc()
        return False, f"⚠️ [Drive] Lỗi khi xử lý file \"{file_name}\": {e}"
def check_google_drive_for_new_files():
    """Chạy theo lịch (xem scheduler.add_job phía dưới) hoặc theo lệnh tay
    "kiểm tra Drive": kiểm tra thư mục Drive GOOGLE_DRIVE_FOLDER_ID, tải file
    MỚI (chưa xử lý lần nào, hoặc bị sửa lại sau lần xử lý trước) về nạp vào
    bot. KHÔNG tự nhắn gì vào nhóm/chat (chỉ ghi log Railway) — nếu gõ lệnh
    tay thì nơi gọi hàm này (handle_text_message) tự lo việc trả lời.
    Trả về số file đã nạp thành công trong lần chạy này."""
    if not GOOGLE_DRIVE_FOLDER_ID:
        return 0
    try:
        files = gdrive_reader.list_files_in_folder(GOOGLE_DRIVE_FOLDER_ID)
    except Exception:
        print("[DRIVE-DEBUG] Loi khi lay danh sach file tu Drive:")
        traceback.print_exc()
        return 0
    print(f"[DRIVE-DEBUG] tim thay {len(files)} file .xlsx/.xlsm trong thu muc Drive")
    so_file_da_nap = 0
    for f in files:
        file_id = f["id"]
        file_name = f["name"]
        modified_time = f["modifiedTime"]
        if storage.is_drive_file_processed(file_id, modified_time):
            continue
        print(f"[DRIVE-DEBUG] file moi: {file_name} ({file_id})")
        tmp_path = os.path.join(TMP_DIR, f"drive_{uuid.uuid4().hex}.xlsx")
        try:
            gdrive_reader.download_file(file_id, tmp_path)
            ok, message = _process_drive_file(tmp_path, file_name)
            storage.mark_drive_file_processed(file_id, file_name, modified_time)
            print(f"[DRIVE-DEBUG] {message}")
            if ok:
                so_file_da_nap += 1
        except Exception:
            print(f"[DRIVE-DEBUG] Loi khi tai/xu ly file {file_name}:")
            traceback.print_exc()
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
    return so_file_da_nap
DRIVE_CHECK_COMMAND_PATTERN = re.compile(r"^\s*kiem\s*tra\s*drive\s*$", re.IGNORECASE)
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
            # (09/11/2026) File LN Fresh (POS 3933 "Báo cáo hiệu quả bán Fresh siêu thị") cho nút LN FRESH
            if menu_bot.la_file_ln_fresh(tmp_path):
                reply_text(messaging_api, event.reply_token, _luu_file_ln_fresh(tmp_path, original_name))
                return
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
                # MỚI 12/09/2026: song song trích DT Fresh/FMCG theo ngày cho
                # lệnh "TTFF" (tiến độ thưởng, tách riêng TD cũ) — không ảnh
                # hưởng luồng MTKM/TD hiện có, lỗi ở đây không làm hỏng phần dưới.
                try:
                    so_ngay_ttff = thuong_freshfmcg.extract_and_save(tmp_path)
                    print(f"[TTFF-DEBUG] da luu {so_ngay_ttff} ngay DT Fresh/FMCG")
                except Exception:
                    traceback.print_exc()
                so_ngay = count_distinct_dates(tmp_path)
                if so_ngay >= 2:
                    # File trai nhieu ngay (vd 01/08 -> hien tai) -> bao cao THUONG
                    # + tu dong cap nhat luon MTKM bang du lieu cua NGAY GAN NHAT trong file
                    thuong_payload = read_thuong_period_rows(tmp_path)
                    storage.save_thuong_report(
                        thuong_payload["ten_st"], thuong_payload, _now_vn_time_str()
                    )
                    ds_ngay, _p = _luu_mtkm_tat_ca_ngay(tmp_path)
                    ngay_bd_disp = _dt.strptime(thuong_payload["ngay_bat_dau"], "%Y-%m-%d").strftime("%d/%m")
                    ngay_kt_disp = _dt.strptime(thuong_payload["ngay_ket_thuc"], "%Y-%m-%d").strftime("%d/%m")
                    reply = (
                        f"✅ Đã lưu dữ liệu THƯỞNG ({thuong_payload['so_ngay_da_qua']} ngày, "
                        f"{ngay_bd_disp} - {ngay_kt_disp}).\n"
                        f"Tổng thưởng dự kiến: {thuong_payload['tong_thuong_du_kien']:,.0f} đ".replace(",", ".")
                        + "\n" + _tin_xac_nhan_mtkm(ds_ngay)
                        + "\n\nGõ \"TD\"/\"THƯỞNG\" hoặc \"MTKM\" để xem báo cáo tương ứng."
                    )
                    reply_text(messaging_api, event.reply_token, reply)
                else:
                    # File 1 ngày -> lưu ngày đó (ghi đè nếu nạp lại), bot tự cộng dồn
                    ds_ngay, _p = _luu_mtkm_tat_ca_ngay(tmp_path)
                    reply = (
                        "✅ " + _tin_xac_nhan_mtkm(ds_ngay)
                        + "\n\nGõ \"MTKM\" để xem thẻ báo cáo."
                    )
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
            elif file_type == "fresh":
                fresh_rows = read_fresh_rows(tmp_path)
                storage.save_fresh_records(fresh_rows)
                ngay_set = sorted({r["ngay"] for r in fresh_rows})
                ngay_bd_disp = _dt.strptime(ngay_set[0], "%Y-%m-%d").strftime("%d/%m/%Y")
                ngay_kt_disp = _dt.strptime(ngay_set[-1], "%Y-%m-%d").strftime("%d/%m/%Y")
                reply = (
                    f"✅ Đã lưu dữ liệu HỦY TỒN + MMKK ({len(ngay_set)} ngày, "
                    f"{ngay_bd_disp} - {ngay_kt_disp}), {len(fresh_rows)} dòng.\n\n"
                    f"Tag bot + gõ \"hủy mmkk hôm qua\", \"doanh thu thủy hải sản\", "
                    f"hoặc \"phân tích số liệu\" để xem báo cáo."
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
# Nhận ẢNH — chỉ LƯU TẠM (RAM) làm "ảnh gần nhất" của đúng đoạn chat, để
# lệnh "phân tích số liệu" (gõ sau đó, có tag bot, trong nhóm) đọc lại đúng
# ảnh này bằng Claude Vision. KHÔNG phân tích ngay lúc nhận ảnh.
# ---------------------------------------------------------------------------
@handler.add(MessageEvent, message=ImageMessageContent)
def handle_image_message(event):
    message_id = event.message.id
    source_type = event.source.type
    if source_type == "group":
        target_id = event.source.group_id
    else:
        target_id = getattr(event.source, "user_id", None)
    with ApiClient(configuration) as api_client:
        blob_api = MessagingApiBlob(api_client)
        messaging_api = MessagingApi(api_client)
        try:
            content = blob_api.get_message_content(message_id)
            _luu_anh_gan_nhat(target_id, content)
            # Theo yêu cầu anh Quí: nhận ảnh xong KHÔNG trả lời gì cả (im
            # lặng lưu tạm), tránh làm phiền nhóm chat. Chỉ khi lưu bị lỗi
            # (ảnh tải về thất bại) mới báo, để anh biết mà gửi lại.
        except Exception:
            traceback.print_exc()
            try:
                reply_text(messaging_api, event.reply_token, "Em nhận ảnh bị lỗi, anh gửi lại giúp em.")
            except Exception:
                traceback.print_exc()
# ---------------------------------------------------------------------------
# Nhận lệnh text — hoạt động cả khi gõ TRỰC TIẾP trong nhóm LINE
# ---------------------------------------------------------------------------
@handler.add(MessageEvent, message=TextMessageContent)
def handle_text_message(event):
    text = (event.message.text or "").strip()
    text_kd = _bo_dau(text)  # bản không dấu, dùng cho các lệnh mới (mục "Lệnh mới" phía dưới)
    da_tag_bot = _da_tag_bot_thuc_su(event, text)
    print(f"[FRESH-DEBUG] text raw = {text!r}")
    print(f"[FRESH-DEBUG] text_kd (khong dau) = {text_kd!r}")
    print(f"[FRESH-DEBUG] da_tag_bot = {da_tag_bot} | huy_mmkk_match = {bool(HUY_MMKK_TRIGGER.search(text_kd))} "
          f"| phan_tich_match = {bool(PHAN_TICH_TRIGGER.search(text_kd))} | seafood_match = {bool(SEAFOOD_TRIGGER.search(text_kd))}")
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
        # NHẬT KÝ TIN NHẮN NHÂN VIÊN (MỚI 08/09/2026) — lưu lại MỌI tin nhắn
        # trong nhóm của bạn nào đã ĐĂNG KÝ tên (qua "DK <Tên>" hoặc tự học
        # qua tag bài phân line), làm "báo cáo thực tế" cho tính năng "nhận
        # xét mục tiêu". Không cần cú pháp gì đặc biệt, không trả lời gì cả
        # (im lặng lưu), không ảnh hưởng các lệnh xử lý bên dưới.
        if source_type == "group":
            _uid_nguoi_gui = getattr(event.source, "user_id", None)
            if _uid_nguoi_gui and text:
                _ten_ngan_nguoi_gui = storage.get_ten_ngan_tu_user_id(_uid_nguoi_gui)
                if _ten_ngan_nguoi_gui:
                    try:
                        storage.save_nhat_ky_nhan_vien(
                            _ngay_hien_tai_str(), _uid_nguoi_gui, _now_vn_time_str(), text
                        )
                    except Exception:
                        traceback.print_exc()
        # Bài PHÂN LINE hàng ngày (THU NGÂN / FRESH / FMCG) — tự nhận diện,
        # KHÔNG cần lệnh gì cả. Chỉ xử lý khi gõ trong nhóm (cần group_id để
        # tra tên hiển thị + tag lại sau này).
        if _is_phan_line_message(text):
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
                _tu_dong_dang_ky_tu_mention(text, mentionees)
                data = _parse_phan_line(text, mentionees)
                print(f"[PHANLINE-DEBUG] ket qua tach: {json.dumps(data, ensure_ascii=False)}")
                try:
                    from zoneinfo import ZoneInfo
                    now_vn = _dt.now(ZoneInfo("Asia/Ho_Chi_Minh"))
                except Exception:
                    now_vn = _dt.now()
                # (01/11/2026) Bài phân line áp cho ngày nào:
                #  - trong bài có ghi ngày (VD "02/11") -> đúng ngày đó
                #  - đăng từ 21h trở đi (sau giờ đóng cửa, thường là sau CHỐT SỐ) -> ngày MAI
                #  - còn lại -> hôm nay
                m_ngay_pl = re.search(r"(?<![\d/])(\d{1,2})\s*/\s*(\d{1,2})(?![\d/])", text)
                ngay_pl = None
                if m_ngay_pl:
                    try:
                        ngay_pl = _dt(now_vn.year, int(m_ngay_pl.group(2)), int(m_ngay_pl.group(1)))
                    except ValueError:
                        ngay_pl = None
                if ngay_pl is None:
                    ngay_pl = now_vn + _td(days=1) if now_vn.hour >= 21 else now_vn
                ngay_str = ngay_pl.strftime("%Y-%m-%d")
                storage.save_phan_line(ngay_str, data)
                # (02/11/2026) Xoá bài phân line của các ngày đã qua (giữ hôm qua để còn chốt số / sửa số)
                try:
                    storage.xoa_phan_line_truoc(_ngay_hom_qua_str())
                except Exception:
                    traceback.print_exc()
                # (01/11/2026) Trả ngay THẺ MỤC TIÊU NGÀY TỪNG BẠN (thẻ riêng, không dính MTKM)
                bubble = None
                try:
                    ch = muc_tieu_ngay.tinh_cua_hang(ngay_str)
                    if ch is not None:
                        tb = muc_tieu_ngay.chia_tung_ban(ch, data)
                        bubble = muc_tieu_flex.build_the_muc_tieu(ngay_str, ch, tb)
                except Exception:
                    traceback.print_exc()
                if bubble:
                    send_card_with_ack(messaging_api, event.reply_token, target_id, FlexMessage(
                        alt_text=f"Mục tiêu ngày từng bạn {ngay_pl.strftime('%d/%m')}",
                        contents=FlexContainer.from_dict(bubble)))
                else:
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
                    tra_loi = f"Đã ghi nhận {ten_chuan}, cảm ơn bạn!"
                    # (02/11/2026) Có báo cáo gửi lúc chưa đăng ký -> ghi luôn, không phải nhắn lại
                    try:
                        cho_items, cho_gio = storage.lay_va_xoa_bao_cao_cho_dk(user_id, _ngay_hien_tai_str())
                        if cho_items:
                            storage.save_bao_cao_ban_hang(_ngay_hien_tai_str(), ten_chuan, cho_items, cho_gio)
                            tra_loi += ("\nEm ghi luôn số bạn báo lúc " + (cho_gio or "nãy") + ":\n"
                                        + muc_tieu_ngay.tom_tat_bao_cao(ten_chuan, cho_items))
                    except Exception:
                        traceback.print_exc()
                    reply_text(messaging_api, event.reply_token, tra_loi)
                else:
                    reply_text(messaging_api, event.reply_token, "Không lấy được ID của bạn, thử lại giúp em nhé.")
            else:
                reply_text(messaging_api, event.reply_token,
                           "Tên chưa đúng — gõ đúng 1 trong: Mi, Quyên, Sang, Thi, Ánh, Linh, Tiền, Son. VD: DK Mi")
            return
        # Lệnh lấy Group ID
        if GROUP_ID_COMMAND_PATTERN.match(text):
            if source_type == "group":
                reply_text(messaging_api, event.reply_token, f"Group ID của nhóm này:\n{event.source.group_id}")
            else:
                reply_text(messaging_api, event.reply_token, "Lệnh này chỉ dùng được trong nhóm (group) nhé anh.")
            return
        # (01/11/2026) Lệnh CHỐT SỐ — thẻ chốt số cuối ngày: mục tiêu ngày cửa hàng so với POS,
        # MT ngày mai, từng bạn (mục tiêu / báo cáo / % đạt), ai báo cao hơn POS, 1 dòng nhận xét.
        # (03/11/2026) Lệnh "THU NGÂN" / "FRESH": tổng kết 3 ngày đã CHỐT SỐ gần nhất của từng bạn
        # (ai cao nhất, ai thấp nhất so với mục tiêu được giao mỗi ngày)
        # (04/11/2026) Thêm ngày -> xem chi tiết từng bạn của 1 ngày: "THU NGÂN 03/11", "FRESH hôm qua"
        _m_xh = re.fullmatch(r"\s*(thu\s*ngan|fresh)\s*(?:(hom\s*qua|hom\s*nay)|(\d{1,2})\s*[/-]\s*(\d{1,2}))?\s*",
                             text_kd)
        if _m_xh and (_m_xh.group(2) or _m_xh.group(3)):
            vai_tro_xh = "thu_ngan" if _m_xh.group(1).startswith("thu") else "fresh"
            ten_xh = "THU NGÂN" if vai_tro_xh == "thu_ngan" else "FRESH"
            try:
                if _m_xh.group(2):
                    ngay_xh = _ngay_hom_qua_str() if "qua" in _m_xh.group(2) else _ngay_hien_tai_str()
                else:
                    ngay_xh = f"{_now_vn_datetime().year}-{int(_m_xh.group(4)):02d}-{int(_m_xh.group(3)):02d}"
                dm_xh = f"{ngay_xh[8:10]}/{ngay_xh[5:7]}"
                kq_ng = muc_tieu_ngay.tinh_tung_ban_ngay(vai_tro_xh, ngay_xh)
                if not kq_ng or not kq_ng["so_nguoi"]:
                    notify_text(messaging_api, event.reply_token, target_id,
                                f"Ngày {dm_xh} chưa có mục tiêu {ten_xh} từng bạn anh ạ (bot không có bài phân line và chưa "
                                f"CHỐT SỐ ngày đó). Anh đăng lại bài phân line có ghi ngày {dm_xh} rồi gõ "
                                f"CHỐT SỐ {dm_xh} là xem được.")
                    return
                send_card_with_ack(messaging_api, event.reply_token, target_id, FlexMessage(
                    alt_text=f"{ten_xh} {dm_xh} – mục tiêu & báo cáo từng bạn",
                    contents=FlexContainer.from_dict(muc_tieu_flex.build_the_tung_ban_ngay(kq_ng))))
            except Exception as e:
                traceback.print_exc()
                try:
                    notify_text(messaging_api, event.reply_token, target_id, f"Có lỗi khi xem {ten_xh}: {e}")
                except Exception:
                    traceback.print_exc()
            return
        if _m_xh:
            try:
                vai_tro_xh = "thu_ngan" if _m_xh.group(1).startswith("thu") else "fresh"
                kq_xh = muc_tieu_ngay.tinh_xep_hang(vai_tro_xh, _ngay_hien_tai_str(), 3)
                if not kq_xh["nguoi"]:
                    notify_text(messaging_api, event.reply_token, target_id,
                                f"Chưa có ngày nào đã CHỐT SỐ có mục tiêu {'THU NGÂN' if vai_tro_xh == 'thu_ngan' else 'FRESH'} "
                                "anh ạ – cuối ngày anh gõ CHỐT SỐ thì bot mới lưu kết quả từng bạn.")
                    return
                send_card_with_ack(messaging_api, event.reply_token, target_id, FlexMessage(
                    alt_text=f"{'Thu ngân' if vai_tro_xh == 'thu_ngan' else 'Fresh'} – tổng kết {len(kq_xh['ngays'])} ngày",
                    contents=FlexContainer.from_dict(muc_tieu_flex.build_the_xep_hang(kq_xh))))
            except Exception as e:
                traceback.print_exc()
                try:
                    notify_text(messaging_api, event.reply_token, target_id, f"Có lỗi khi tổng kết: {e}")
                except Exception:
                    traceback.print_exc()
            return
        if CHOT_SO_PATTERN.search(text_kd) and len(text_kd) <= 60:
            try:
                ngay_chot = _ngay_hien_tai_str()
                m_ngay = re.search(r"(\d{1,2})\s*[/-]\s*(\d{1,2})", text_kd)
                if m_ngay:
                    ngay_chot = f"{_now_vn_datetime().year}-{int(m_ngay.group(2)):02d}-{int(m_ngay.group(1)):02d}"
                kq_chot = muc_tieu_ngay.tinh_chot_so(ngay_chot)
                if kq_chot is not None:
                    try:
                        luu = {vt: [{k: d.get(k) for k in ("ten", "ca", "mt", "bc", "pct", "da_bao")}
                                    for d in kq_chot["bang"].get(vt, {}).get("dong", [])]
                               for vt in ("thu_ngan", "fresh")}
                        luu["bao_vuot"] = kq_chot.get("bao_vuot") or []  # (03/11/2026) để lệnh THU NGÂN/FRESH đánh dấu BC KHỐNG
                        storage.save_chot_so(ngay_chot, luu)
                    except Exception:
                        traceback.print_exc()
                if kq_chot is None:
                    notify_text(messaging_api, event.reply_token, target_id,
                                "Tháng này chưa cài mục tiêu ngày nên chưa chốt số được anh ạ.")
                    return
                bubble = muc_tieu_flex.build_the_chot_so(ngay_chot, kq_chot)
                alt = f"Chốt số {ngay_chot[8:]}/{ngay_chot[5:7]}"
                # (02/11/2026) Đông người -> thẻ nặng: LINE giới hạn ~50KB/1 tin carousel, nên tách
                # thẻ cửa hàng và thẻ từng bạn thành 2 tin (vẫn chung 1 lần trả lời, miễn phí)
                if bubble.get("type") == "carousel" and len(bubble["contents"]) > 1 and \
                        len(json.dumps(bubble, ensure_ascii=False, separators=(",", ":")).encode()) > 45000:
                    phan_sau = bubble["contents"][1:]
                    the_sau = phan_sau[0] if len(phan_sau) == 1 else {"type": "carousel", "contents": phan_sau}
                    tin = [FlexMessage(alt_text=alt, contents=FlexContainer.from_dict(bubble["contents"][0])),
                           FlexMessage(alt_text=alt + " · từng bạn", contents=FlexContainer.from_dict(the_sau))]
                else:
                    tin = FlexMessage(alt_text=alt, contents=FlexContainer.from_dict(bubble))
                send_card_with_ack(messaging_api, event.reply_token, target_id, tin)
            except Exception as e:
                traceback.print_exc()
                try:
                    notify_text(messaging_api, event.reply_token, target_id, f"Có lỗi khi chốt số: {e}")
                except Exception:
                    traceback.print_exc()
            return
        # (09/11/2026) 5 NÚT MENU (chat riêng) — gõ đúng chữ trong nhóm cũng ra thẻ:
        # DOANH THU LŨY KẾ · LN FRESH · MTKM (có sẵn) · MỤC TIÊU TỪNG BẠN · KẾT QUẢ 7 NGÀY · CÀI MENU
        # (10/11/2026) + "MENU": thẻ 5 nút bấm trong nhóm
        _lenh_menu = menu_bot.nhan_lenh(text_kd)
        if _lenh_menu:
            try:
                _xu_ly_lenh_menu(_lenh_menu, event, messaging_api, api_client, source_type, target_id)
            except Exception as e:
                traceback.print_exc()
                try:
                    notify_text(messaging_api, event.reply_token, target_id, f"Có lỗi khi tạo bảng: {e}")
                except Exception:
                    traceback.print_exc()
            return
        # (02/11/2026) Anh sửa số báo cáo:
        #   "XÓA BC <tên> [dd/mm]"           -> xoá báo cáo của bạn đó (VD ghi nhầm)
        #   "BC <tên> [dd/mm]: <số bán>"      -> nhập hộ / sửa hộ số của bạn đó (VD "BC Tiền 01/11: c2 8 lốc, sk 1tr7")
        _m_xoa = re.match(r"^\s*xoa\s*(?:bc|bao\s*cao)\s+([a-z]+)(?:\s+(\d{1,2})\s*/\s*(\d{1,2}))?\s*$", text_kd)
        _m_ho = re.match(r"^\s*(?:bc|báo cáo|bao cao)\s+([^\s:/\d]+)\s*(?:(\d{1,2})\s*/\s*(\d{1,2}))?\s*:\s*(.+)$",
                         unicodedata.normalize("NFC", text).lower(), re.S)
        for _m_lenh in (_m_xoa, _m_ho):
            if not _m_lenh:
                continue
            _ten_lenh = muc_tieu_ngay.nhan_ten(_m_lenh.group(1))
            if not _ten_lenh:
                continue
            _ngay_lenh = _ngay_hien_tai_str()
            if _m_lenh.group(2):
                _ngay_lenh = f"{_now_vn_datetime().year}-{int(_m_lenh.group(3)):02d}-{int(_m_lenh.group(2)):02d}"
            _nhan_ngay = f"{_ngay_lenh[8:]}/{_ngay_lenh[5:7]}"
            try:
                if _m_lenh is _m_xoa:
                    _ok = storage.xoa_bao_cao_ban_hang(_ngay_lenh, _ten_lenh)
                    notify_text(messaging_api, event.reply_token, target_id,
                                f"Đã xoá báo cáo của {_ten_lenh} ngày {_nhan_ngay}." if _ok
                                else f"Ngày {_nhan_ngay} chưa có báo cáo nào của {_ten_lenh}.")
                else:
                    _it, _gh = muc_tieu_ngay.doc_bao_cao(_m_lenh.group(4))
                    if not _it:
                        notify_text(messaging_api, event.reply_token, target_id,
                                    "Em chưa đọc được số nào, anh ghi lại giúp em (VD: BC Tiền 01/11: c2 8 lốc, sk 1tr7)")
                    else:
                        storage.save_bao_cao_ban_hang(_ngay_lenh, _ten_lenh, _it, "nhập hộ " + _now_vn_time_str())
                        notify_text(messaging_api, event.reply_token, target_id,
                                    f"(Nhập hộ {_nhan_ngay}) " + muc_tieu_ngay.tom_tat_bao_cao(_ten_lenh, _it, _gh))
            except Exception:
                traceback.print_exc()
            return
        # (01/11/2026) BÁO CÁO BÁN HÀNG CUỐI NGÀY của các bạn — nhắn tự do trong nhóm
        # (VD "Nay em bán trứng đc 8 vỉ Đm 1079 skho 1504 c2 6 lốc tt anh @Quí").
        # Bot đọc số, lưu theo tên đã đăng ký, trả 1 dòng nhắc lại số đã hiểu.
        if source_type == "group" and text:
            try:
                _mention_obj = getattr(event.message, "mention", None)
                _mens = (getattr(_mention_obj, "mentionees", None) or []) if _mention_obj else []
                _che = []
                for _m in _mens:
                    _ix = _mention_index_py(text, getattr(_m, "index", None))
                    _ln = getattr(_m, "length", None)
                    if _ix is not None and _ln:
                        _che.append((_ix, _ix + _ln))
                _items, _ghi = muc_tieu_ngay.doc_bao_cao(text, _che)
                # tag bot (hỏi bot) thì không tính là tag báo cáo — các bạn báo cáo thường tag anh Quí
                if muc_tieu_ngay.la_tin_bao_cao(text, _items, co_tag=bool(_mens) and not da_tag_bot):
                    _uid = getattr(event.source, "user_id", None)
                    _ten = storage.get_ten_ngan_tu_user_id(_uid) if _uid else None
                    if _ten:
                        storage.save_bao_cao_ban_hang(_ngay_hien_tai_str(), _ten, _items, _now_vn_time_str())
                        # reply lỗi (VD bot vừa khởi động lại, tin đến trễ) thì gửi bằng push để bạn vẫn thấy
                        notify_text(messaging_api, event.reply_token, target_id,
                                    muc_tieu_ngay.tom_tat_bao_cao(_ten, _items, _ghi))
                        return
                    if _uid:
                        storage.save_bao_cao_cho_dk(_uid, _ngay_hien_tai_str(), _items, _now_vn_time_str())
                    notify_text(messaging_api, event.reply_token, target_id,
                                "Em đã giữ số bạn báo nhưng chưa biết tài khoản này là bạn nào. "
                                "Bạn gõ \"DK <tên>\" 1 lần (VD: DK Tiền) là em tự ghi vào báo cáo của bạn, không cần nhắn lại số.")
                    return
            except Exception:
                traceback.print_exc()
        # Lệnh "kiểm tra drive" (16/09/2026, ĐƠN GIẢN HOÁ THEO YÊU CẦU ANH QUÍ) —
        # chạy tay ngay lập tức thay vì đợi tới lịch 4 tiếng/lần. Chỉ trả lời
        # ĐÚNG nơi anh gõ lệnh (chat riêng hoặc nhóm), KHÔNG tự nhắn gì thêm
        # vào nhóm/chat khác. Không kể chi tiết đã nạp file gì, chỉ báo xong.
        if DRIVE_CHECK_COMMAND_PATTERN.match(text_kd):
            if not GOOGLE_DRIVE_FOLDER_ID:
                reply_text(messaging_api, event.reply_token,
                           "Chưa cấu hình GOOGLE_DRIVE_FOLDER_ID trên Railway, chưa bật được tính năng này.")
                return
            try:
                check_google_drive_for_new_files()
                reply_text(messaging_api, event.reply_token, "Dạ em kiểm tra xong ạ")
            except Exception as e:
                traceback.print_exc()
                reply_text(messaging_api, event.reply_token, f"Có lỗi khi kiểm tra Drive: {e}")
            return
        # Lệnh DT — báo cáo doanh thu
        if DT_COMMAND_PATTERN.match(text):
            try:
                result = build_revenue_report_messages(_base_url())
                if result is None:
                    notify_text(messaging_api, event.reply_token, target_id, "Chưa có dữ liệu doanh thu nào được lưu. Anh gửi file Excel doanh thu trước nhé.")
                    return
                flex_message = result
                send_card_with_ack(messaging_api, event.reply_token, target_id, flex_message)
                storage.save_last_command(target_id, "DT", _now_vn_time_str())
                _xoa_anh_gan_nhat(target_id)
            except Exception as e:
                traceback.print_exc()
                try:
                    notify_text(messaging_api, event.reply_token, target_id, f"Có lỗi khi tạo báo cáo doanh thu: {e}")
                except Exception:
                    traceback.print_exc()
            return
        # Lệnh DTDK — Doanh Thu Dự Kiến (báo cáo riêng, theo tháng + tiến độ target năm)
        if DTDK_COMMAND_PATTERN.match(text):
            try:
                ten_st = "BHX_STR_CLD - Thửa 1289 An Nghiệp"
                bubble = dtdk_report.build_dtdk(ten_st)
                if bubble is None:
                    reply_text(messaging_api, event.reply_token,
                               "Chưa có dữ liệu doanh thu nào được lưu. Anh gửi file Excel doanh thu trước nhé.")
                    return
                reply_flex(messaging_api, event.reply_token, "Doanh thu theo tháng (DTDK)", bubble)
            except Exception as e:
                traceback.print_exc()
                try:
                    reply_text(messaging_api, event.reply_token, f"Có lỗi khi tạo báo cáo DTDK: {e}")
                except Exception:
                    traceback.print_exc()
            return
        # Lệnh MỤC TIÊU KHUYẾN MÃI — báo cáo ngành hàng
        if MTKM_COMMAND_PATTERN.match(text):
            try:
                flex_message = build_category_report_message()
                if flex_message is None:
                    notify_text(messaging_api, event.reply_token, target_id, "Chưa có dữ liệu ngành hàng nào được lưu. Anh gửi file Excel chi tiết ngành hàng trước nhé.")
                    return
                send_card_with_ack(messaging_api, event.reply_token, target_id, flex_message)
                storage.save_last_command(target_id, "MTKM", _now_vn_time_str())
                _xoa_anh_gan_nhat(target_id)
            except Exception as e:
                traceback.print_exc()
                try:
                    notify_text(messaging_api, event.reply_token, target_id, f"Có lỗi khi tạo báo cáo ngành hàng: {e}")
                except Exception:
                    traceback.print_exc()
            return
        # Lệnh TD / THƯỞNG — báo cáo thưởng FRESH + FMCG
        if TD_COMMAND_PATTERN.match(text):
            try:
                flex_message = build_thuong_report_message()
                if flex_message is None:
                    notify_text(messaging_api, event.reply_token, target_id, "Chưa có dữ liệu thưởng nào được lưu. Anh gửi file Excel doanh thu chi tiết (từ đầu tháng đến hiện tại) trước nhé.")
                    return
                send_card_with_ack(messaging_api, event.reply_token, target_id, flex_message)
                storage.save_last_command(target_id, "TD", _now_vn_time_str())
            except Exception as e:
                traceback.print_exc()
                try:
                    notify_text(messaging_api, event.reply_token, target_id, f"Có lỗi khi tạo báo cáo thưởng: {e}")
                except Exception:
                    traceback.print_exc()
            return
        # Lệnh TTFF — TIẾN ĐỘ THƯỞNG FRESH/FMCG (MỚI 12/09/2026, Base cố định,
        # TÁCH RIÊNG khỏi TD/THƯỞNG cũ). Dùng đúng 1 file "doanh thu chi tiết"
        # đã gửi cho MTKM/TD — không cần gửi thêm file nào khác.
        if TTFF_COMMAND_PATTERN.match(text):
            try:
                ten_st = "BHX_STR_CLD - Thửa 1289 An Nghiệp"
                payload = thuong_freshfmcg.build_report_payload(ten_st)
                if payload is None:
                    notify_text(messaging_api, event.reply_token, target_id,
                               "Chưa có dữ liệu cho báo cáo này. Anh gửi file Excel \"doanh thu chi tiết\" trước nhé.")
                    return
                bubble = thuong_freshfmcg.build_freshfmcg_flex_message(payload)
                flex_message = FlexMessage(
                    alt_text="Tiến độ thưởng Fresh & FMCG",
                    contents=FlexContainer.from_dict(bubble),
                )
                send_card_with_ack(messaging_api, event.reply_token, target_id, flex_message)
                storage.save_last_command(target_id, "TTFF", _now_vn_time_str())
            except Exception as e:
                traceback.print_exc()
                try:
                    notify_text(messaging_api, event.reply_token, target_id, f"Có lỗi khi tạo báo cáo tiến độ thưởng: {e}")
                except Exception:
                    traceback.print_exc()
            return
        # Lệnh HỦY MMKK — Công việc 1 (đúng 1 ngày) / Công việc 2 (nhiều ngày)
        # / so sánh 2 ngày cụ thể (MỚI 08/09/2026, đã CHỈNH LẠI theo phản hồi
        # của anh Quí: "so sánh" trả lời bằng văn phong PHÂN TÍCH tự nhiên —
        # giống đang trò chuyện 1-1 — chứ KHÔNG phải 2 thẻ card số liệu thô
        # đặt cạnh nhau). Trong nhóm: bắt buộc tag tên bot. Chat riêng 1-1:
        # không cần tag (giống DT/MTKM) — trước đây lệnh này CHỈ hoạt động
        # trong nhóm, hỏi riêng 1-1 sẽ bị rớt xuống luồng AI trả lời tự do.
        # 13/09/2026: cho phép "so sánh ngày X và Y" kích hoạt CHỈ CẦN có tag
        # bot (trong nhóm) mà KHÔNG cần thêm chữ "hủy mmkk" — trước đây phải
        # gõ đủ "hủy mmkk so sánh ngày..." mới khớp, anh Quí phản hồi muốn gõ
        # tag xong là ra luôn cho gọn.
        _co_so_sanh_pre, _ngay_a_pre, _ngay_b_pre = fresh_report.parse_so_sanh_2_ngay(text)
        if (HUY_MMKK_TRIGGER.search(text_kd) or _co_so_sanh_pre) and (source_type != "group" or da_tag_bot):
            try:
                ten_st = "BHX_STR_CLD - Thửa 1289 An Nghiệp"
                co_so_sanh, ngay_a, ngay_b = _co_so_sanh_pre, _ngay_a_pre, _ngay_b_pre
                if co_so_sanh:
                    ket_qua = fresh_report.build_so_sanh_2_ngay(ngay_a, ngay_b)
                    if ket_qua is None:
                        reply_text(messaging_api, event.reply_token,
                                   f"Chưa có dữ liệu HỦY TỒN + MMKK ngày {ngay_a.strftime('%d/%m')} lẫn ngày {ngay_b.strftime('%d/%m')}. Anh gửi file cho bot trước nhé.")
                        return
                    reply_text(messaging_api, event.reply_token, ket_qua)
                    return
                mode, ngay_tu, ngay_den = fresh_report.parse_date_request(text)
                if mode == "single":
                    bubble = fresh_report.build_cong_viec_1(ten_st, ngay_tu)
                    if bubble is None:
                        reply_text(messaging_api, event.reply_token,
                                   f"Chưa có dữ liệu HỦY TỒN + MMKK ngày {ngay_tu.strftime('%d/%m/%Y')}. Anh gửi file cho bot trước nhé.")
                        return
                    reply_flex(messaging_api, event.reply_token,
                               f"Hủy tồn + MMKK {ngay_tu.strftime('%d/%m/%Y')}", bubble)
                    storage.save_last_command(target_id, f"huy_mmkk:{ngay_tu.strftime('%Y-%m-%d')}", _now_vn_time_str())
                else:
                    bubble = fresh_report.build_cong_viec_2(ten_st, ngay_tu, ngay_den)
                    if bubble is None:
                        reply_text(messaging_api, event.reply_token,
                                   f"Chưa có dữ liệu HỦY TỒN + MMKK trong khoảng {ngay_tu.strftime('%d/%m')} - {ngay_den.strftime('%d/%m')}. Anh gửi file cho bot trước nhé.")
                        return
                    reply_flex(messaging_api, event.reply_token,
                               "Tổng hủy tồn + MMKK theo nhóm", bubble)
            except Exception as e:
                traceback.print_exc()
                try:
                    reply_text(messaging_api, event.reply_token, f"Có lỗi khi xử lý hủy tồn/MMKK: {e}")
                except Exception:
                    traceback.print_exc()
            return
        # Lệnh DOANH THU THỦY HẢI SẢN — tương tự, cho phép cả chat riêng 1-1.
        if SEAFOOD_TRIGGER.search(text_kd) and (source_type != "group" or da_tag_bot):
            try:
                ten_st = "BHX_STR_CLD - Thửa 1289 An Nghiệp"
                bubble = fresh_report.build_doanh_thu_thuy_hai_san(ten_st)
                if bubble is None:
                    reply_text(messaging_api, event.reply_token,
                               "Chưa có dữ liệu HỦY TỒN + MMKK nào được lưu. Anh gửi file cho bot trước nhé.")
                    return
                reply_flex(messaging_api, event.reply_token, "Doanh thu thủy hải sản", bubble)
            except Exception as e:
                traceback.print_exc()
                try:
                    reply_text(messaging_api, event.reply_token, f"Có lỗi khi xử lý doanh thu thủy hải sản: {e}")
                except Exception:
                    traceback.print_exc()
            return
        # Lệnh PHÂN TÍCH SỐ LIỆU — thứ tự ưu tiên:
        # 1. Có ẢNH gần nhất còn hiệu lực (gửi trong vòng 2 tiếng, đúng đoạn
        #    chat này) -> đọc ảnh bằng Claude Vision rồi phân tích.
        # 2. Không có ảnh, nhưng lệnh gần nhất trong đoạn chat này là "DT",
        #    "MTKM" HOẶC "TD" (MỞ RỘNG 08/09/2026, trước chỉ có DT/MTKM) ->
        #    phân tích luôn báo cáo đó bằng data đã có sẵn trong hệ thống
        #    (dùng Claude, không cần gửi ảnh/file gì thêm).
        # 3. Không khớp 2 trường hợp trên -> quay lại luồng cũ: bắt buộc
        #    phải gõ "hủy mmkk <ngày>" trước, có kết quả rồi mới phân tích.
        if (PHAN_TICH_TRIGGER.search(text_kd) and (source_type != "group" or da_tag_bot)
                and not HUY_MMKK_TRIGGER.search(text_kd)):
            anh_data = _lay_anh_gan_nhat(target_id)
            if anh_data is not None:
                try:
                    ket_qua = ai_assistant.phan_tich_anh(anh_data, target_id=target_id)
                    reply_text(messaging_api, event.reply_token, ket_qua)
                except Exception:
                    traceback.print_exc()
                    try:
                        reply_text(messaging_api, event.reply_token, "Có lỗi khi em phân tích ảnh, thử lại giúp em nhé.")
                    except Exception:
                        traceback.print_exc()
                return
            last_cmd_bao_cao = storage.get_last_command(target_id)
            if last_cmd_bao_cao in ("DT", "MTKM", "TD"):
                try:
                    if last_cmd_bao_cao == "DT":
                        ket_qua = ai_assistant.phan_tich_doanh_thu(target_id=target_id)
                    elif last_cmd_bao_cao == "MTKM":
                        ket_qua = ai_assistant.phan_tich_nganh_hang(target_id=target_id)
                    else:
                        ket_qua = ai_assistant.phan_tich_thuong(target_id=target_id)
                    reply_text(messaging_api, event.reply_token, ket_qua)
                except Exception:
                    traceback.print_exc()
                    try:
                        reply_text(messaging_api, event.reply_token, "Có lỗi khi em phân tích, thử lại giúp em nhé.")
                    except Exception:
                        traceback.print_exc()
                return
            try:
                _mode, ngay, _ngay_den = fresh_report.parse_date_request(text)
                ngay_key = f"huy_mmkk:{ngay.strftime('%Y-%m-%d')}"
                last_cmd = storage.get_last_command(target_id)
                if last_cmd != ngay_key:
                    reply_text(
                        messaging_api, event.reply_token,
                        f"Anh tag bot + gõ \"hủy mmkk {ngay.strftime('%d/%m')}\" trước, "
                        f"có kết quả xong mới gõ \"phân tích số liệu\" nhé."
                    )
                    return
                so_sanh = fresh_report.wants_comparison(text)
                ket_qua = fresh_report.build_phan_tich(ngay, so_sanh_voi_hom_qua=so_sanh)
                if ket_qua is None:
                    reply_text(messaging_api, event.reply_token,
                               f"Chưa có dữ liệu ngày {ngay.strftime('%d/%m/%Y')}. Anh gửi file cho bot trước nhé.")
                    return
                reply_text(messaging_api, event.reply_token, ket_qua)
            except Exception as e:
                traceback.print_exc()
                try:
                    reply_text(messaging_api, event.reply_token, f"Có lỗi khi phân tích: {e}")
                except Exception:
                    traceback.print_exc()
            return
        # Lệnh "NHẬN XÉT <loại data>" (MỚI 16/09/2026) — anh Quí tag bot, nói
        # thẳng muốn nhận xét cái gì (mmkk/hủy tồn, doanh thu, khuyến mãi/
        # ngành hàng), có thể kèm ngày cụ thể hoặc "hôm nay" ngay trong câu.
        # KHÔNG cần gõ lệnh lấy báo cáo trước — bot tự lấy đúng data tương
        # ứng rồi nhận xét luôn. VD: "nhận xét mmkk ngày 15/09 @TROLY",
        # "nhận xét doanh thu hôm nay như nào @TROLY", "nhận xét kết quả
        # khuyến mãi hôm nay @TROLY".
        if (NHAN_XET_TRIGGER.search(text_kd) and (source_type != "group" or da_tag_bot)
                and (NHAN_XET_MMKK_RE.search(text_kd) or NHAN_XET_DOANH_THU_RE.search(text_kd)
                     or NHAN_XET_KM_RE.search(text_kd))):
            try:
                if NHAN_XET_MMKK_RE.search(text_kd):
                    mode, ngay_tu, ngay_den = fresh_report.parse_date_request(text)
                    if mode == "single":
                        ket_qua = fresh_report.build_phan_tich(
                            ngay_tu, so_sanh_voi_hom_qua=fresh_report.wants_comparison(text)
                        )
                        if ket_qua is None:
                            ket_qua = f"Chưa có dữ liệu Hủy tồn + MMKK ngày {ngay_tu.strftime('%d/%m/%Y')}."
                    else:
                        ket_qua = ai_assistant.phan_tich_fresh_khoang_ngay(ngay_tu, ngay_den, target_id=target_id)
                elif NHAN_XET_DOANH_THU_RE.search(text_kd):
                    ket_qua = ai_assistant.phan_tich_doanh_thu(target_id=target_id)
                else:
                    ket_qua = ai_assistant.phan_tich_nganh_hang(target_id=target_id)
                reply_text(messaging_api, event.reply_token, ket_qua)
            except Exception:
                traceback.print_exc()
                try:
                    reply_text(messaging_api, event.reply_token, "Có lỗi khi em nhận xét, thử lại giúp em nhé.")
                except Exception:
                    traceback.print_exc()
            return
        # Lệnh NHẬN XÉT MỤC TIÊU (MỚI 08/09/2026) — anh Quí tag bot + nêu tên
        # 1 bạn nhân viên + có ý "nhận xét" (VD "nhận xét giúp anh Quyên hôm
        # nay") -> bot so mục tiêu bạn đó được giao qua bài phân line với
        # báo cáo thực tế (nhật ký tin nhắn bạn đó tự gửi trong nhóm).
        if NHAN_XET_TRIGGER.search(text_kd) and (source_type != "group" or da_tag_bot):
            try:
                ten_match = TEN_NHAN_XET_RE.search(text)
                if not ten_match:
                    reply_text(
                        messaging_api, event.reply_token,
                        "Anh nêu rõ tên bạn nhân viên trong câu giúp em nhé "
                        "(VD: \"nhận xét giúp anh Quyên hôm nay\")."
                    )
                    return
                ten_chuan = ten_match.group(1)
                user_id_nv = storage.get_user_id_da_dang_ky(ten_chuan)
                if not user_id_nv:
                    reply_text(
                        messaging_api, event.reply_token,
                        f"Em chưa có ID LINE của {ten_chuan} (chưa từng được tag trong bài phân line, "
                        f"cũng chưa đăng ký). Nhờ {ten_chuan} gõ \"DK {ten_chuan}\" trong nhóm 1 lần giúp em."
                    )
                    return
                ngay_str = _ngay_hom_qua_str() if "hom qua" in text_kd else _ngay_hien_tai_str()
                phan_line_data = storage.get_phan_line(ngay_str) or {}
                muc_tieu_text = ""
                for ca in ("sang", "chieu"):
                    mt = phan_line_data.get(ca, {}).get("muc_tieu", {}).get(user_id_nv)
                    if mt:
                        muc_tieu_text = mt
                        break
                bao_cao_rows = storage.get_nhat_ky_nhan_vien(ngay_str, user_id_nv)
                if not muc_tieu_text and not bao_cao_rows:
                    reply_text(
                        messaging_api, event.reply_token,
                        f"Em chưa có mục tiêu lẫn tin nhắn báo cáo của {ten_chuan} ngày "
                        f"{ngay_str}. Anh kiểm tra lại đã đăng bài phân line có tag {ten_chuan} "
                        f"chưa, và {ten_chuan} đã nhắn báo cáo trong nhóm chưa nhé."
                    )
                    return
                ket_qua = ai_assistant.nhan_xet_muc_tieu(
                    ten_chuan, ngay_str, muc_tieu_text, bao_cao_rows, target_id=target_id
                )
                reply_text(messaging_api, event.reply_token, ket_qua)
            except Exception as e:
                traceback.print_exc()
                try:
                    reply_text(messaging_api, event.reply_token, f"Có lỗi khi em nhận xét: {e}")
                except Exception:
                    traceback.print_exc()
            return
        # ---- TROLY TRẢ LỜI TỰ DO (MỚI - dùng Claude API) ----
        # Không khớp lệnh cố định nào ở trên -> coi như câu hỏi tự do.
        # Trong nhóm: bắt buộc phải tag tên bot mới trả lời (tránh bot xen
        # vào chat thường của mọi người). Chat riêng 1-1 với bot: hỏi gì
        # cũng được, không cần tag vì đã là nhắn riêng cho bot rồi.
        if source_type == "group" and not da_tag_bot:
            return
        cau_hoi = BOT_NAME_STRIP_PATTERN.sub("", text).strip()
        if not cau_hoi:
            return
        try:
            tra_loi = ai_assistant.hoi_ai(cau_hoi, target_id=target_id)
            reply_text(messaging_api, event.reply_token, tra_loi)
        except Exception:
            traceback.print_exc()
            try:
                reply_text(messaging_api, event.reply_token, "Em chưa trả lời được câu này, thử lại sau giúp em nhé.")
            except Exception:
                traceback.print_exc()
        return
def push_text(messaging_api, target_id, text):
    text = text[:4900]
    messaging_api.push_message(
        PushMessageRequest(to=target_id, messages=[TextMessage(text=text)])
    )
@handler.add(PostbackEvent)
def handle_postback(event):
    """(10/11/2026) Bấm ô trong MENU (thẻ trong nhóm hoặc menu chat riêng): không hiện chữ trong khung chat,
    bot chỉ trả đúng bảng."""
    lenh = menu_bot.lenh_tu_postback(getattr(event.postback, "data", ""))
    if not lenh:
        return
    with ApiClient(configuration) as api_client:
        messaging_api = MessagingApi(api_client)
        source_type = event.source.type
        target_id = event.source.group_id if source_type == "group" else getattr(event.source, "user_id", None)
        print(f"[MENU-POSTBACK] {lenh} tu {source_type}")
        try:
            _xu_ly_lenh_menu(lenh, event, messaging_api, api_client, source_type, target_id, chi_bang=True)
        except Exception as e:
            traceback.print_exc()
            try:
                notify_text(messaging_api, event.reply_token, target_id, f"Có lỗi khi tạo bảng: {e}")
            except Exception:
                traceback.print_exc()
def _gui_bang(messaging_api, reply_token, target_id, flex_message, chi_bang):
    """chi_bang=True (bấm nút menu): chỉ trả thẻ, không kèm câu "Em gửi Anh và Team..."."""
    if not chi_bang:
        send_card_with_ack(messaging_api, reply_token, target_id, flex_message)
        return
    try:
        messaging_api.reply_message(ReplyMessageRequest(reply_token=reply_token, messages=[flex_message]))
    except Exception:
        traceback.print_exc()
        messaging_api.push_message(PushMessageRequest(to=target_id, messages=[flex_message]))
import threading as _threading
_KHOA_DRIVE = _threading.Lock()
def _nap_drive_truoc_khi_tra_bang():
    """(10/11/2026) Trước khi trả bảng cần số POS (doanh thu, LN Fresh, mục tiêu ngày): kiểm tra thư mục Drive
    của bot, có file POS mới (em vừa nạp) thì nạp luôn — không phải đợi lịch 4 tiếng."""
    if not _KHOA_DRIVE.acquire(blocking=False):
        return
    try:
        so = check_google_drive_for_new_files()
        if so:
            print(f"[MENU-DRIVE] nap them {so} file truoc khi tra bang")
    except Exception:
        traceback.print_exc()
    finally:
        _KHOA_DRIVE.release()
def _xu_ly_lenh_menu(lenh, event, messaging_api, api_client, source_type, target_id, chi_bang=False):
    """(09/11/2026) Trả thẻ cho 4 nút menu mới + lệnh CÀI MENU. (10/11) + "mtkm" cho nút bấm."""
    hom_nay = _ngay_hien_tai_str()
    if lenh in ("doanh_thu", "ln_fresh", "mtkm"):
        _nap_drive_truoc_khi_tra_bang()
    if lenh == "mtkm":
        flex_message = build_category_report_message()
        if flex_message is None:
            notify_text(messaging_api, event.reply_token, target_id,
                        "Chưa có dữ liệu ngành hàng nào được lưu. Anh gửi file Excel chi tiết ngành hàng trước nhé.")
            return
        _gui_bang(messaging_api, event.reply_token, target_id, flex_message, chi_bang)
        return
    if lenh == "menu":  # (10/11/2026) thẻ 5 nút, gõ ở nhóm hay chat riêng đều được
        reply_flex(messaging_api, event.reply_token, "Menu TROLY", menu_bot.build_the_menu())
        return
    if lenh == "cai_menu":
        if source_type != "user":
            notify_text(messaging_api, event.reply_token, target_id,
                        "Menu chỉ hiện trong chat riêng với bot – anh vào chat riêng với TROLY rồi gõ CÀI MENU nhé.")
            return
        menu_bot.cai_menu_cho(event.source.user_id, messaging_api, MessagingApiBlob(api_client))
        notify_text(messaging_api, event.reply_token, target_id,
                    "✅ Đã gắn menu. Anh bấm \"Menu\" ở dưới khung chat (nếu chưa thấy thì thoát ra vào lại chat).")
        return
    if lenh == "doanh_thu":
        kq = menu_bot.tinh_doanh_thu(hom_nay)
        if kq is None:
            notify_text(messaging_api, event.reply_token, target_id,
                        "Tháng này chưa có file doanh thu chi tiết (POS 76) của ngày nào trước hôm nay anh ạ.")
            return
        bubble, alt = menu_bot.build_the_doanh_thu(kq), "Doanh thu lũy kế"
    elif lenh == "ln_fresh":
        p = menu_bot.lay_ln_fresh_moi_nhat()
        if p is None:
            notify_text(messaging_api, event.reply_token, target_id,
                        "Bot chưa có file LN Fresh (POS 3933) anh ạ – anh gửi file vào chat hoặc bỏ vào thư mục Drive của bot.")
            return
        bubble, alt = menu_bot.build_the_ln_fresh(p), "LN Fresh"
    elif lenh == "muc_tieu_tung_ban":
        bubble, loi = menu_bot.build_the_muc_tieu_hom_nay(hom_nay)
        if bubble is None:
            notify_text(messaging_api, event.reply_token, target_id, loi)
            return
        alt = f"Mục tiêu từng bạn {hom_nay[8:]}/{hom_nay[5:7]}"
    else:  # bay_ngay
        bubble = menu_bot.build_the_7_ngay(menu_bot.tinh_7_ngay(hom_nay))
        if bubble is None:
            notify_text(messaging_api, event.reply_token, target_id,
                        "Chưa có ngày nào đã CHỐT SỐ có mục tiêu từng bạn anh ạ.")
            return
        alt = "Kết quả 7 ngày gần nhất"
    _gui_bang(messaging_api, event.reply_token, target_id,
              FlexMessage(alt_text=alt, contents=FlexContainer.from_dict(bubble)), chi_bang)
def notify_text(messaging_api, reply_token, target_id, text):
    """Gửi text ưu tiên bằng REPLY (miễn phí, không tính hạn mức tin nhắn tháng
    của LINE OA); chỉ khi reply lỗi (token hết hạn...) mới thử push."""
    try:
        reply_text(messaging_api, reply_token, text)
    except Exception:
        traceback.print_exc()
        push_text(messaging_api, target_id, text)
def send_card_with_ack(messaging_api, reply_token, target_id, flex_message):
    """Gửi câu xác nhận + thẻ báo cáo trong CÙNG 1 reply (miễn phí). Trước đây
    thẻ gửi bằng push nên khi LINE OA hết hạn mức tin nhắn tháng (HTTP 429
    "You have reached your monthly limit") bot chỉ trả lời được câu xác nhận
    mà không có thẻ. Nếu reply lỗi thì mới thử push như cũ.
    flex_message có thể là 1 thẻ hoặc danh sách thẻ (tối đa 4, vì 1 reply được 5 tin)."""
    the = flex_message if isinstance(flex_message, list) else [flex_message]
    try:
        messaging_api.reply_message(
            ReplyMessageRequest(
                reply_token=reply_token,
                messages=[TextMessage(text="Em gửi Anh và Team 8363 luôn ạ")] + the[:4],
            )
        )
    except Exception:
        traceback.print_exc()
        messaging_api.push_message(
            PushMessageRequest(to=target_id, messages=the[:5])
        )
def reply_text(messaging_api, reply_token, text):
    text = text[:4900]
    messaging_api.reply_message(
        ReplyMessageRequest(reply_token=reply_token, messages=[TextMessage(text=text)])
    )
def reply_flex(messaging_api, reply_token, alt_text, bubble):
    flex_message = FlexMessage(
        alt_text=alt_text,
        contents=FlexContainer.from_dict(bubble),
    )
    messaging_api.reply_message(
        ReplyMessageRequest(reply_token=reply_token, messages=[flex_message])
    )
def reply_flex_multi(messaging_api, reply_token, alt_text, bubbles):
    """Gửi nhiều thẻ Flex (card) cùng lúc trong 1 tin nhắn, dạng carousel
    (vuốt ngang xem từng thẻ) — dùng cho "so sánh ngày X và ngày Y"."""
    carousel = {"type": "carousel", "contents": bubbles}
    flex_message = FlexMessage(
        alt_text=alt_text,
        contents=FlexContainer.from_dict(carousel),
    )
    messaging_api.reply_message(
        ReplyMessageRequest(reply_token=reply_token, messages=[flex_message])
    )
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
