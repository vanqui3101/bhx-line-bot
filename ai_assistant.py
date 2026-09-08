"""
ai_assistant.py - TROLY TRẢ LỜI TỰ DO (tích hợp Claude API).

Cho phép bot trả lời câu hỏi TỰ DO (không cần đúng cú pháp lệnh cố định),
dựa trên dữ liệu ĐÃ LƯU trong SQLite (doanh thu, ngành hàng, thưởng, tồn
kho, hủy tồn+MMKK) + vài thông tin tĩnh về cửa hàng/nhân sự (khai báo ngay
trong file này, vì bot không có cách nào khác để biết những thứ đó).

CẤU HÌNH (Environment Variable mới, bắt buộc để tính năng này chạy được):
- ANTHROPIC_API_KEY   (lấy trong Anthropic Console -> API Keys)

NGUYÊN TẮC TRẢ LỜI (đã thống nhất với anh Quí):
- CHỈ trả lời dựa trên dữ liệu đưa vào system prompt bên dưới.
- Câu hỏi ngoài phạm vi dữ liệu đã có -> trả lời thẳng "chưa có dữ liệu",
  KHÔNG suy đoán/bịa số.
- Xưng "em", gọi người hỏi là "anh". Trả lời ngắn gọn, đi thẳng vào việc.
- Bot NHỚ được cuộc trò chuyện gần đây của TỪNG đoạn chat riêng (nhóm/1-1
  đều tách riêng, không lẫn giữa các nhóm/người khác nhau) — hỏi tiếp câu
  sau vẫn hiểu ngữ cảnh câu trước, không cần lặp lại từ đầu. Nhớ trong vòng
  2 tiếng không hỏi gì thêm thì quên (RAM, không lưu SQLite, mất khi bot
  khởi động lại).
- Khi thiếu thông tin để trả lời/phân tích chính xác, bot chủ động HỎI LẠI
  anh Quí cần biết thêm gì (giống cách 1 trợ lý thật sẽ hỏi lại), thay vì tự
  suy đoán hoặc chỉ trả lời cụt "chưa có dữ liệu".

GIỚI HẠN CỦA BẢN NÀY (nói rõ để anh Quí biết, tránh kỳ vọng sai):
- Chưa có data "Lợi nhuận theo ngành hàng" (giá vốn cơ bản) và chưa có data
  "Danh sách Fresh theo tần suất KM" (2 ngày / 2 lần/tuần / 1 lần/tuần) -
  hai phần này mới chỉ trao đổi qua chat, CHƯA có trong hệ thống bot. Nếu
  sau này muốn bot biết luôn 2 phần này thì cần bổ sung thêm (báo lại để
  làm tiếp), bản này CHƯA có.
- Nhân sự/giờ hoạt động bên dưới là DỮ LIỆU TĨNH (gõ chết trong code) — nếu
  có gì thay đổi (nhân viên mới, tính cách khác, đổi giờ mở cửa...) thì
  phải SỬA LẠI trong file này rồi deploy lại, bot không tự cập nhật được.
"""
import os
import time
import base64
import requests
from collections import defaultdict

import storage

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
ANTHROPIC_MODEL = "claude-sonnet-5"
ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"

# ---------------------------------------------------------------------------
# BỘ NHỚ HỘI THOẠI TẠM (RAM, không lưu SQLite) — mỗi đoạn chat (nhóm hoặc
# riêng) có lịch sử RIÊNG, không lẫn với đoạn chat khác. Quên sau 2 tiếng
# không hỏi gì thêm, và chỉ giữ tối đa 1 số lượt gần nhất để câu hỏi không
# phình to quá mức mỗi lần gọi API.
# ---------------------------------------------------------------------------
_LICH_SU_CHAT = {}  # target_id -> {"turns": [{"role":..., "content":...}], "luc": epoch}
LICH_SU_HET_HAN_GIAY = 2 * 60 * 60  # 2 tiếng
LICH_SU_TOI_DA_LUOT = 12  # tối đa 12 lượt (~6 câu hỏi + 6 câu trả lời)


def _lay_lich_su(target_id):
    if not target_id:
        return []
    info = _LICH_SU_CHAT.get(target_id)
    if not info:
        return []
    if time.time() - info["luc"] > LICH_SU_HET_HAN_GIAY:
        _LICH_SU_CHAT.pop(target_id, None)
        return []
    return info["turns"]


def _luu_luot_chat(target_id, vai_tro, noi_dung):
    if not target_id or not noi_dung:
        return
    info = _LICH_SU_CHAT.setdefault(target_id, {"turns": [], "luc": time.time()})
    info["turns"].append({"role": vai_tro, "content": noi_dung})
    info["turns"] = info["turns"][-LICH_SU_TOI_DA_LUOT:]
    info["luc"] = time.time()

# ---------------------------------------------------------------------------
# DỮ LIỆU TĨNH (gõ tay, sửa trực tiếp ở đây khi có gì thay đổi)
# ---------------------------------------------------------------------------
STORE_INFO = (
    "Tên siêu thị: BHX_STR_CLD - Thửa 1289 An Nghiệp\n"
    "Mã siêu thị: 8363 | Tỉnh/thành: Cần Thơ\n"
    "Giờ hoạt động: 5h30 - 21h\n"
    "Quản lý: Trương Văn Quí (chức danh Quản Lý Thực Tập 1), mã nhân viên 227216"
)

NHAN_SU = (
    "- Lữ Thị Hà Mi (sinh 2000, vào làm từ 16/11/2019 - thâm niên cao nhất): "
    "luôn đứng vị trí Thu ngân khi phân line. Làm việc chịu khó, siêng năng, "
    "chịu khó học hỏi. Tính cách hiền, dễ nhờ vả.\n"
    "- Hồng Thị Nhựt Thi (sinh 2003, vào làm từ 11/04/2023): làm việc nhanh, "
    "gọn, đẹp, kỹ tính. Tính tình nóng, hay la mắng. Được đánh giá có tiềm "
    "năng lên làm quản lý.\n"
    "- Nguyễn Văn Sang (sinh 1999, vào làm từ 26/07/2022): nhân sự đa năng, "
    "đứng được mọi vị trí (thu ngân, kho...), không đứng vị trí FRESH khi "
    "phân line. Chịu khó, làm việc rất nhanh. Tính cách cọc cằn nhưng vui vẻ.\n"
    "- Nguyễn Kim Quyên (sinh 1991, vào làm từ 29/04/2021 - lão làng, lớn "
    "tuổi nhất nhóm nhân viên): ưu tiên đứng Thu ngân. Làm việc nhanh, kỹ "
    "tính, gọn gàng, ngăn nắp. Tính cách hiền lành, vui vẻ, hòa đồng.\n"
    "- Tăng Thị Ánh (sinh 2004, vào làm từ 12/06/2024): luôn hoàn thành tốt "
    "công việc siêu thị và việc được giao, hỗ trợ Quí khi vắng mặt. Làm việc "
    "nhanh, gọn, đẹp.\n"
    "- Trần Nhựt Linh (sinh 2005, vào làm từ 02/12/2024 - nhỏ tuổi nhất siêu "
    "thị): chủ yếu đứng kho, làm việc tạm ổn, chưa biết nhiều việc nhưng "
    "việc được giao luôn hoàn tất.\n"
    "- Diệp Sam Son (sinh 1997, vào làm từ 07/06/2024): vị trí tiếp đón "
    "khách hàng (TĐKH), thực hiện giao hàng online, đôi khi hỗ trợ nhân "
    "viên khác. Tính cách vui vẻ, hòa đồng, nhưng thường bị đồng nghiệp "
    "nhận xét là lười và ỷ lại vào công sức người khác.\n"
    "Lưu ý: việc phân nhân viên đại diện xử lý công việc thường ngày do "
    "chính Quí tự quyết định. Tất cả nhân viên đều có thể hỗ trợ tốt khi "
    "Quí vắng mặt, không chỉ riêng 1 người."
)


def _tien(x):
    try:
        return f"{x:,.0f}".replace(",", ".") + " đ"
    except Exception:
        return "0 đ"


def _an_toan(ham, mac_dinh):
    """Gọi 1 hàm lấy data, nếu lỗi thì trả về giá trị mặc định thay vì làm
    sập cả tính năng trả lời tự do."""
    try:
        return ham()
    except Exception:
        return mac_dinh


def _context_doanh_thu():
    rows = storage.get_all_revenue_rows()
    if not rows:
        return "Chưa có dữ liệu doanh thu nào được lưu."
    theo_thang = defaultdict(float)
    for r in rows:
        ngay = r.get("ngay") or ""
        if len(ngay) < 7:
            continue
        thang = ngay[:7]
        theo_thang[thang] += (r.get("dt_offline") or 0) + (r.get("dt_online") or 0)
    dong = [f"  Tháng {t}: {_tien(v)}" for t, v in sorted(theo_thang.items())]
    latest, latest_records, prev, prev_records = storage.get_latest_and_previous()
    if not latest:
        return "Chưa có dữ liệu doanh thu nào được lưu."
    tong_latest = sum((r.get("dt_offline") or 0) + (r.get("dt_online") or 0) for r in latest_records)
    ket_qua = [f"Ngày mới nhất có dữ liệu: {latest} — tổng doanh thu {_tien(tong_latest)}"]
    ket_qua.append("Doanh thu cộng dồn theo tháng đã lưu:")
    ket_qua.extend(dong)
    return "\n".join(ket_qua)


def _context_nganh_hang():
    ngay, ten_st, payload, gio = storage.get_latest_category_report()
    if not ngay or not payload:
        return "Chưa có dữ liệu ngành hàng (MTKM) nào được lưu."
    nam = payload.get("nam", {}) or {}
    banh = payload.get("banh_trung_thu", {}) or {}
    c2 = payload.get("c2", {}) or {}
    return (
        f"Báo cáo ngành hàng (MTKM) mới nhất — ngày {ngay}:\n"
        f"  Nấm: {_tien(nam.get('doanh_thu', 0))}\n"
        f"  Bánh trung thu: {banh.get('tong_sl', 0):.0f} cái\n"
        f"  C2: {c2.get('tong_chai', 0):.0f} chai"
    )


def _context_thuong():
    ten_st, payload, gio = storage.get_latest_thuong_report()
    if not ten_st or not payload:
        return "Chưa có dữ liệu thưởng (FRESH + FMCG) nào được lưu."
    return (
        f"Báo cáo thưởng ngành hàng mới nhất ({payload.get('ngay_bat_dau', '?')} "
        f"đến {payload.get('ngay_ket_thuc', '?')}): "
        f"tổng thưởng dự kiến {_tien(payload.get('tong_thuong_du_kien', 0))}"
    )


FRESH_SO_NGAY_TOI_DA_TRONG_CONTEXT = 45  # tránh context phình quá to nếu data lâu ngày


def _context_fresh():
    """QUAN TRỌNG: trước đây hàm này chỉ đưa ra 1 con số TỔNG CỘNG DỒN cả
    giai đoạn -> khiến AI trả lời tự do KHÔNG tách được số liệu từng ngày
    riêng lẻ (dù lệnh cố định "hủy mmkk <ngày>" vẫn tách đúng, vì lệnh đó
    đọc thẳng từ storage.get_fresh_records_by_date(), không qua hàm này).
    Đã sửa: liệt kê số liệu THEO TỪNG NGÀY (dùng đúng hàm
    get_fresh_records_by_date() cho mỗi ngày — hàm đã được lệnh "hủy mmkk"
    dùng và test đúng từ trước), để AI trả lời tự do cũng trả lời được câu
    hỏi về đúng 1 ngày cụ thể, không chỉ số cộng dồn."""
    dates = storage.get_fresh_distinct_dates()
    if not dates:
        return "Chưa có dữ liệu Hủy tồn + MMKK (Fresh) nào được lưu."
    ngay_sorted = sorted(dates)
    bi_cat_bot = len(ngay_sorted) > FRESH_SO_NGAY_TOI_DA_TRONG_CONTEXT
    ngay_hien = ngay_sorted[-FRESH_SO_NGAY_TOI_DA_TRONG_CONTEXT:] if bi_cat_bot else ngay_sorted
    dong = [
        f"Dữ liệu Hủy tồn + MMKK đã lưu cho các ngày: {ngay_sorted[0]} đến {ngay_sorted[-1]}"
        + (f" (chỉ liệt kê chi tiết {FRESH_SO_NGAY_TOI_DA_TRONG_CONTEXT} ngày gần nhất bên dưới)" if bi_cat_bot else "")
        + ". Số liệu TỪNG NGÀY (đơn vị: kg cho Rau củ/Trái cây/Thịt/Thủy hải sản, hộp cho Trứng):"
    ]
    tong_theo_nhom = defaultdict(lambda: {"huy": 0.0, "mmkk": 0.0})
    for ngay_str in ngay_hien:
        rows = storage.get_fresh_records_by_date(ngay_str)
        theo_nhom = defaultdict(lambda: {"huy": 0.0, "mmkk": 0.0})
        for r in rows:
            nh = r.get("nganh_hang") or "Khác"
            huy_val = r.get("sl_huy") or 0
            mmkk_val = r.get("sl_mmkk") or 0
            theo_nhom[nh]["huy"] += huy_val
            theo_nhom[nh]["mmkk"] += mmkk_val
            tong_theo_nhom[nh]["huy"] += huy_val
            tong_theo_nhom[nh]["mmkk"] += mmkk_val
        chi_tiet = "; ".join(
            f"{nh}: hủy {s['huy']:.1f}, mmkk {s['mmkk']:.1f}"
            for nh, s in sorted(theo_nhom.items())
        )
        dong.append(f"  Ngày {ngay_str}: {chi_tiet if chi_tiet else '(không có dữ liệu)'}")
    dong.append(f"Tổng cộng dồn {ngay_sorted[0]} đến {ngay_sorted[-1]} (theo ngành hàng):")
    for nh, s in sorted(tong_theo_nhom.items()):
        dong.append(f"  {nh}: hủy tồn {s['huy']:.1f}, mất mát kiểm kê {s['mmkk']:.1f}")
    dong.append(
        "(Muốn xem bảng chi tiết từng sản phẩm đúng 1 ngày, hoặc phân tích "
        "nguyên nhân, anh tag bot + gõ \"hủy mmkk <ngày>\" hoặc \"phân tích số "
        "liệu\" như bình thường.)"
    )
    return "\n".join(dong)


def _build_system_prompt():
    doanh_thu = _an_toan(_context_doanh_thu, "Chưa có dữ liệu doanh thu.")
    nganh_hang = _an_toan(_context_nganh_hang, "Chưa có dữ liệu ngành hàng.")
    thuong = _an_toan(_context_thuong, "Chưa có dữ liệu thưởng.")
    fresh = _an_toan(_context_fresh, "Chưa có dữ liệu Hủy tồn + MMKK.")

    return f"""Bạn là TROLY, trợ lý ảo hỗ trợ anh Quí — quản lý 1 cửa hàng Bách Hóa Xanh.

THÔNG TIN CỬA HÀNG:
{STORE_INFO}

NHÂN SỰ (7 bạn nhân viên dưới quyền anh Quí):
{NHAN_SU}

DỮ LIỆU DOANH THU ĐÃ LƯU:
{doanh_thu}

DỮ LIỆU NGÀNH HÀNG (MTKM) ĐÃ LƯU:
{nganh_hang}

DỮ LIỆU THƯỞNG ĐÃ LƯU:
{thuong}

DỮ LIỆU HỦY TỒN + MMKK (FRESH) ĐÃ LƯU:
{fresh}

QUY TẮC TRẢ LỜI (bắt buộc tuân thủ):
1. CHỈ được trả lời dựa trên đúng những dữ liệu liệt kê ở trên. Câu hỏi nào
   nằm ngoài phạm vi dữ liệu đã có (vd hỏi về giá vốn/lợi nhuận theo ngành
   hàng, hỏi về danh sách khuyến mãi Fresh theo tần suất, hỏi số liệu ngày
   chưa được gửi...) thì trả lời thẳng "Em chưa có dữ liệu này, anh cung
   cấp thêm giúp em" — TUYỆT ĐỐI không suy đoán hay bịa số liệu.
2. Xưng "em", gọi người hỏi là "anh". Trả lời ngắn gọn, đi thẳng vào việc,
   không dài dòng, không rào đón.
3. Nếu anh hỏi những gì đã có lệnh báo cáo dạng thẻ đẹp sẵn (doanh thu ->
   gõ "DT", ngành hàng -> "MTKM", thưởng -> "TD"/"THƯỞNG", doanh thu theo
   tháng -> "DTDK"), có thể trả lời gọn con số rồi gợi ý anh gõ đúng lệnh
   đó để xem bản đầy đủ, đẹp hơn.
"""


def hoi_ai(cau_hoi, target_id=None):
    """Gửi câu hỏi tự do lên Claude API, trả về chuỗi text để bot reply lại
    trong LINE. Không bao giờ raise exception ra ngoài — luôn trả về 1
    chuỗi (kể cả khi lỗi), để app.py chỉ cần reply_text() thẳng kết quả.

    target_id: truyền vào (group_id hoặc user_id) để bot NHỚ lịch sử hội
    thoại riêng của đúng đoạn chat đó (không bắt buộc — không truyền thì
    vẫn trả lời được, chỉ là không nhớ ngữ cảnh câu trước)."""
    if not ANTHROPIC_API_KEY:
        return "Chưa cấu hình được AI (thiếu ANTHROPIC_API_KEY trên Railway), anh báo lại giúp em."
    if not cau_hoi or not cau_hoi.strip():
        return "Anh hỏi em gì đó cụ thể giúp em nhé."
    cau_hoi = cau_hoi.strip()
    try:
        system_prompt = _build_system_prompt()
        messages = list(_lay_lich_su(target_id))
        messages.append({"role": "user", "content": cau_hoi})
        body = {
            "model": ANTHROPIC_MODEL,
            "max_tokens": 700,
            "system": system_prompt,
            "messages": messages,
        }
        resp = requests.post(
            ANTHROPIC_URL,
            headers={
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json=body,
            timeout=25,
        )
        if resp.status_code >= 300:
            print("Loi goi Claude API:", resp.status_code, resp.text)
            return "Em gọi AI bị lỗi, thử lại sau giúp em nhé."
        data = resp.json()
        parts = data.get("content") or []
        text = "".join(p.get("text", "") for p in parts if p.get("type") == "text")
        text = text.strip()
        ket_qua = text or "Em chưa nghĩ ra câu trả lời, anh hỏi lại giúp em."
        _luu_luot_chat(target_id, "user", cau_hoi)
        _luu_luot_chat(target_id, "assistant", ket_qua)
        return ket_qua
    except requests.exceptions.RequestException:
        return "Em không kết nối được tới AI lúc này, thử lại sau giúp em nhé."
    except Exception:
        import traceback
        traceback.print_exc()
        return "Có lỗi khi em xử lý câu hỏi này, thử lại giúp em nhé."


# ---------------------------------------------------------------------------
# ĐỌC ẢNH + PHÂN TÍCH SỐ LIỆU TỪ ẢNH (MỚI - dùng Claude Vision)
# ---------------------------------------------------------------------------
def _build_system_prompt_doc_anh():
    return f"""Bạn là TROLY, trợ lý ảo hỗ trợ anh Quí — quản lý 1 cửa hàng Bách Hóa Xanh.

THÔNG TIN CỬA HÀNG:
{STORE_INFO}

NHÂN SỰ (7 bạn nhân viên dưới quyền anh Quí):
{NHAN_SU}

NHIỆM VỤ: anh vừa gửi 1 tấm ảnh chụp số liệu/bảng dữ liệu (báo cáo bán hàng,
hủy tồn, doanh thu, hoặc số liệu khác tương tự). Bạn cần:
1. Đọc chính xác nội dung số liệu trong ảnh. Nếu có chỗ mờ/không đọc được
   thì nói rõ phần đó không đọc được, TUYỆT ĐỐI không suy đoán hay bịa số.
2. Phân tích: chỉ ra ưu điểm, nhược điểm/vấn đề, nguyên nhân cụ thể (sản
   phẩm/ngành hàng nào gây vấn đề), và đề xuất cách khắc phục/kiểm soát —
   phải thẳng, đúng thực tế, không nói giảm nói tránh.
3. Xưng "em", gọi người hỏi là "anh" (trừ khi có cơ sở rõ ràng người hỏi là
   1 trong 7 bạn nhân viên ở trên thì gọi đúng tên bạn đó). Trả lời ngắn
   gọn, đi thẳng vào việc, có thể dùng gạch đầu dòng cho dễ đọc trong LINE.
"""


def _build_system_prompt_phan_tich_du_lieu():
    return f"""Bạn là TROLY, trợ lý ảo hỗ trợ anh Quí — quản lý 1 cửa hàng Bách Hóa Xanh.

THÔNG TIN CỬA HÀNG:
{STORE_INFO}

NHÂN SỰ (7 bạn nhân viên dưới quyền anh Quí):
{NHAN_SU}

NHIỆM VỤ: phân tích đúng số liệu báo cáo được cung cấp bên dưới (đã lấy sẵn
từ hệ thống, KHÔNG suy đoán/bịa thêm số liệu ngoài phạm vi đã cho). Chỉ ra
ưu điểm, nhược điểm/vấn đề, nguyên nhân cụ thể, và đề xuất cách khắc phục/
kiểm soát — phải thẳng, đúng thực tế, không nói giảm nói tránh, tư duy phân
tích sâu (không trả lời hời hợt, qua loa).
Xưng "em", gọi người hỏi là "anh". Trả lời ngắn gọn, đi thẳng vào việc, có
thể dùng gạch đầu dòng cho dễ đọc trong LINE.
"""


def phan_tich_du_lieu(tieu_de, noi_dung, target_id=None):
    """Phân tích 1 đoạn dữ liệu TEXT đã có sẵn trong hệ thống (vd báo cáo
    doanh thu/ngành hàng) bằng Claude. Không bao giờ raise ra ngoài.

    target_id: truyền vào để lưu kết quả phân tích này vào chung lịch sử
    hội thoại của đoạn chat đó — nhờ vậy sau đó anh hỏi tiếp qua hoi_ai()
    (câu hỏi tự do) vẫn hiểu đang nói về báo cáo nào."""
    if not ANTHROPIC_API_KEY:
        return "Chưa cấu hình được AI (thiếu ANTHROPIC_API_KEY trên Railway), anh báo lại giúp em."
    if not noi_dung or not noi_dung.strip():
        return "Chưa có dữ liệu để phân tích, anh gõ lệnh lấy báo cáo trước giúp em."
    try:
        system_prompt = _build_system_prompt_phan_tich_du_lieu()
        body = {
            "model": ANTHROPIC_MODEL,
            "max_tokens": 1000,
            "system": system_prompt,
            "messages": [{
                "role": "user",
                "content": f"Phân tích giúp em báo cáo \"{tieu_de}\" sau:\n\n{noi_dung}",
            }],
        }
        resp = requests.post(
            ANTHROPIC_URL,
            headers={
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json=body,
            timeout=30,
        )
        if resp.status_code >= 300:
            print("Loi goi Claude API (phan tich du lieu):", resp.status_code, resp.text)
            return "Em phân tích bị lỗi, thử lại sau giúp em nhé."
        data = resp.json()
        parts = data.get("content") or []
        text = "".join(p.get("text", "") for p in parts if p.get("type") == "text")
        text = text.strip()
        ket_qua = text or "Em chưa phân tích được, anh hỏi lại giúp em."
        _luu_luot_chat(target_id, "user", f"[Yêu cầu phân tích số liệu] {tieu_de}")
        _luu_luot_chat(target_id, "assistant", ket_qua)
        return ket_qua
    except requests.exceptions.RequestException:
        return "Em không kết nối được tới AI lúc này, thử lại sau giúp em nhé."
    except Exception:
        import traceback
        traceback.print_exc()
        return "Có lỗi khi em phân tích, thử lại giúp em nhé."


def phan_tich_doanh_thu(target_id=None):
    """Phân tích báo cáo DOANH THU đang có trong hệ thống (dùng khi anh gõ
    'DT' rồi tag bot + 'phân tích số liệu' ngay sau đó)."""
    noi_dung = _an_toan(_context_doanh_thu, "Chưa có dữ liệu doanh thu.")
    return phan_tich_du_lieu("Báo cáo doanh thu", noi_dung, target_id=target_id)


def phan_tich_nganh_hang(target_id=None):
    """Phân tích báo cáo NGÀNH HÀNG/MTKM đang có trong hệ thống (dùng khi anh
    gõ 'MTKM' rồi tag bot + 'phân tích số liệu' ngay sau đó)."""
    noi_dung = _an_toan(_context_nganh_hang, "Chưa có dữ liệu ngành hàng (MTKM).")
    return phan_tich_du_lieu("Báo cáo ngành hàng (MTKM)", noi_dung, target_id=target_id)


def phan_tich_anh(image_bytes, media_type="image/jpeg", target_id=None):
    """Gửi ảnh (bytes) cho Claude Vision để đọc + phân tích số liệu trong ảnh.
    Không bao giờ raise ra ngoài — luôn trả về 1 chuỗi text để bot reply
    thẳng trong LINE."""
    if not ANTHROPIC_API_KEY:
        return "Chưa cấu hình được AI (thiếu ANTHROPIC_API_KEY trên Railway), anh báo lại giúp em."
    if not image_bytes:
        return "Em chưa nhận được ảnh, anh gửi lại giúp em."
    try:
        image_b64 = base64.b64encode(image_bytes).decode("utf-8")
        system_prompt = _build_system_prompt_doc_anh()
        body = {
            "model": ANTHROPIC_MODEL,
            "max_tokens": 1200,
            "system": system_prompt,
            "messages": [{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {"type": "base64", "media_type": media_type, "data": image_b64},
                    },
                    {
                        "type": "text",
                        "text": "Đọc số liệu trong ảnh trên rồi phân tích giúp em theo đúng quy tắc đã nêu.",
                    },
                ],
            }],
        }
        resp = requests.post(
            ANTHROPIC_URL,
            headers={
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json=body,
            timeout=40,
        )
        if resp.status_code >= 300:
            print("Loi goi Claude Vision API:", resp.status_code, resp.text)
            return "Em đọc ảnh bị lỗi, thử lại sau giúp em nhé."
        data = resp.json()
        parts = data.get("content") or []
        text = "".join(p.get("text", "") for p in parts if p.get("type") == "text")
        text = text.strip()
        ket_qua = text or "Em chưa đọc được nội dung trong ảnh, anh gửi ảnh rõ hơn giúp em."
        _luu_luot_chat(target_id, "user", "[Anh gửi 1 ảnh chụp số liệu để phân tích]")
        _luu_luot_chat(target_id, "assistant", ket_qua)
        return ket_qua
    except requests.exceptions.RequestException:
        return "Em không kết nối được tới AI lúc này, thử lại sau giúp em nhé."
    except Exception:
        import traceback
        traceback.print_exc()
        return "Có lỗi khi em xử lý ảnh này, thử lại giúp em nhé."
