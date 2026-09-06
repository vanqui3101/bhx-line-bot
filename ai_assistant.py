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
import requests
from collections import defaultdict

import storage

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
ANTHROPIC_MODEL = "claude-sonnet-5"
ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"

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


def _context_fresh():
    dates = storage.get_fresh_distinct_dates()
    if not dates:
        return "Chưa có dữ liệu Hủy tồn + MMKK (Fresh) nào được lưu."
    ngay_den, ngay_tu = dates[0], dates[-1]
    rows = storage.get_fresh_records_range(ngay_tu, ngay_den)
    theo_nhom = defaultdict(lambda: {"nhap": 0.0, "xuat": 0.0, "huy": 0.0, "mmkk": 0.0})
    for r in rows:
        nh = r.get("nganh_hang") or "Khác"
        theo_nhom[nh]["nhap"] += r.get("sl_nhap") or 0
        theo_nhom[nh]["xuat"] += r.get("sl_xuat") or 0
        theo_nhom[nh]["huy"] += r.get("sl_huy") or 0
        theo_nhom[nh]["mmkk"] += r.get("sl_mmkk") or 0
    dong = [f"Dữ liệu Hủy tồn + MMKK đã lưu, từ {ngay_tu} đến {ngay_den} (tổng theo ngành hàng):"]
    for nh, s in sorted(theo_nhom.items()):
        dong.append(
            f"  {nh}: nhập {s['nhap']:.1f}, xuất {s['xuat']:.1f}, "
            f"hủy tồn {s['huy']:.1f}, mất mát kiểm kê {s['mmkk']:.1f}"
        )
    dong.append(
        "(Đây là số tổng cộng dồn cả giai đoạn. Muốn xem chi tiết đúng 1 ngày "
        "hoặc phân tích nguyên nhân, anh tag bot + gõ \"hủy mmkk <ngày>\" hoặc "
        "\"phân tích số liệu\" như bình thường.)"
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


def hoi_ai(cau_hoi):
    """Gửi câu hỏi tự do lên Claude API, trả về chuỗi text để bot reply lại
    trong LINE. Không bao giờ raise exception ra ngoài — luôn trả về 1
    chuỗi (kể cả khi lỗi), để app.py chỉ cần reply_text() thẳng kết quả."""
    if not ANTHROPIC_API_KEY:
        return "Chưa cấu hình được AI (thiếu ANTHROPIC_API_KEY trên Railway), anh báo lại giúp em."
    if not cau_hoi or not cau_hoi.strip():
        return "Anh hỏi em gì đó cụ thể giúp em nhé."
    try:
        system_prompt = _build_system_prompt()
        body = {
            "model": ANTHROPIC_MODEL,
            "max_tokens": 700,
            "system": system_prompt,
            "messages": [{"role": "user", "content": cau_hoi.strip()}],
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
        return text or "Em chưa nghĩ ra câu trả lời, anh hỏi lại giúp em."
    except requests.exceptions.RequestException:
        return "Em không kết nối được tới AI lúc này, thử lại sau giúp em nhé."
    except Exception:
        import traceback
        traceback.print_exc()
        return "Có lỗi khi em xử lý câu hỏi này, thử lại giúp em nhé."
