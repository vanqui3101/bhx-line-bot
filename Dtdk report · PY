"""
dtdk_report.py - Báo cáo "DTDK" (Doanh Thu Dự Kiến).

Báo cáo RIÊNG, KHÔNG thay thế lệnh "DT" cũ (DT vẫn giữ nguyên để theo dõi
thường xuyên). DTDK dùng để xem khi cần, gồm:
  1) Tháng hiện tại: tổng đã có + trung bình/ngày + dự kiến cả tháng
  2) Bảng tổng doanh thu theo từng tháng đã qua (Offline + Online gộp, chưa VAT)
  3) Tiến độ so với Target năm (gốc) và Mục tiêu tăng thêm (+15%)

Dùng chung dữ liệu với báo cáo DT/MTKM/TD — không cần file riêng, chỉ cần đã
gửi file "doanh thu chi tiết" (doanh thu theo siêu thị) cho bot như trước giờ.

Cú pháp lệnh: gõ đúng "DTDK" trong nhóm (không cần tag bot, giống DT/MTKM/TD).
"""
import calendar
from datetime import datetime

import storage
from flex_builder import build_dtdk_flex_message

# Target năm 2026 (chốt qua chat) và mục tiêu tăng thêm +15% so với target gốc.
TARGET_NAM = 19_153_000_000
TARGET_TANG_THEM_PCT = 0.15


def _today_vn():
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("Asia/Ho_Chi_Minh")).date()
    except Exception:
        return datetime.now().date()


def build_dtdk(ten_st):
    """Trả về bubble dict cho báo cáo DTDK, hoặc None nếu chưa có dữ liệu."""
    rows = storage.get_all_revenue_rows()
    if not rows:
        return None

    # Cộng dồn theo ngày trước (phòng khi 1 ngày có nhiều dòng siêu thị).
    theo_ngay = {}
    for r in rows:
        ngay = r.get("ngay")
        if not ngay:
            continue
        tong = (r.get("dt_offline") or 0) + (r.get("dt_online") or 0)
        theo_ngay[ngay] = theo_ngay.get(ngay, 0) + tong

    if not theo_ngay:
        return None

    # Gom theo tháng (khoá "YYYY-MM").
    theo_thang = {}
    for ngay, tong in theo_ngay.items():
        key = ngay[:7]
        theo_thang.setdefault(key, []).append(tong)

    today = _today_vn()
    thang_hien_tai_key = today.strftime("%Y-%m")

    ds_ngay_hien_tai = theo_thang.get(thang_hien_tai_key, [])
    so_ngay_hien_tai = len(ds_ngay_hien_tai)
    tong_hien_tai = sum(ds_ngay_hien_tai)
    tb_ngay = tong_hien_tai / so_ngay_hien_tai if so_ngay_hien_tai else 0
    so_ngay_trong_thang = calendar.monthrange(today.year, today.month)[1]
    du_kien_thang = tb_ngay * so_ngay_trong_thang

    month_rows = []
    grand = 0.0
    for key in sorted(theo_thang.keys()):
        if key == thang_hien_tai_key:
            continue
        _y, m = key.split("-")
        tong_thang = sum(theo_thang[key])
        grand += tong_thang
        month_rows.append((f"Tháng {int(m)}", tong_thang))

    target_tang_them = TARGET_NAM * (1 + TARGET_TANG_THEM_PCT)

    return build_dtdk_flex_message(
        ten_st=ten_st,
        thang_hien_tai_label=f"Tháng {today.month} ({so_ngay_hien_tai}/{so_ngay_trong_thang} ngày)",
        tong_hien_tai=tong_hien_tai,
        tb_ngay=tb_ngay,
        du_kien_thang=du_kien_thang,
        month_rows=month_rows,
        grand=grand,
        target_nam=TARGET_NAM,
        target_tang_them=target_tang_them,
    )
