"""
mtkm_tracker.py - Tính LŨY KẾ + MỤC TIÊU NGÀY cho thẻ MTKM (bản 05/10/2026).

Thẻ MTKM mới có 3 khối:
  1) Doanh thu NẤM      - mục tiêu THÁNG (anh Nam giao), cộng dồn cả tháng
  2) THI ĐUA FMCG TUẦN   - mỗi tuần chỉ đo ĐÚNG 1 ngành, mục tiêu = mức M2 trong
                           file THI_DUA_FMCG_T10 của anh Nam, cộng dồn trong
                           tuần, hết ngày cuối tuần tự chuyển sang ngành tuần kế
  3) NƯỚC GIẶT 888 3,2kg - mục tiêu THÁNG (túi), cộng dồn cả tháng

Nguồn số liệu: bảng category_reports (mỗi ngày 1 dòng, lưu từ file POS 76
"Doanh thu chi tiết"). Nạp lại file của cùng 1 ngày -> GHI ĐÈ ngày đó, không
cộng trùng. Lần nạp sau 21h là số chốt của ngày.

Cách tính mục tiêu ngày (dồn phần thiếu sang các ngày sau):
    Mục tiêu hôm nay = (Mục tiêu kỳ - Lũy kế đến HẾT HÔM QUA) / Số ngày còn lại (tính cả hôm nay)
    Mai cần          = (Mục tiêu kỳ - Lũy kế đến HIỆN TẠI)   / Số ngày còn lại SAU hôm nay

Doanh thu dùng cột "Thành tiền phải thu khách hàng (chưa VAT)" — khớp với
cột "Doanh thu" (chưa VAT) mà file anh Nam dùng làm base.

MUỐN ĐỔI MỤC TIÊU / THÊM THÁNG MỚI: chỉ cần sửa 3 cấu hình bên dưới
(MUC_TIEU_THANG, TUAN_THI_DUA, SO_LIEU_NEN), không phải sửa chỗ khác.
"""
import math
from datetime import datetime, timedelta

import storage

# ---------------------------------------------------------------------------
# CẤU HÌNH MỤC TIÊU
# ---------------------------------------------------------------------------
# Mục tiêu THÁNG cho Nấm (đồng, chưa VAT) và Nước giặt 888 (túi), theo "YYYY-MM".
MUC_TIEU_THANG = {
    "2026-10": {"nam": 25_800_000, "ng888": 300},
}

# Tên các "Ngành hàng" (đúng chữ trong cột Ngành hàng của POS 76) cho từng ngành thi đua
NGANH_DMSK = [
    "Sản Phẩm Từ Sữa - Bảo Quản Mát",           # NH 1354
    "Kem các loại",                              # NH 1355
    "Sữa - Thức uống bổ dưỡng các loại",         # NH 1056
    "Thực phẩm đông lạnh - Hàng mát các loại",   # NH 990
]
NGANH_HPMP = [
    "Hóa phẩm các loại",      # Chăm sóc nhà cửa
    "Mỹ phẩm các loại",       # Chăm sóc cá nhân
    "Làm Đẹp",                # Chăm sóc cá nhân (sữa rửa mặt, dưỡng da...)
]
NGANH_BANH_KEO = [
    "Bánh kẹo - Trà - Cà phê - Bột Dinh Dưỡng các loại",  # NH 1054
]

# 3 tuần thi đua FMCG tháng 10 (siêu thị 8363). Mục tiêu = M2 (triệu -> đồng).
TUAN_THI_DUA = [
    {
        "ten": "TUẦN 2", "nganh": "Đông mát – Sữa – Kem", "ds_nganh": NGANH_DMSK,
        "tu": "2026-10-05", "den": "2026-10-11",
        "m1": 101_100_000, "m2": 105_900_000,
    },
    {
        "ten": "TUẦN 3", "nganh": "Hóa phẩm – Mỹ phẩm", "ds_nganh": NGANH_HPMP,
        "tu": "2026-10-12", "den": "2026-10-18",
        "m1": 62_000_000, "m2": 67_600_000,
    },
    {
        "ten": "TUẦN 4", "nganh": "Bánh kẹo", "ds_nganh": NGANH_BANH_KEO,
        "tu": "2026-10-19", "den": "2026-10-25",
        "m1": 47_800_000, "m2": 50_100_000,
    },
]

# Số liệu NỀN cho những ngày bot chưa có file (chỉ dùng khi trong máy CHƯA lưu
# ngày đó — nếu sau này anh nạp file ngày đó thì số trong file được ưu tiên).
# Đã đối chiếu từ file POS 76 thật ngày 05/10/2026.
SO_LIEU_NEN = {
    "2026-10-01": {"nam": 449_523, "ng888": 0},
    "2026-10-02": {"nam": 558_095, "ng888": 4},
    "2026-10-03": {"nam": 571_428, "ng888": 2},
    "2026-10-04": {"nam": 1_005_713, "ng888": 6},
}


# ---------------------------------------------------------------------------
# HÀM TÍNH
# ---------------------------------------------------------------------------
def _d(s):
    return datetime.strptime(s, "%Y-%m-%d").date()


def _s(d):
    return d.strftime("%Y-%m-%d")


def _gia_tri_nganh(payload, ds_nganh):
    items = (payload.get("nganh_hang") or {}).get("items") or []
    ten_set = set(ds_nganh)
    return sum(float(it.get("thanh_tien") or 0) for it in items if it.get("ten") in ten_set)


def _tinh_ky(muc_tieu, theo_ngay, tu, den, ngay_bc, la_so_luong=False):
    """theo_ngay: dict {date: giá trị}. Trả về dict số liệu của 1 kỳ (tháng/tuần)."""
    luy_ke_truoc = sum(v for d, v in theo_ngay.items() if tu <= d < ngay_bc)
    hom_nay = theo_ngay.get(ngay_bc, 0.0)
    luy_ke = luy_ke_truoc + hom_nay
    so_ngay_con_lai_ca_hom_nay = (den - ngay_bc).days + 1
    so_ngay_sau_hom_nay = (den - ngay_bc).days

    muc_tieu_hom_nay = max(muc_tieu - luy_ke_truoc, 0) / so_ngay_con_lai_ca_hom_nay
    con_lai = max(muc_tieu - luy_ke, 0)
    mai_can = (con_lai / so_ngay_sau_hom_nay) if so_ngay_sau_hom_nay > 0 else None

    if la_so_luong:
        muc_tieu_hom_nay = math.ceil(muc_tieu_hom_nay)
        mai_can = math.ceil(mai_can) if mai_can is not None else None

    return {
        "muc_tieu": muc_tieu,
        "hom_nay": hom_nay,
        "luy_ke": luy_ke,
        "muc_tieu_hom_nay": muc_tieu_hom_nay,
        "con_thieu_hom_nay": max(muc_tieu_hom_nay - hom_nay, 0),
        "con_lai": con_lai,
        "so_ngay_sau_hom_nay": so_ngay_sau_hom_nay,
        "mai_can": mai_can,
        "pct": (luy_ke / muc_tieu * 100) if muc_tieu else 0.0,
        "ngay_thu": (ngay_bc - tu).days + 1,
        "tong_so_ngay": (den - tu).days + 1,
    }


def tinh_mtkm(ngay_bc_str):
    """Tính toàn bộ số liệu cho thẻ MTKM tại ngày báo cáo ngay_bc_str
    ("YYYY-MM-DD" = ngày của file mới nhất anh nạp)."""
    ngay_bc = _d(ngay_bc_str)
    dau_thang = ngay_bc.replace(day=1)
    cuoi_thang = (dau_thang.replace(day=28) + timedelta(days=4))
    cuoi_thang = cuoi_thang - timedelta(days=cuoi_thang.day)

    da_luu = storage.get_category_reports_range(_s(dau_thang), _s(ngay_bc))

    nam_theo_ngay, ng888_theo_ngay, ngay_thieu = {}, {}, []
    d = dau_thang
    while d <= ngay_bc:
        key = _s(d)
        if key in da_luu:
            p = da_luu[key]
            nam_theo_ngay[d] = float((p.get("nam") or {}).get("doanh_thu") or 0)
            ng888_theo_ngay[d] = float((p.get("nuoc_giat_888") or {}).get("sl") or 0)
        elif key in SO_LIEU_NEN:
            nam_theo_ngay[d] = float(SO_LIEU_NEN[key]["nam"])
            ng888_theo_ngay[d] = float(SO_LIEU_NEN[key]["ng888"])
        elif d < ngay_bc:
            ngay_thieu.append(d.strftime("%d/%m"))
        d += timedelta(days=1)

    ket_qua = {
        "ngay": ngay_bc_str,
        "ngay_thieu": ngay_thieu,
        "nam": None,
        "ng888": None,
        "tuan": None,
        "tuan_ghi_chu": None,
    }

    mt = MUC_TIEU_THANG.get(ngay_bc.strftime("%Y-%m"))
    if mt:
        ket_qua["nam"] = _tinh_ky(mt["nam"], nam_theo_ngay, dau_thang, cuoi_thang, ngay_bc)
        ket_qua["ng888"] = _tinh_ky(mt["ng888"], ng888_theo_ngay, dau_thang, cuoi_thang, ngay_bc, la_so_luong=True)
        ket_qua["tu_thang"] = _s(dau_thang)
        ket_qua["den_thang"] = _s(cuoi_thang)

    # ---- Tuần thi đua đang chạy ----
    tuan = next((t for t in TUAN_THI_DUA if _d(t["tu"]) <= ngay_bc <= _d(t["den"])), None)
    if tuan:
        tu, den = _d(tuan["tu"]), _d(tuan["den"])
        da_luu_tuan = storage.get_category_reports_range(tuan["tu"], _s(ngay_bc))
        theo_ngay = {_d(k): _gia_tri_nganh(v, tuan["ds_nganh"]) for k, v in da_luu_tuan.items()}
        so = _tinh_ky(tuan["m2"], theo_ngay, tu, den, ngay_bc)
        if so["luy_ke"] >= tuan["m2"]:
            trang_thai = "ĐÃ ĐẠT M2"
        elif so["luy_ke"] >= tuan["m1"]:
            trang_thai = "ĐÃ ĐẠT M1"
        elif ngay_bc == den:
            trang_thai = "CHƯA ĐẠT M1"
        else:
            trang_thai = None
        thieu_tuan = []
        dd = tu
        while dd < ngay_bc:
            if dd not in theo_ngay:
                thieu_tuan.append(dd.strftime("%d/%m"))
            dd += timedelta(days=1)
        so.update({
            "ten": tuan["ten"], "nganh": tuan["nganh"],
            "tu": tuan["tu"], "den": tuan["den"],
            "m1": tuan["m1"], "m2": tuan["m2"],
            "trang_thai": trang_thai, "ngay_thieu": thieu_tuan,
        })
        ket_qua["tuan"] = so
    else:
        sap_toi = next((t for t in TUAN_THI_DUA if _d(t["tu"]) > ngay_bc), None)
        if sap_toi:
            ket_qua["tuan_ghi_chu"] = (
                f"{sap_toi['ten']} ({sap_toi['nganh']}) bắt đầu "
                f"{_d(sap_toi['tu']).strftime('%d/%m')}"
            )
        else:
            ket_qua["tuan_ghi_chu"] = "Đã hết các tuần thi đua đã cài đặt"
    return ket_qua


def tom_tat_text(kq):
    """Bản chữ ngắn gọn (dùng cho tin xác nhận khi nạp file + AI phân tích)."""
    def tr(x):
        return f"{x / 1_000_000:.2f}".replace(".", ",") + " tr"

    dong = []
    if kq.get("nam"):
        n = kq["nam"]
        dong.append(f"Nấm: hôm nay {tr(n['hom_nay'])} / cần {tr(n['muc_tieu_hom_nay'])} · "
                    f"lũy kế {tr(n['luy_ke'])} / {tr(n['muc_tieu'])} ({n['pct']:.0f}%)")
    if kq.get("tuan"):
        t = kq["tuan"]
        dong.append(f"{t['ten']} {t['nganh']}: hôm nay {tr(t['hom_nay'])} / cần {tr(t['muc_tieu_hom_nay'])} · "
                    f"lũy kế {tr(t['luy_ke'])} / M2 {tr(t['m2'])} ({t['pct']:.0f}%)")
    elif kq.get("tuan_ghi_chu"):
        dong.append(f"Thi đua tuần: {kq['tuan_ghi_chu']}")
    if kq.get("ng888"):
        g = kq["ng888"]
        dong.append(f"Nước giặt 888: hôm nay {g['hom_nay']:.0f} / cần {g['muc_tieu_hom_nay']:.0f} túi · "
                    f"lũy kế {g['luy_ke']:.0f} / {g['muc_tieu']:.0f} túi ({g['pct']:.0f}%)")
    if kq.get("ngay_thieu"):
        dong.append(f"⚠ Chưa có dữ liệu ngày: {', '.join(kq['ngay_thieu'])} (đang tính = 0)")
    return "\n".join(dong)
