"""
mtkm_tracker.py - Tính LŨY KẾ + MỤC TIÊU NGÀY cho thẻ MTKM (bản 20/10/2026).

Thẻ MTKM dạng BẢNG DỌC, 2 phần (cùng kiểu cột Mục tiêu / DT hiện tại / % đạt):
  - HÔM NAY      : mục tiêu hôm nay (đã dồn phần thiếu các ngày trước), DT hôm nay
  - LŨY KẾ THÁNG : mục tiêu tháng (cố định), DT từ đầu kỳ đến hiện tại

Các dòng (chốt với anh Quí 20/10/2026):
  1) Sữa khô                 - mục tiêu 231,86tr  (phần của 471tr, chia theo tỉ lệ T9)
  2) Đông mát – SC – Kem     - mục tiêu 239,14tr  (phần của 471tr, chia theo tỉ lệ T9)
     Tổng SK + ĐM            - mục tiêu 471tr (anh Nam giao)
  3) Thủy hải sản + Trứng    - mục tiêu 105% doanh thu T9 (72,08tr) = 75,69tr
  4) Thịt nhập khẩu 15–31/10 - mục tiêu +10% so TB T7–T8 (2,891tr/ngày x 17 ngày x 1,1)

Nguồn số liệu: bảng category_reports (mỗi ngày 1 dòng, lưu từ file POS 76
"Doanh thu chi tiết"). Nạp lại file của cùng 1 ngày -> GHI ĐÈ ngày đó, không
cộng trùng. Lần nạp sau 21h là số chốt của ngày.

Cách tính mục tiêu ngày (dồn phần thiếu sang các ngày sau):
    Mục tiêu hôm nay = (Mục tiêu kỳ - Lũy kế đến HẾT HÔM QUA) / Số ngày còn lại (tính cả hôm nay)
    Mai cần          = (Mục tiêu kỳ - Lũy kế đến HIỆN TẠI)   / Số ngày còn lại SAU hôm nay

Doanh thu dùng cột "Thành tiền phải thu khách hàng (chưa VAT)".

MUỐN ĐỔI MỤC TIÊU / THÊM THÁNG MỚI: chỉ sửa MUC_TIEU_THANG (và SO_LIEU_NEN nếu cần).
"""
import math
from datetime import datetime, timedelta

import storage

# ---------------------------------------------------------------------------
# CẤU HÌNH
# ---------------------------------------------------------------------------
NGANH_SUA_KHO = ["Sữa - Thức uống bổ dưỡng các loại"]
NGANH_DONG_MAT = [
    "Sản Phẩm Từ Sữa - Bảo Quản Mát",
    "Kem các loại",
    "Thực phẩm đông lạnh - Hàng mát các loại",
]
NGANH_THUY_HAI_SAN = ["Thủy hải sản các loại"]
# Theo cột "Ngành hàng - Phân tích" (Trứng nằm trong ngành gốc "Rau củ các loại",
# Thịt nhập khẩu nằm trong "Thịt gia cầm gia súc các loại")
PT_TRUNG = ["Trứng Các Loại"]
PT_THIT_NK = ["Thịt Nhập Khẩu"]

# "nganh": khớp cột Ngành hàng; "pt": khớp cột Ngành hàng - Phân tích (không phân biệt hoa thường)
MUC_TIEU_THANG = {
    "2026-10": {
        "khoi": [
            {"key": "sk", "ten": "Sữa khô", "nganh": NGANH_SUA_KHO, "pt": [],
             "muc_tieu": 231_860_000, "tu": "2026-10-01", "den": "2026-10-31"},
            {"key": "dm", "ten": "Đông mát–SC–Kem", "nganh": NGANH_DONG_MAT, "pt": [],
             "muc_tieu": 239_140_000, "tu": "2026-10-01", "den": "2026-10-31"},
            {"key": "thst", "ten": "Thủy hải sản + Trứng", "nganh": NGANH_THUY_HAI_SAN, "pt": PT_TRUNG,
             "muc_tieu": 75_687_000, "tu": "2026-10-01", "den": "2026-10-31"},
            {"key": "nk", "ten": "Thịt NK 15–31/10", "nganh": [], "pt": PT_THIT_NK,
             "muc_tieu": 54_068_000, "tu": "2026-10-15", "den": "2026-10-31"},
        ],
        # Dòng tổng (đặt ngay sau 2 dòng thành phần)
        "tong": {"key": "tong_skdm", "ten": "Tổng SK + ĐM", "gom": ["sk", "dm"],
                 "muc_tieu": 471_000_000, "sau_key": "dm"},
    },
}

# Số liệu NỀN 01–19/10 (đối chiếu từ file POS 76 thật). Chỉ dùng khi ngày đó
# trong máy CHƯA có file nạp đủ cột "Ngành hàng - Phân tích" — nạp file mới
# thì số trong file được ưu tiên. Thứ tự: sk, dm, thst, nk (đồng).
_NEN = {
    "2026-10-01": (4_371_113, 5_198_425, 2_237_416, 2_947_000),
    "2026-10-02": (3_500_095, 6_796_394, 1_742_556, 1_663_652),
    "2026-10-03": (5_111_297, 6_034_628, 1_606_731, 1_660_500),
    "2026-10-04": (5_844_443, 7_598_251, 1_626_112, 1_907_704),
    "2026-10-05": (2_797_316, 6_715_146, 1_698_318, 1_817_573),
    "2026-10-06": (2_749_351, 4_178_752, 1_589_865, 2_514_494),
    "2026-10-07": (2_611_942, 4_752_993, 2_091_097, 2_465_404),
    "2026-10-08": (2_512_869, 5_844_815, 1_819_586, 2_623_785),
    "2026-10-09": (4_129_723, 7_158_584, 1_827_459, 3_528_500),
    "2026-10-10": (5_894_074, 9_395_701, 1_755_837, 2_509_960),
    "2026-10-11": (5_196_392, 8_688_981, 2_805_906, 3_794_620),
    "2026-10-12": (6_547_130, 7_072_077, 1_310_512, 2_472_746),
    "2026-10-13": (3_394_080, 5_962_243, 1_525_505, 1_807_087),
    "2026-10-14": (3_552_223, 7_673_245, 1_396_202, 2_546_461),
    "2026-10-15": (3_819_167, 5_943_382, 2_569_881, 3_121_302),
    "2026-10-16": (13_813_887, 6_793_198, 2_403_462, 3_491_453),
    "2026-10-17": (7_037_128, 8_463_612, 2_627_440, 3_895_764),
    "2026-10-18": (4_199_999, 6_649_449, 2_840_507, 4_212_626),
    "2026-10-19": (5_564_631, 6_478_009, 2_502_412, 3_438_117),
}
SO_LIEU_NEN = {d: dict(zip(("sk", "dm", "thst", "nk"), v)) for d, v in _NEN.items()}


# ---------------------------------------------------------------------------
# HÀM TÍNH
# ---------------------------------------------------------------------------
def _d(s):
    return datetime.strptime(s, "%Y-%m-%d").date()


def _s(d):
    return d.strftime("%Y-%m-%d")


def _gia_tri_khoi(payload, khoi):
    """Doanh thu 1 khối trong 1 ngày từ payload đã lưu. Trả None nếu payload
    thiếu cột cần thiết (file cũ chưa có 'phan_tich')."""
    tong = 0.0
    if khoi["nganh"]:
        ten_set = {t.lower() for t in khoi["nganh"]}
        items = (payload.get("nganh_hang") or {}).get("items") or []
        tong += sum(float(it.get("thanh_tien") or 0) for it in items
                    if str(it.get("ten") or "").lower() in ten_set)
    if khoi["pt"]:
        pt = payload.get("phan_tich")
        if pt is None:
            return None
        ten_set = {t.lower() for t in khoi["pt"]}
        tong += sum(float(v or 0) for k, v in pt.items() if str(k).lower() in ten_set)
    return tong


def _tinh_ky(muc_tieu, theo_ngay, tu, den, ngay_bc, la_so_luong=False):
    """theo_ngay: dict {date: giá trị}. Trả về dict số liệu của 1 kỳ."""
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

    tong_so_ngay = (den - tu).days + 1
    ngay_thu = (ngay_bc - tu).days + 1
    return {
        "muc_tieu": muc_tieu,
        "hom_nay": hom_nay,
        "luy_ke": luy_ke,
        "muc_tieu_hom_nay": muc_tieu_hom_nay,
        "pct_hom_nay": (hom_nay / muc_tieu_hom_nay * 100) if muc_tieu_hom_nay else 100.0,
        "con_thieu_hom_nay": max(muc_tieu_hom_nay - hom_nay, 0),
        "con_lai": con_lai,
        "so_ngay_sau_hom_nay": so_ngay_sau_hom_nay,
        "mai_can": mai_can,
        "pct": (luy_ke / muc_tieu * 100) if muc_tieu else 0.0,
        # % thời gian đã qua của kỳ — để biết lũy kế đang nhanh hay chậm
        "pct_tien_do": ngay_thu / tong_so_ngay * 100,
        "ngay_thu": ngay_thu,
        "tong_so_ngay": tong_so_ngay,
    }


def tinh_mtkm(ngay_bc_str):
    """Tính toàn bộ số liệu cho thẻ MTKM tại ngày báo cáo ngay_bc_str
    ("YYYY-MM-DD" = ngày của file mới nhất anh nạp).
    Trả về {"ngay", "dong": [ {key, ten, la_tong, ...số liệu _tinh_ky}, ... ],
            "ghi_chu": [..], "ngay_thieu": [..]}"""
    ngay_bc = _d(ngay_bc_str)
    cau_hinh = MUC_TIEU_THANG.get(ngay_bc.strftime("%Y-%m"))
    ket_qua = {"ngay": ngay_bc_str, "dong": [], "ghi_chu": [], "ngay_thieu": []}
    if not cau_hinh:
        ket_qua["ghi_chu"].append("Chưa cài mục tiêu cho tháng này")
        return ket_qua

    dau_thang = ngay_bc.replace(day=1)
    da_luu = storage.get_category_reports_range(_s(dau_thang), _s(ngay_bc))

    # Giá trị theo ngày của từng khối
    theo_ngay = {k["key"]: {} for k in cau_hinh["khoi"]}
    ngay_thieu = []
    d = dau_thang
    while d <= ngay_bc:
        key = _s(d)
        p = da_luu.get(key)
        nen = SO_LIEU_NEN.get(key)
        thieu_ngay = False
        for k in cau_hinh["khoi"]:
            v = _gia_tri_khoi(p, k) if p is not None else None
            if v is None and nen is not None:
                v = float(nen.get(k["key"], 0))
            if v is None:
                if p is None or k["pt"]:
                    thieu_ngay = True
                v = 0.0
            theo_ngay[k["key"]][d] = v
        if thieu_ngay and d < ngay_bc:
            ngay_thieu.append(d.strftime("%d/%m"))
        d += timedelta(days=1)
    ket_qua["ngay_thieu"] = ngay_thieu

    so_khoi = {}
    for k in cau_hinh["khoi"]:
        tu, den = _d(k["tu"]), _d(k["den"])
        if ngay_bc < tu:
            ket_qua["ghi_chu"].append(f"{k['ten']}: bắt đầu {tu.strftime('%d/%m')}")
            continue
        ngay_tinh = min(ngay_bc, den)
        so = _tinh_ky(k["muc_tieu"], theo_ngay[k["key"]], tu, den, ngay_tinh)
        so.update({"key": k["key"], "ten": k["ten"], "la_tong": False, "tu": k["tu"], "den": k["den"]})
        so_khoi[k["key"]] = so
        ket_qua["dong"].append(so)
        tong = cau_hinh.get("tong")
        if tong and tong["sau_key"] == k["key"]:
            gop = {}
            for kk in tong["gom"]:
                for dd, vv in theo_ngay[kk].items():
                    gop[dd] = gop.get(dd, 0.0) + vv
            st = _tinh_ky(tong["muc_tieu"], gop, dau_thang, _d(k["den"]), ngay_tinh)
            st.update({"key": tong["key"], "ten": tong["ten"], "la_tong": True, "tu": k["tu"], "den": k["den"]})
            ket_qua["dong"].append(st)
    return ket_qua


def tom_tat_text(kq):
    """Bản chữ ngắn gọn (dùng cho tin xác nhận khi nạp file + AI phân tích)."""
    def tr(x):
        return f"{x / 1_000_000:.2f}".replace(".", ",") + " tr"

    dong = []
    for s in kq.get("dong", []):
        dong.append(
            f"{s['ten']}: hôm nay {tr(s['hom_nay'])} / MT {tr(s['muc_tieu_hom_nay'])} ({s['pct_hom_nay']:.0f}%) · "
            f"lũy kế {tr(s['luy_ke'])} / {tr(s['muc_tieu'])} ({s['pct']:.1f}%, tiến độ cần {s['pct_tien_do']:.0f}%)"
            + (f" · mai cần {tr(s['mai_can'])}" if s.get("mai_can") is not None else "")
        )
    for g in kq.get("ghi_chu", []):
        dong.append(g)
    if kq.get("ngay_thieu"):
        dong.append(f"⚠ Chưa có dữ liệu ngày: {', '.join(kq['ngay_thieu'])} (đang tính = 0)")
    return "\n".join(dong)
