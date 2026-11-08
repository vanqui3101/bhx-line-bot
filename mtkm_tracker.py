"""
mtkm_tracker.py - Tính MỤC TIÊU NGÀY + LŨY KẾ cho thẻ MTKM (bản 01/11/2026).

Thẻ MTKM tháng 11 (chốt với anh Quí 01/11/2026) — 1 bảng gọn, mỗi dòng:
  MT ngày | Hôm nay | % ngày | % tháng
Các dòng:
  Sữa khô · Sữa mát · Kem · Đông lạnh – hàng mát · Tổng ĐMSK   (base T9 × 120% = 471tr)
  Cá + Trứng   (base T10 × 102%; riêng tuần 1 01–07/11 phải chạy nhịp +15%)
  Heo C.P (kg) (base T9 728,7kg × 110%, tính cả da mỡ)
  Nước giặt 888 (túi) (mục tiêu 20 túi/ngày = 600 túi/tháng, từ 02/11/2026)
Kèm phần MỤC TIÊU TỪNG BẠN: đọc bài phân line của ngày, chia MT ngày
  - Thu ngân: sữa khô / sữa mát / kem / đông lạnh / 888
  - Fresh:    cá + trứng / heo C.P
  theo tỉ lệ ca sáng 40% – ca chiều 60%, rồi chia đều cho số bạn trong ca.

Cách tính mục tiêu ngày (dồn phần thiếu sang các ngày sau):
    MT ngày = (Mục tiêu kỳ - Lũy kế đến HẾT HÔM QUA) / Số ngày còn lại (tính cả hôm nay)

Nguồn số liệu: bảng category_reports (mỗi ngày 1 dòng, lưu từ file POS 76
"Doanh thu chi tiết"). Doanh thu = cột "Thành tiền phải thu khách hàng (chưa VAT)".

MUỐN ĐỔI MỤC TIÊU / THÊM THÁNG MỚI: chỉ sửa MUC_TIEU_THANG bên dưới.
"""
import math
from datetime import datetime, timedelta

import storage

# ---------------------------------------------------------------------------
# CẤU HÌNH
# ---------------------------------------------------------------------------
NGANH_SUA_KHO = ["Sữa - Thức uống bổ dưỡng các loại"]
NGANH_SUA_MAT = ["Sản Phẩm Từ Sữa - Bảo Quản Mát"]
NGANH_KEM = ["Kem các loại"]
NGANH_DONG_LANH = ["Thực phẩm đông lạnh - Hàng mát các loại"]
NGANH_THUY_HAI_SAN = ["Thủy hải sản các loại"]
PT_TRUNG = ["Trứng Các Loại"]          # theo cột "Ngành hàng - Phân tích"

# Tỉ lệ chia mục tiêu ngày theo ca (anh Quí chốt 01/11/2026)
TY_LE_CA = {"sang": 0.4, "chieu": 0.6}

# Mỗi dòng: key, ten, nguon (nganh/pt hoặc truong trong payload), don_vi,
#   muc_tieu (cả kỳ), tu/den (kỳ), giai_doan (tuỳ chọn: mục tiêu riêng cho 1
#   đoạn ngày, dùng để tính MT ngày trong đoạn đó), cho (thu_ngan/fresh/None)
MUC_TIEU_THANG = {
    "2026-11": {
        "dong": [
            {"key": "sk", "ten": "Sữa khô", "nganh": NGANH_SUA_KHO, "don_vi": "tr",
             "muc_tieu": 231_864_000, "cho": "thu_ngan"},
            {"key": "sm", "ten": "Sữa mát", "nganh": NGANH_SUA_MAT, "don_vi": "tr",
             "muc_tieu": 104_808_000, "cho": "thu_ngan"},
            {"key": "kem", "ten": "Kem", "nganh": NGANH_KEM, "don_vi": "tr",
             "muc_tieu": 21_060_000, "cho": "thu_ngan"},
            {"key": "dl", "ten": "Đông lạnh", "nganh": NGANH_DONG_LANH, "don_vi": "tr",
             "muc_tieu": 113_268_000, "cho": "thu_ngan"},
            {"key": "dmsk", "ten": "Tổng ĐMSK", "gom": ["sk", "sm", "kem", "dl"], "don_vi": "tr",
             "muc_tieu": 471_000_000, "la_tong": True},
            # (06/11/2026) Bỏ Hóa mỹ phẩm khỏi thẻ MTKM (mục tiêu HMP từng bạn vẫn ở muc_tieu_ngay.py)
            # (06/11/2026) Tách Cá và Trứng thành 2 dòng: Cá tính tiền, Trứng tính vỉ (như mục tiêu từng bạn)
            # (06/11/2026) Cá: doanh thu cá T10 58,31tr x 102% = 59,48tr, bỏ mục tiêu riêng tuần 1.
            # MT ngày = (59,48tr - đã bán tới hết hôm qua) / số ngày còn lại tính cả hôm nay.
            {"key": "ca", "ten": "Cá (THS)", "nganh": NGANH_THUY_HAI_SAN, "don_vi": "tr",
             "muc_tieu": 59_480_000, "cho": "fresh"},
            {"key": "trung", "ten": "Trứng (vỉ)", "truong": "trung_vi", "don_vi": "vi",
             "muc_tieu": 750, "cho": "fresh"},
            {"key": "heo", "ten": "Heo C.P (kg)", "truong": "heo_cp_kg", "don_vi": "kg",
             "muc_tieu": 801.6, "cho": "fresh"},
            # (08/11/2026) Nấm: phải đạt 25tr cả tháng, hiện từ 09/11 (lũy kế tính từ 01/11)
            {"key": "nam", "ten": "Nấm", "truong": "nam.doanh_thu", "don_vi": "tr",
             "muc_tieu": 25_000_000, "cho": "thu_ngan", "ap_dung_tu": "2026-11-09"},
            # (02/11/2026) Nhóm đếm số lượng: Celano (sữa lốc, chỉ trong chu kỳ thi đua) · C2 · nước giặt 888
            {"key": "celano", "ten": "Celano (lốc)", "truong": "celano_loc", "don_vi": "loc", "muc_tieu": 0,
             "chi_trong_giai_doan": True,
             "giai_doan": [{"tu": "2026-11-02", "den": "2026-11-08", "muc_tieu": 183, "nhan": "Thi đua 02–08/11"}]},
            {"key": "c2", "ten": "C2 (lốc)", "truong": "c2_loc", "don_vi": "loc", "muc_tieu": 900},
            {"key": "ng888", "ten": "888 (túi)", "truong": "nuoc_giat_888.sl", "don_vi": "tui",
             "muc_tieu": 600, "cho": "thu_ngan"},  # (02/11/2026) 20 túi/ngày
            # (08/11/2026) Thi đua tuần 2: kem đánh răng P/S 2 mã bán hết tồn (18 + 30 = 48 tuýp)
            {"key": "kdr", "ten": "Kem đánh răng P/S", "truong": "kdr_ps_tuyp", "don_vi": "tuyp", "muc_tieu": 0,
             "cho": "thu_ngan", "chi_trong_giai_doan": True,
             "giai_doan": [{"tu": "2026-11-09", "den": "2026-11-15", "muc_tieu": 48, "nhan": "Thi đua 09–15/11"}]},
        ],
    },
}


# ---------------------------------------------------------------------------
# HÀM TÍNH
# ---------------------------------------------------------------------------
def _d(s):
    return datetime.strptime(s, "%Y-%m-%d").date()


def _s(d):
    return d.strftime("%Y-%m-%d")


def _gia_tri(payload, dong):
    """Giá trị 1 dòng trong 1 ngày từ payload đã lưu. None nếu payload thiếu dữ liệu cần."""
    if "truong" in dong:
        cur = payload
        for k in dong["truong"].split("."):
            if not isinstance(cur, dict) or k not in cur:
                return None
            cur = cur[k]
        return float(cur or 0)
    tong = 0.0
    if dong.get("nganh"):
        ten_set = {t.lower() for t in dong["nganh"]}
        items = (payload.get("nganh_hang") or {}).get("items") or []
        tong += sum(float(it.get("thanh_tien") or 0) for it in items
                    if str(it.get("ten") or "").lower() in ten_set)
    if dong.get("pt"):
        pt = payload.get("phan_tich")
        if pt is None:
            return None
        ten_set = {t.lower() for t in dong["pt"]}
        tong += sum(float(v or 0) for k, v in pt.items() if str(k).lower() in ten_set)
    return tong


def _mt_ngay(muc_tieu, theo_ngay, tu, den, ngay_bc):
    luy_ke_truoc = sum(v for d, v in theo_ngay.items() if tu <= d < ngay_bc)
    so_ngay_con_lai = (den - ngay_bc).days + 1
    return max(muc_tieu - luy_ke_truoc, 0) / so_ngay_con_lai


def tinh_mtkm(ngay_bc_str):
    """Tính số liệu thẻ MTKM tại ngày báo cáo (ngày của file mới nhất anh nạp).
    Trả về {"ngay", "dong": [...], "tung_ban": [...], "ghi_chu": [...], "ngay_thieu": [...],
            "ngay_thu", "so_ngay_thang", "pct_tien_do"}"""
    ngay_bc = _d(ngay_bc_str)
    cau_hinh = MUC_TIEU_THANG.get(ngay_bc.strftime("%Y-%m"))
    dau_thang = ngay_bc.replace(day=1)
    cuoi_thang = (dau_thang.replace(day=28) + timedelta(days=4))
    cuoi_thang = cuoi_thang - timedelta(days=cuoi_thang.day)
    so_ngay_thang = cuoi_thang.day
    ngay_thu = ngay_bc.day
    ket_qua = {"ngay": ngay_bc_str, "dong": [], "tung_ban": [], "ghi_chu": [], "ngay_thieu": [],
               "ngay_thu": ngay_thu, "so_ngay_thang": so_ngay_thang,
               "pct_tien_do": ngay_thu / so_ngay_thang * 100}
    if not cau_hinh:
        ket_qua["ghi_chu"].append("Chưa cài mục tiêu MTKM cho tháng này")
        return ket_qua

    da_luu = storage.get_category_reports_range(_s(dau_thang), _s(ngay_bc))
    theo_ngay = {}     # key -> {date: value}
    thieu = set()
    def _gd_cua(dong):
        return next((g for g in dong.get("giai_doan", []) if _d(g["tu"]) <= ngay_bc <= _d(g["den"])), None)

    ds_dong = [dg for dg in cau_hinh["dong"] if not (dg.get("chi_trong_giai_doan") and not _gd_cua(dg))
               and not (dg.get("ap_dung_tu") and ngay_bc < _d(dg["ap_dung_tu"]))]
    for dong in ds_dong:
        if dong.get("la_tong"):
            continue
        theo_ngay[dong["key"]] = {}
        d = _d(_gd_cua(dong)["tu"]) if dong.get("chi_trong_giai_doan") else dau_thang
        while d <= ngay_bc:
            p = da_luu.get(_s(d))
            v = _gia_tri(p, dong) if p is not None else None
            if v is None:
                if d < ngay_bc or p is None:
                    thieu.add(d)
                v = 0.0
            theo_ngay[dong["key"]][d] = v
            d += timedelta(days=1)
    ket_qua["ngay_thieu"] = [x.strftime("%d/%m") for x in sorted(thieu)]

    for dong in ds_dong:
        if dong.get("la_tong"):
            gop = {}
            for kk in dong["gom"]:
                for dd, vv in theo_ngay[kk].items():
                    gop[dd] = gop.get(dd, 0.0) + vv
            tn = gop
        else:
            tn = theo_ngay[dong["key"]]
        # mục tiêu ngày: nếu đang trong 1 giai đoạn riêng (VD tuần 1) thì tính theo giai đoạn đó
        gd = next((g for g in dong.get("giai_doan", []) if _d(g["tu"]) <= ngay_bc <= _d(g["den"])), None)
        tien_do_rieng = None
        if gd:
            mt_ngay = _mt_ngay(gd["muc_tieu"], tn, _d(gd["tu"]), _d(gd["den"]), ngay_bc)
            # (06/11/2026) Giai đoạn đẩy nhanh (Cá tuần 1) đã gần đạt -> MT ngày không thấp hơn nhịp tháng
            # (trước đây Cá 06/11 ra 0,22tr vì tuần 1 chỉ còn thiếu 0,3tr). Thi đua chu kỳ (Celano) giữ nguyên.
            if not dong.get("chi_trong_giai_doan") and dong.get("muc_tieu"):
                mt_ngay = max(mt_ngay, _mt_ngay(dong["muc_tieu"], tn, dau_thang, cuoi_thang, ngay_bc))
            nhan_gd = gd.get("nhan") or f"Tuần 1 ({_d(gd['tu']).strftime('%d')}–{_d(gd['den']).strftime('%d/%m')})"
            if dong.get("chi_trong_giai_doan"):
                # thi đua theo chu kỳ: % tính trên mục tiêu chu kỳ, tiến độ theo số ngày của chu kỳ
                so_ngay_gd = (_d(gd["den"]) - _d(gd["tu"])).days + 1
                tien_do_rieng = ((ngay_bc - _d(gd["tu"])).days + 1) / so_ngay_gd * 100
        else:
            mt_ngay = _mt_ngay(dong["muc_tieu"], tn, dau_thang, cuoi_thang, ngay_bc)
            nhan_gd = None
        if dong["don_vi"] in ("tui", "loc", "vi", "tuyp"):
            mt_ngay = math.ceil(mt_ngay - 1e-9)
        hom_nay = tn.get(ngay_bc, 0.0)
        luy_ke = sum(v for dd, v in tn.items() if dd <= ngay_bc)
        ket_qua["dong"].append({
            "key": dong["key"], "ten": dong["ten"], "don_vi": dong["don_vi"],
            "la_tong": dong.get("la_tong", False), "cho": dong.get("cho"),
            "muc_tieu": dong["muc_tieu"], "mt_ngay": mt_ngay, "hom_nay": hom_nay,
            "luy_ke": luy_ke, "nhan_gd": nhan_gd,
            "pct_ngay": (hom_nay / mt_ngay * 100) if mt_ngay else 100.0,
            "pct_thang": ((luy_ke / gd["muc_tieu"] * 100) if (tien_do_rieng is not None and gd["muc_tieu"])
                          else (luy_ke / dong["muc_tieu"] * 100) if dong["muc_tieu"] else 0.0),
            "tien_do_rieng": tien_do_rieng,
        })

    # (01/11/2026) Mục tiêu TỪNG BẠN đã tách sang thẻ riêng (muc_tieu_ngay.py — gửi khi anh đăng
    # phân line + lệnh CHỐT SỐ), thẻ MTKM chỉ còn theo dõi cửa hàng trong ngày.
    return ket_qua


def _ten(uid):
    try:
        return storage.get_ten_ngan_tu_user_id(uid) or "(chưa đăng ký tên)"
    except Exception:
        return "(chưa đăng ký tên)"


def chia_muc_tieu(dong_list, phan_line):
    """Chia MT ngày cho từng bạn theo bài phân line: ca sáng 40% – ca chiều 60%,
    chia đều trong ca. Ca nào không có người thì ca kia nhận 100%.
    Bạn có mặt cả 2 ca được cộng cả 2 phần. Trả về list theo thứ tự:
    thu ngân sáng → thu ngân chiều → fresh sáng → fresh chiều."""
    if not phan_line:
        return []
    ket_qua = []
    for vai_tro, nhan_vt in (("thu_ngan", "Thu ngân"), ("fresh", "Fresh")):
        ds_ca = {ca: list((phan_line.get(ca) or {}).get(f"{vai_tro}_users") or []) for ca in ("sang", "chieu")}
        co_ca = [ca for ca in ("sang", "chieu") if ds_ca[ca]]
        if not co_ca:
            continue
        ty_le = {ca: (TY_LE_CA[ca] if len(co_ca) == 2 else 1.0) for ca in co_ca}
        dong_cua_vt = [d for d in dong_list if d.get("cho") == vai_tro]
        theo_nguoi = {}   # uid -> {"ca": [...], "mt": {key: value}}
        for ca in co_ca:
            phan = ty_le[ca] / len(ds_ca[ca])
            for uid in ds_ca[ca]:
                ng = theo_nguoi.setdefault(uid, {"ca": [], "mt": {}})
                ng["ca"].append(ca)
                for d in dong_cua_vt:
                    ng["mt"][d["key"]] = ng["mt"].get(d["key"], 0.0) + d["mt_ngay"] * phan
        # Dòng tính theo TÚI (888): chia số nguyên cho đủ đúng MT ngày (phần dư lớn nhất được +1)
        for d in dong_cua_vt:
            if d["don_vi"] != "tui" or not theo_nguoi:
                continue
            tong = int(round(d["mt_ngay"]))
            phan_nguyen = {u: math.floor(ng["mt"][d["key"]] + 1e-9) for u, ng in theo_nguoi.items()}
            con = tong - sum(phan_nguyen.values())
            du = sorted(theo_nguoi, key=lambda u: -(theo_nguoi[u]["mt"][d["key"]] - phan_nguyen[u]))
            for u in du[:max(con, 0)]:
                phan_nguyen[u] += 1
            for u in theo_nguoi:
                theo_nguoi[u]["mt"][d["key"]] = phan_nguyen[u]
        for uid, ng in theo_nguoi.items():
            ca_txt = "cả ngày" if len(ng["ca"]) == 2 else ("sáng" if ng["ca"][0] == "sang" else "chiều")
            ket_qua.append({
                "ten": _ten(uid), "vai_tro": f"{nhan_vt} {ca_txt}", "loai": vai_tro,
                "thu_tu": (0 if vai_tro == "thu_ngan" else 1, 0 if ng["ca"][0] == "sang" else 1),
                "mt": ng["mt"],
            })
    ket_qua.sort(key=lambda x: x["thu_tu"])
    return ket_qua


def _fmt(v, don_vi):
    if don_vi == "tr":
        return f"{v / 1_000_000:.2f}".replace(".", ",") + "tr"
    if don_vi == "kg":
        return f"{v:.1f}".replace(".", ",") + "kg"
    return f"{v:.0f} " + {"loc": "lốc", "vi": "vỉ", "tuyp": "tuýp"}.get(don_vi, "túi")


def tom_tat_text(kq):
    """Bản chữ ngắn gọn (dùng cho tin xác nhận khi nạp file + AI phân tích)."""
    dong = []
    for s in kq.get("dong", []):
        if s.get("tien_do_rieng") is not None:
            dong.append(f"{s['ten']}: hôm nay {_fmt(s['hom_nay'], s['don_vi'])} / MT ngày {_fmt(s['mt_ngay'], s['don_vi'])}"
                        f" ({s['pct_ngay']:.0f}%) · lũy kế chu kỳ {_fmt(s['luy_ke'], s['don_vi'])}")
            continue
        dong.append(
            f"{s['ten']}: hôm nay {_fmt(s['hom_nay'], s['don_vi'])} / MT ngày {_fmt(s['mt_ngay'], s['don_vi'])}"
            f" ({s['pct_ngay']:.0f}%) · lũy kế {_fmt(s['luy_ke'], s['don_vi'])} / {_fmt(s['muc_tieu'], s['don_vi'])}"
            f" ({s['pct_thang']:.1f}%, tiến độ cần {kq['pct_tien_do']:.0f}%)"
        )
    for b in kq.get("tung_ban", []):
        nhan = {"sk": "Sữa khô", "sm": "Sữa mát", "kem": "Kem", "dl": "Đông lạnh",
                "ng888": "888", "ca_trung": "Cá + trứng", "heo": "Heo C.P"}
        mt = " · ".join(f"{nhan.get(k, k)} {_fmt(v, 'tui' if k == 'ng888' else ('kg' if k == 'heo' else 'tr'))}"
                        for k, v in b["mt"].items())
        dong.append(f"- {b['ten']} ({b['vai_tro']}): {mt}")
    for g in kq.get("ghi_chu", []):
        dong.append(g)
    if kq.get("ngay_thieu"):
        dong.append(f"⚠ Chưa có dữ liệu ngày: {', '.join(kq['ngay_thieu'])} (đang tính = 0)")
    return "\n".join(dong)
