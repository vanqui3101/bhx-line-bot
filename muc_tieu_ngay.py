"""
muc_tieu_ngay.py - MỤC TIÊU NGÀY TỪNG BẠN + CHỐT SỐ CUỐI NGÀY (bản 01/11/2026).

Quy trình anh Quí chốt 01/11/2026:
  1. Sáng anh đăng bài phân line (THU NGÂN / FRESH / FMCG theo ca sáng – chiều,
     tên gõ thường hoặc tag đều được) -> bot trả ngay THẺ "MỤC TIÊU NGÀY TỪNG BẠN"
     (thẻ riêng, KHÔNG dính thẻ MTKM theo dõi trong ngày).
  2. Cuối ngày (chậm nhất 21h45) từng bạn nhắn số bán được vào nhóm, viết tự do
     (VD "Nay em bán trứng đc 8 vỉ Đm 1079 skho 1504 c2 6 lốc") -> bot đọc số,
     trả 1 dòng nhắc lại số đã hiểu. Nhắn lại thì mục nào có trong tin mới được
     cập nhật.
  3. Anh nạp file POS 76 chốt + gõ "CHỐT SỐ" -> thẻ: mục tiêu ngày cửa hàng so với
     POS, MT ngày mai, từng bạn (mục tiêu / báo cáo / % đạt), đối chiếu tổng báo
     cáo với POS (ai báo cao hơn POS), và 1 dòng nhận xét tự động.

Cách tính:
  - MT ngày cửa hàng = (Mục tiêu tháng - đã bán từ đầu tháng tới hết hôm qua)
                        / số ngày còn lại (tính cả hôm nay)  -> thiếu/dư tự chia đều.
  - Từng bạn: ca sáng 40% / ca chiều 60% (ca nào không có người thì ca kia 100%),
    chia đều cho số bạn cùng vai trò trong ca; đứng 2 ca thì cộng 2 phần.
    Mục đếm số nguyên (C2 lốc, 888 túi, trứng vỉ) chia nguyên, tổng đúng bằng MT ngày.
  - % đạt từng bạn = trung bình % các mục được giao (mỗi mục tối đa 200%).

MUỐN ĐỔI MỤC TIÊU / THÊM THÁNG: sửa MUC_TIEU_THANG bên dưới.
"""
import math
import re
import unicodedata
from datetime import datetime, timedelta

import storage

# ---------------------------------------------------------------------------
# CẤU HÌNH
# ---------------------------------------------------------------------------
TY_LE_CA = {"sang": 0.4, "chieu": 0.6}
TEN_NHAN_VIEN = ["Mi", "Quyên", "Sang", "Thi", "Ánh", "Linh", "Tiền", "Son"]

# Thứ tự cột trên thẻ
MUC_THU_NGAN = ["sk", "celano", "dm", "hmp", "c2", "ng888"]
MUC_FRESH = ["ca", "trung", "heo"]

MUC = {
    "sk": {"ten": "Sữa khô", "don_vi": "tr", "vai_tro": "thu_ngan",
           "nganh": ["Sữa - Thức uống bổ dưỡng các loại"]},
    # (02/11/2026) Thi đua Celano 4 SKU (chu kỳ 02–08/11) — tính theo LỐC 4 hộp
    "celano": {"ten": "Celano", "don_vi": "lốc", "vai_tro": "thu_ngan", "truong": "celano_loc", "nguyen": True},
    "dm": {"ten": "Đông mát", "don_vi": "tr", "vai_tro": "thu_ngan",
           "nganh": ["Sản Phẩm Từ Sữa - Bảo Quản Mát", "Kem các loại",
                     "Thực phẩm đông lạnh - Hàng mát các loại"]},
    # (02/11/2026) Hoá mỹ phẩm = chăm sóc cá nhân + chăm sóc nhà cửa (giống POS 835)
    "hmp": {"ten": "Hóa mỹ phẩm", "don_vi": "tr", "vai_tro": "thu_ngan",
            "nganh": ["Mỹ phẩm các loại", "Làm Đẹp", "Trang Thiết Bị Y Tế",
                      "Hóa phẩm các loại", "Dụng cụ gia đình", "Dụng cụ nhà bếp"]},
    "c2": {"ten": "C2", "don_vi": "lốc", "vai_tro": "thu_ngan", "truong": "c2_loc", "nguyen": True},
    "ng888": {"ten": "888", "don_vi": "túi", "vai_tro": "thu_ngan", "truong": "nuoc_giat_888.sl", "nguyen": True},
    "ca": {"ten": "Cá", "don_vi": "tr", "vai_tro": "fresh", "nganh": ["Thủy hải sản các loại"]},
    "trung": {"ten": "Trứng", "don_vi": "vỉ", "vai_tro": "fresh", "truong": "trung_vi", "nguyen": True},
    "heo": {"ten": "Heo C.P", "don_vi": "kg", "vai_tro": "fresh", "truong": "heo_cp_kg"},
}

# Mục tiêu tháng (chốt với anh Quí 01/11/2026)
#  sk: base T9 x120% (chia theo tỉ trọng 4 ngành ĐMSK 471tr) · dm = sữa mát + kem + đông lạnh · hmp: base T10 x115%
#  c2: anh Quí chốt 01/11: trung bình 30 lốc/ngày (thùng TH24 = 4 lốc, tính cả online) · 888: tối thiểu 200 túi
#  trung: 25 vỉ/ngày (sáng 10 + chiều 15) · heo: T9 728,7kg x110%
#  ca: (cá + trứng 75,953tr) - tiền 750 vỉ trứng (giá TB T10 ~25.039đ/vỉ = 18,779tr) = 57,174tr;
#      tuần 1 (01-07/11) cá + trứng phải chạy nhịp +15% (19,983tr) -> cá tuần 1 = 19,983 - 175 vỉ x 25.039 = 15,601tr
NHAN_NGAN = {"sk": "sữa khô", "celano": "Celano", "dm": "đông mát", "hmp": "hóa mỹ phẩm", "c2": "C2", "ng888": "888",
             "ca": "cá", "trung": "trứng", "heo": "heo C.P"}

MUC_TIEU_THANG = {
    "2026-11": {
        "sk": {"muc_tieu": 231_864_000},
        "hmp": {"muc_tieu": 280_100_000, "giao_tu": "2026-11-02"},   # base T10 243,56tr x 115%; chia từng bạn từ 02/11
        "dm": {"muc_tieu": 239_136_000},
        "c2": {"muc_tieu": 900},
        # Thi đua Celano KV anh Nam: chu kỳ 02–08/11, mục tiêu 6.015.000đ (= 1,5 lần doanh thu T10 4,01tr)
        # quy ra 183 lốc theo giá TB T10 32.866đ/lốc. Chỉ tính trong chu kỳ.
        "celano": {"chi_trong_giai_doan": True,
                   "giai_doan": [{"tu": "2026-11-02", "den": "2026-11-08", "muc_tieu": 183,
                                  "muc_tieu_tien": 6_015_000}]},
        "ng888": {"muc_tieu": 600},  # (02/11/2026) anh nâng lên 20 túi/ngày x 30 ngày (trước: 200 túi/tháng)
        "ca": {"muc_tieu": 57_174_000,
               "giai_doan": [{"tu": "2026-11-01", "den": "2026-11-07", "muc_tieu": 15_601_000}]},
        "trung": {"muc_tieu": 750},
        "heo": {"muc_tieu": 801.6},
    },
}


# ---------------------------------------------------------------------------
# TIỆN ÍCH
# ---------------------------------------------------------------------------
def _d(s):
    return datetime.strptime(s, "%Y-%m-%d").date()


def _s(d):
    return d.strftime("%Y-%m-%d")


def bo_dau(s):
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D").lower()


def fmt_so(v, don_vi, gon=False):
    """tr -> triệu 2 số lẻ · kg -> 1 số lẻ · lốc/túi/vỉ -> số nguyên."""
    v = v or 0
    if don_vi == "tr":
        return f"{v / 1_000_000:.2f}".replace(".", ",")
    if don_vi == "kg":
        return f"{v:.1f}".replace(".", ",")
    return f"{v:.0f}"


def pct_txt(p):
    """% hiển thị: chưa tới 100 thì làm tròn XUỐNG (99,8% -> 99%), tránh hiện "100%" mà chưa đạt."""
    if p is None:
        return "–"
    return f"{math.floor(p) if p < 100 else round(p):.0f}%"


def fmt_co_don_vi(v, don_vi):
    if don_vi == "tr":
        return fmt_so(v, "tr") + "tr"
    if don_vi == "kg":
        return fmt_so(v, "kg") + "kg"
    return f"{fmt_so(v, don_vi)} {don_vi}"


def nhan_ten(doan):
    """Dò 1 tên nhân viên trong 1 đoạn chữ ngắn (VD "* mi", "Quyên", "anh").
    So khớp có dấu trước, rồi mới không dấu (chỉ khi ra đúng 1 tên)."""
    t = unicodedata.normalize("NFC", str(doan or "")).strip().lower()
    t = re.sub(r"^[\s\*\-•+·.>]+|[\s.:;,!]+$", "", t)
    if not t:
        return None
    for ten in TEN_NHAN_VIEN:
        if t == ten.lower():
            return ten
    khop = [ten for ten in TEN_NHAN_VIEN if bo_dau(ten) == bo_dau(t)]
    return khop[0] if len(khop) == 1 else None


def nhan_nhieu_ten(dong):
    """Dòng chỉ gồm tên (1 hoặc nhiều tên cách nhau dấu phẩy / và / &).
    Trả về list tên nếu MỌI phần đều là tên hợp lệ, ngược lại []."""
    t = re.sub(r"^[\s\*\-•+·.>]+", "", str(dong or "")).strip()
    if not t or len(t) > 40:
        return []
    phan = [p for p in re.split(r"\s*(?:,|/|&|\+|\bvà\b|\bva\b)\s*", t) if p.strip()]
    ten = [nhan_ten(p) for p in phan]
    return ten if ten and all(ten) else []


# ---------------------------------------------------------------------------
# 1. MỤC TIÊU NGÀY CỬA HÀNG
# ---------------------------------------------------------------------------
def _gia_tri(payload, key):
    """Số của 1 mục trong 1 ngày từ payload POS 76 đã lưu. None nếu payload thiếu trường."""
    if payload is None:
        return None
    cfg = MUC[key]
    if "truong" in cfg:
        cur = payload
        for k in cfg["truong"].split("."):
            if not isinstance(cur, dict) or k not in cur:
                return None
            cur = cur[k]
        return float(cur or 0)
    ten_set = {t.lower() for t in cfg["nganh"]}
    items = (payload.get("nganh_hang") or {}).get("items") or []
    return sum(float(it.get("thanh_tien") or 0) for it in items if str(it.get("ten") or "").lower() in ten_set)


def _tien_trung(payload):
    if not payload:
        return None
    return float((payload.get("phan_tich") or {}).get("Trứng Các Loại", 0) or 0)


def _mt_con_lai(muc_tieu, theo_ngay, tu, den, ngay):
    """(Mục tiêu kỳ - đã bán từ 'tu' tới HẾT ngày trước 'ngay') / số ngày còn lại tính cả 'ngay'."""
    da_ban = sum(v for d, v in theo_ngay.items() if tu <= d < ngay)
    con = (den - ngay).days + 1
    return max(muc_tieu - da_ban, 0) / con if con > 0 else 0.0


def _mt_ngay_muc(cfg_thang, key, theo_ngay, ngay, dau_thang, cuoi_thang):
    c = cfg_thang[key]
    gd = next((g for g in c.get("giai_doan", []) if _d(g["tu"]) <= ngay <= _d(g["den"])), None)
    if gd:
        v = _mt_con_lai(gd["muc_tieu"], theo_ngay, _d(gd["tu"]), _d(gd["den"]), ngay)
    else:
        v = _mt_con_lai(c["muc_tieu"], theo_ngay, dau_thang, cuoi_thang, ngay)
    return math.ceil(v - 1e-9) if MUC[key].get("nguyen") else v


def tinh_cua_hang(ngay_str):
    """Mục tiêu ngày + số POS của cửa hàng tại 1 ngày.
    Trả về {"ngay", "ngay_thu", "so_ngay", "co_pos", "ngay_thieu", "muc": {key: {...}}} hoặc None nếu tháng chưa cài mục tiêu."""
    ngay = _d(ngay_str)
    cfg_thang = MUC_TIEU_THANG.get(ngay.strftime("%Y-%m"))
    if not cfg_thang:
        return None
    dau_thang = ngay.replace(day=1)
    cuoi_thang = (dau_thang.replace(day=28) + timedelta(days=4))
    cuoi_thang = cuoi_thang - timedelta(days=cuoi_thang.day)
    da_luu = storage.get_category_reports_range(_s(dau_thang), _s(ngay))
    p_hom_nay = da_luu.get(ngay_str)
    thieu = []
    muc = {}
    for key in MUC_THU_NGAN + MUC_FRESH:
        c = cfg_thang.get(key)
        if not c:
            continue
        gd_nay = next((g for g in c.get("giai_doan", []) if _d(g["tu"]) <= ngay <= _d(g["den"])), None)
        if c.get("chi_trong_giai_doan") and not gd_nay:
            continue  # thi đua theo chu kỳ: ngoài chu kỳ thì không hiện
        theo_ngay = {}
        d = dau_thang
        while d <= ngay:
            v = _gia_tri(da_luu.get(_s(d)), key)
            if v is None and d < ngay and _s(d) not in thieu:
                thieu.append(_s(d))
            theo_ngay[d] = v or 0.0
            d += timedelta(days=1)
        mt = _mt_ngay_muc(cfg_thang, key, theo_ngay, ngay, dau_thang, cuoi_thang)
        mai = ngay + timedelta(days=1)
        mt_mai = None
        if mai <= cuoi_thang and not (c.get("chi_trong_giai_doan") and mai > _d(gd_nay["den"])):
            theo_ngay_mai = dict(theo_ngay)
            mt_mai = _mt_ngay_muc(cfg_thang, key, theo_ngay_mai, mai, dau_thang, cuoi_thang)
        pos = _gia_tri(p_hom_nay, key) if p_hom_nay is not None else None
        muc[key] = {"mt_ngay": mt, "pos": pos, "mt_mai": mt_mai,
                    "pct": (pos / mt * 100) if (pos is not None and mt) else None}
        if gd_nay and gd_nay.get("muc_tieu_tien"):
            # thi đua tính doanh thu: lũy kế tiền trong chu kỳ tới hôm nay
            d = _d(gd_nay["tu"])
            tien = 0.0
            while d <= ngay:
                tien += float((da_luu.get(_s(d)) or {}).get(f"{key}_tien") or 0)
                d += timedelta(days=1)
            muc[key]["chu_ky"] = {"tu": gd_nay["tu"], "den": gd_nay["den"], "tien": tien,
                                  "muc_tieu_tien": gd_nay["muc_tieu_tien"]}
    if "trung" in muc:
        muc["trung"]["tien"] = _tien_trung(p_hom_nay)
    if "heo" in muc:
        muc["heo"]["da_mo"] = (p_hom_nay or {}).get("heo_cp_damo_kg")
    return {"ngay": ngay_str, "ngay_thu": ngay.day, "so_ngay": cuoi_thang.day,
            "co_pos": p_hom_nay is not None,
            # các mục đang áp dụng hôm đó (thi đua theo chu kỳ chỉ có trong chu kỳ)
            "thu_ngan": [k for k in MUC_THU_NGAN if k in muc], "fresh": [k for k in MUC_FRESH if k in muc],
            "ngay_thieu": [datetime.strptime(x, "%Y-%m-%d").strftime("%d/%m") for x in sorted(thieu)],
            "muc": muc}


# ---------------------------------------------------------------------------
# 2. CHIA MỤC TIÊU TỪNG BẠN THEO BÀI PHÂN LINE
# ---------------------------------------------------------------------------
def ds_theo_vai_tro(phan_line, vai_tro):
    """{"sang": [tên...], "chieu": [tên...]} của 1 vai trò (thu_ngan / fresh / fmcg) từ bài phân line đã lưu.
    Bài cũ (chỉ có user_id) thì đổi user_id -> tên đã đăng ký."""
    kq = {}
    for ca in ("sang", "chieu"):
        p = (phan_line or {}).get(ca) or {}
        ten = list(p.get(f"{vai_tro}_names") or [])
        if not ten:
            for uid in p.get(f"{vai_tro}_users") or []:
                t = storage.get_ten_ngan_tu_user_id(uid)
                if t and t not in ten:
                    ten.append(t)
        kq[ca] = ten
    return kq


def _chia_nguyen(tong, phan):
    """Chia số nguyên 'tong' theo tỉ lệ phan {ten: so_thuc} (phần dư lớn nhất được +1)."""
    nguyen = {k: math.floor(v + 1e-9) for k, v in phan.items()}
    con = int(round(tong)) - sum(nguyen.values())
    for k in sorted(phan, key=lambda k: -(phan[k] - nguyen[k]))[:max(con, 0)]:
        nguyen[k] += 1
    return nguyen


def chia_tung_ban(ch, phan_line):
    """ch: kết quả tinh_cua_hang(). Trả về {"thu_ngan": [...], "fresh": [...], "fmcg": {ca: [tên]}}.
    Mỗi bạn: {"ten", "ca": ["sang"/"chieu"], "mt": {key: số}}."""
    kq = {"thu_ngan": [], "fresh": [], "fmcg": ds_theo_vai_tro(phan_line, "fmcg")}
    for vai_tro, ds_muc in (("thu_ngan", ch["thu_ngan"]), ("fresh", ch["fresh"])):
        ds_ca = ds_theo_vai_tro(phan_line, vai_tro)
        co_ca = [ca for ca in ("sang", "chieu") if ds_ca[ca]]
        if not co_ca:
            continue
        ty_le = {ca: (TY_LE_CA[ca] if len(co_ca) == 2 else 1.0) for ca in co_ca}
        nguoi = {}
        for ca in co_ca:
            phan = ty_le[ca] / len(ds_ca[ca])
            for ten in ds_ca[ca]:
                n = nguoi.setdefault(ten, {"ten": ten, "ca": [], "he_so": 0.0, "mt": {}})
                if ca not in n["ca"]:
                    n["ca"].append(ca)
                n["he_so"] += phan
        cfg_thang = MUC_TIEU_THANG.get(ch["ngay"][:7], {})
        for key in ds_muc:
            giao_tu = (cfg_thang.get(key) or {}).get("giao_tu")
            if giao_tu and ch["ngay"] < giao_tu:
                continue  # mục mới thêm giữa tháng: chưa giao cho từng bạn ở các ngày trước
            mt = ch["muc"][key]["mt_ngay"]
            if MUC[key].get("nguyen"):
                chia = _chia_nguyen(mt, {t: n["he_so"] * mt for t, n in nguoi.items()})
                for t, n in nguoi.items():
                    n["mt"][key] = chia[t]
            else:
                for n in nguoi.values():
                    n["mt"][key] = mt * n["he_so"]
        kq[vai_tro] = list(nguoi.values())
    return kq


def ca_txt(ds_ca):
    if len(ds_ca) == 2:
        return "cả ngày"
    return "sáng" if ds_ca and ds_ca[0] == "sang" else "chiều"


# ---------------------------------------------------------------------------
# 3. ĐỌC TIN NHẮN BÁO CÁO CỦA CÁC BẠN (viết tự do)
# ---------------------------------------------------------------------------
# Từ khoá dò trên bản KHÔNG DẤU (giữ nguyên vị trí ký tự so với bản NFC)
_TU_KHOA = [
    ("celano", r"celano|(?<![a-z0-9])cel(?![a-z0-9])"),
    ("hmp", r"hoa\s*my\s*pham|hoa\s*pham|my\s*pham|(?<![a-z0-9])hmp(?![a-z0-9])"),
    ("sk", r"sua\s*kho|skho|s\.\s*kho|(?<![a-z0-9])sk(?![a-z0-9])"),
    ("dm", r"dong\s*mat|d\.\s*mat|(?<![a-z0-9])dm(?![a-z0-9])"),
    ("c2", r"(?<![a-z0-9])c\s?2(?![0-9])"),
    ("ng888", r"nuoc\s*giat(?:\s*888)?|(?<![0-9])888(?![0-9])"),
    ("trung", r"trung(?:\s+cac?\s+loai)?"),
    ("ca", r"thuy\s*hai\s*san|hai\s*san|(?<![a-z0-9])ths(?![a-z0-9])"),
    ("heo", r"(?:thit\s*)?heo(?:\s*c\.?\s*p)?"),
]
_CA_CO_DAU = re.compile(r"(?<!\w)cá(?!\w)")  # "cá" phải dò trên bản CÓ dấu (tránh "các", "ca sáng")
_DON_VI_DEM = r"loc|vi|hop|tui|kg|ky|ki|cay|thung|chai|khay|qua|g"
_SO = re.compile(
    r"(?P<a>\d+(?:[.,]\d+)?)\s*(?:tr|trieu|cu)(?![a-z])\s*(?P<b>\d+)?"     # 2tr2 · 1tr972 · 1,5tr
    r"|(?P<c>\d+(?:[.,]\d+)?)\s*(?:k|nghin|ngan)(?![a-z])"                    # 500k · 750 nghìn
    r"|(?P<d>\d+(?:[.,]\d+)*)"                                                # số trơn
    r"(?:\s*(?P<u>" + _DON_VI_DEM + r")(?![a-z]))?"
)
_DON_VI_HIEN = {"loc": "lốc", "vi": "vỉ", "hop": "hộp", "tui": "túi", "kg": "kg", "ky": "kg", "ki": "kg",
                "cay": "cây", "thung": "thùng", "chai": "chai", "khay": "khay", "qua": "quả", "g": "g"}
_TU_BAO_CAO = re.compile(r"(?<!\w)(bán|ban|bc|báo cáo|bao cao|b/c)(?!\w)")


def _so_tron(s):
    if re.fullmatch(r"\d{1,3}(?:[.,]\d{3})+", s):
        return float(re.sub(r"[.,]", "", s))
    return float(s.replace(",", "."))


def _doc_so(m):
    """-> (gia_tri, loai, don_vi): loai 'tien_tr' (đã ra triệu) | 'tron' | 'dem' (có đơn vị đếm)."""
    if m.group("a") is not None:
        a = _so_tron(m.group("a"))
        b = m.group("b")
        if b:
            a = float(f"{int(a)}.{b}")
        return a, "tien_tr", None
    if m.group("c") is not None:
        return _so_tron(m.group("c")) / 1000, "tien_tr", None
    v = _so_tron(m.group("d"))
    u = m.group("u")
    return v, ("dem" if u else "tron"), u


def _tien_tu_so_tron(v):
    """Số trơn -> triệu: >=100.000 là đồng · >=100 là nghìn (1079 = 1,079tr) · <100 là triệu."""
    if v >= 100_000:
        return v / 1_000_000
    if v >= 100:
        return v / 1000
    return v


def doc_bao_cao(text, che_doan=None):
    """Đọc 1 tin nhắn báo cáo. Trả về (items, ghi) — items {key: số}:
    tiền (sk/dm/ca) đơn vị ĐỒNG; c2 lốc; 888 túi; trứng vỉ; heo kg. ghi: list lưu ý (số không dùng được)."""
    nfc = unicodedata.normalize("NFC", text or "").lower()
    t = bo_dau(nfc)
    if len(t) != len(nfc):  # an toàn: phải giữ nguyên độ dài để khớp vị trí
        t = "".join(bo_dau(ch) or " " for ch in nfc)
    kw = []
    da_che = [False] * len(t)
    # che phần tag @tên (LINE cho biết vị trí) — trong tên tag có mã NV, không phải số bán
    for bd, kt in (che_doan or []):
        for i in range(max(bd, 0), min(kt, len(t))):
            da_che[i] = True
    t = "".join(" " if da_che[i] else ch for i, ch in enumerate(t))
    nfc = "".join(" " if da_che[i] else ch for i, ch in enumerate(nfc))
    da_che = [False] * len(t)
    for key, pat in _TU_KHOA:
        for m in re.finditer(pat, t):
            if any(da_che[m.start():m.end()]):
                continue
            kw.append((m.start(), m.end(), key))
            for i in range(m.start(), m.end()):
                da_che[i] = True
    for m in _CA_CO_DAU.finditer(nfc):
        if not any(da_che[m.start():m.end()]):
            kw.append((m.start(), m.end(), "ca"))
            for i in range(m.start(), m.end()):
                da_che[i] = True
    if not kw:
        return {}, []
    kw.sort()
    # (02/11/2026) che từ khoá bằng "#" (không phải dấu cách): "hmp 1tr c2 2 lốc" trước đây che "c2" thành
    # khoảng trắng -> "1tr    2" bị đọc thành 1tr2 (1,2tr) và C2 mất số
    che = "".join("#" if da_che[i] else ch for i, ch in enumerate(t))
    # che ngày / giờ / "ca 1" để không bị hiểu nhầm là số bán (VD "ngày 1/11", "21h30")
    che = re.sub(r"(?<![\d.,])\d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?(?![\d.,])"      # ngày 1/11, 01-11-2026
                 r"|(?<![\d.,])\d{1,2}(?:h\d{0,2}|:\d{2})(?![a-z0-9])"              # 21h30, 21:30, 8h
                 r"|(?<![a-z])(?:ca|lan)\s?\d(?!\d)"                                  # ca 1, lần 2
                 r"|[_\-]\s?\d{6}(?!\d)|(?<!\d)\d{6}(?=\s?[_\-])",                   # mã NV trong tag (-112006, _227216_)
                 lambda m: " " * len(m.group(0)), che)
    so = [(m.start(), m) for m in _SO.finditer(che) if m.group(0).strip()]
    if not so:
        return {}, []
    kieu_kn = kw[0][0] < so[0][0]  # từ khoá đứng trước số -> "đm 1079"; ngược lại "1tr4 đm"
    dung = set()

    def _sau(i):
        bd, kt = kw[i][1], (kw[i + 1][0] if i + 1 < len(kw) else 10 ** 9)
        return [j for j, (p, _m) in enumerate(so) if bd <= p < kt and j not in dung]

    def _truoc(i):
        bd, kt = (kw[i - 1][1] if i > 0 else -1), kw[i][0]
        return [j for j, (p, _m) in enumerate(so) if bd <= p < kt and j not in dung]

    cap = []
    for i in range(len(kw)):
        a, b = (_sau(i), _truoc(i)) if kieu_kn else (_truoc(i), _sau(i))
        chon = (a[0] if kieu_kn else a[-1]) if a else ((b[-1] if kieu_kn else b[0]) if b else None)
        if chon is not None:
            dung.add(chon)
            cap.append((kw[i][2], so[chon][1]))
    items, ghi = {}, []
    for key, m in cap:
        v, loai, u = _doc_so(m)
        don_vi = MUC[key]["don_vi"]
        gt = None
        if don_vi == "tr":
            if loai == "tien_tr":
                gt = v * 1_000_000
            elif loai == "tron":
                gt = _tien_tu_so_tron(v) * 1_000_000
            else:
                ghi.append(f"{'Cá (THS)' if key == 'ca' else MUC[key]['ten']} {fmt_so(v, 'x')} "
                           f"{_DON_VI_HIEN.get(u, u)} – cần báo bằng tiền")
        elif key == "heo":
            if loai != "tien_tr":
                gt = v / 1000 if u == "g" else v
        else:  # c2 lốc · 888 túi · trứng vỉ
            if loai == "tien_tr":
                ghi.append(f"{MUC[key]['ten']} cần báo theo {don_vi}")
            elif u == "thung" and key in ("c2", "ng888"):
                gt = v * 4
            elif key == "celano" and u in ("hop", "thung"):
                gt = v / 4 if u == "hop" else v * 12
            elif u == "chai" and key == "c2":
                gt = v / 6
            else:
                gt = v
        if gt is not None:
            items[key] = items.get(key, 0) + gt
    return items, ghi


_MUC_TIEU_RE = re.compile(r"m[ụu]c\s*ti[êe]u|ch[ỉi]\s*ti[êe]u|(?<!\w)nhận\s*(?:mt|line|ca)(?!\w)")


def la_tin_bao_cao(text, items=None, co_tag=False):
    """(02/11/2026) Chỉ coi là BÁO CÁO BÁN HÀNG khi tin có chữ "bán" / "BC" / "báo cáo" và có số các mục.
    Tin nhắc lại MỤC TIÊU (có chữ "mục tiêu", "chỉ tiêu"...) thì KHÔNG tính — tránh ghi nhầm mục tiêu
    thành số bán như trường hợp Thi ngày 01/11. (co_tag giữ cho tương thích, không còn dùng.)"""
    if not items:
        return False
    t = unicodedata.normalize("NFC", text or "").lower()
    if _MUC_TIEU_RE.search(t):
        return False
    return bool(_TU_BAO_CAO.search(t))


def tom_tat_bao_cao(ten, items, ghi=None):
    phan = []
    for key in MUC_THU_NGAN + MUC_FRESH:
        if key in items:
            ten_muc = NHAN_NGAN[key]
            phan.append(f"{ten_muc} {fmt_co_don_vi(items[key], MUC[key]['don_vi'])}")
    s = f"✅ {ten}: " + " · ".join(phan)
    if ghi:
        s += "\n⚠ " + " · ".join(ghi)
    return s


# ---------------------------------------------------------------------------
# 4. CHỐT SỐ CUỐI NGÀY
# ---------------------------------------------------------------------------
def _pct(bc, mt):
    if not mt:
        return None
    return min((bc or 0) / mt * 100, 200)


def tinh_chot_so(ngay_str):
    ch = tinh_cua_hang(ngay_str)
    if ch is None:
        return None
    phan_line = storage.get_phan_line(ngay_str)
    tb = chia_tung_ban(ch, phan_line)
    bao_cao = storage.get_bao_cao_ban_hang(ngay_str)
    kq = {"ch": ch, "co_phan_line": bool(phan_line), "fmcg": tb["fmcg"], "bang": {}, "chua_bao": [],
          "bao_vuot": [], "so_bao": 0, "so_nguoi": 0}
    ten_co_muc_tieu = set()
    for vai_tro, ds_muc in (("thu_ngan", ch["thu_ngan"]), ("fresh", ch["fresh"])):
        dong = []
        for n in tb[vai_tro]:
            ten_co_muc_tieu.add(n["ten"])
            bc = (bao_cao.get(n["ten"]) or {}).get("items") or {}
            bc_vt = {k: bc[k] for k in ds_muc if k in bc}
            pcts = [p for p in (_pct(bc_vt.get(k, 0), n["mt"].get(k)) for k in ds_muc) if p is not None]
            dong.append({"ten": n["ten"], "ca": ca_txt(n["ca"]), "mt": n["mt"], "bc": bc_vt,
                         "da_bao": n["ten"] in bao_cao, "gio": (bao_cao.get(n["ten"]) or {}).get("gio"),
                         "pct": sum(pcts) / len(pcts) if pcts and n["ten"] in bao_cao else None})
        # bạn không có trong phân line vai trò này nhưng có báo mục của vai trò này
        for ten, v in bao_cao.items():
            if any(d["ten"] == ten for d in dong):
                continue
            bc_vt = {k: v["items"][k] for k in ds_muc if k in (v.get("items") or {})}
            if bc_vt:
                dong.append({"ten": ten, "ca": "ngoài line", "mt": {}, "bc": bc_vt, "da_bao": True, "pct": None,
                             "gio": v.get("gio")})
        dong.sort(key=lambda d: (d["pct"] is None, -(d["pct"] or 0)))
        tong_bc = {k: sum((d["bc"].get(k) or 0) for d in dong) for k in ds_muc}
        co_tu_van = {}
        for k in ds_muc:
            pos = ch["muc"][k]["pos"]
            co_tu_van[k] = (tong_bc[k] / pos * 100) if pos else None
            nguoi_bao = [d for d in dong if (d["bc"].get(k) or 0) > 0]
            if pos is not None and nguoi_bao:
                # chừa sai số làm tròn: tiền 2% (tối thiểu 20.000đ), kg 2% (tối thiểu 0,5kg); đếm cái thì không chừa
                dung_sai = {"tr": max(pos * 0.02, 20_000), "kg": max(pos * 0.02, 0.5)}.get(MUC[k]["don_vi"], 0)
                if tong_bc[k] > pos + dung_sai:
                    mot_minh = [d["ten"] for d in nguoi_bao if d["bc"][k] > pos + dung_sai]
                    kq["bao_vuot"].append({"muc": k, "tong": tong_bc[k], "pos": pos,
                                           "ten": mot_minh or [d["ten"] for d in nguoi_bao],
                                           "chac": bool(mot_minh) or len(nguoi_bao) == 1})
        kq["bang"][vai_tro] = {"dong": dong, "tong_bc": tong_bc, "co_tu_van": co_tu_van}
    kq["so_nguoi"] = len(ten_co_muc_tieu)
    kq["chua_bao"] = sorted(t for t in ten_co_muc_tieu if t not in bao_cao)
    kq["so_bao"] = kq["so_nguoi"] - len(kq["chua_bao"])
    kq["nhan_xet"] = _nhan_xet(kq)
    return kq


def _nhan_xet(kq):
    ch = kq["ch"]
    cau = []
    if ch["co_pos"]:
        dat = [MUC[k]["ten"] for k in ch["thu_ngan"] + ch["fresh"] if (ch["muc"][k]["pct"] or 0) >= 100]
        thieu = sorted(((ch["muc"][k]["pct"] or 0, k) for k in ch["thu_ngan"] + ch["fresh"]
                        if (ch["muc"][k]["pct"] or 0) < 100))
        s = f"Cửa hàng đạt {len(dat)}/{len(ch['thu_ngan'] + ch['fresh'])} mục tiêu ngày"
        if thieu:
            s += " – hụt nhiều nhất " + ", ".join(f"{MUC[k]['ten']} {pct_txt(p)}" for p, k in thieu[:2])
        cau.append(s)
    else:
        cau.append("Chưa có file POS hôm nay nên chưa so được mục tiêu ngày cửa hàng")
    nguoi = [d for vt in ("thu_ngan", "fresh") for d in kq["bang"][vt]["dong"] if d["pct"] is not None]
    if nguoi:
        cao = max(nguoi, key=lambda d: d["pct"])
        thap = min(nguoi, key=lambda d: d["pct"])
        cau.append(f"cao nhất {cao['ten']} {pct_txt(cao['pct'])}")
        if len(nguoi) > 1:
            yeu = min(((_pct(thap["bc"].get(k, 0), thap["mt"].get(k)) or 0, k) for k in thap["mt"] if thap["mt"].get(k)),
                      default=None)
            s = f"thấp nhất {thap['ten']} {pct_txt(thap['pct'])}"
            if yeu:
                s += f" (yếu {NHAN_NGAN[yeu[1]]})"
            cau.append(s)
    if kq["bao_vuot"]:
        theo_ten = {}
        nhom = []
        for b in kq["bao_vuot"]:
            if b["chac"]:
                for t in b["ten"]:
                    theo_ten.setdefault(t, []).append(NHAN_NGAN[b["muc"]])
            else:
                nhom.append(NHAN_NGAN[b["muc"]])
        if theo_ten:
            cau.append("BC KHỐNG: " + ", ".join(f"{t} ({', '.join(v)})" for t, v in theo_ten.items()))
        if nhom:
            cau.append("có BC KHỐNG ở " + ", ".join(nhom) + " (tổng báo cao hơn thực tế) – cần kiểm tra lại")
    if kq["chua_bao"]:
        cau.append("chưa báo: " + ", ".join(kq["chua_bao"]))
    s = "; ".join(cau) + "."
    return s[0].upper() + s[1:]
