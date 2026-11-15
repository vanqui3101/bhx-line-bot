"""
MTKM CHO SIÊU THỊ PHỤ (15/11/2026) — anh Quí làm giúp siêu thị bạn, TÁCH RIÊNG HOÀN TOÀN khỏi 8363.

- Hiện có: 14536 Mỹ Xuyên — chỉ đo NẤM (mục tiêu doanh thu nấm cả tháng 10,6tr).
- Nguồn: file POS 76 (Doanh thu chi tiết) của siêu thị đó; nấm = Nhóm hàng "Nấm Các Loại" (giống 8363),
  tiền = "Thành tiền phải thu khách hàng (chưa VAT)".
- Số lưu riêng ở bảng mtkm_st_phu (ma_st, ngay) — không ghi vào bảng MTKM/thưởng/chốt số/mục tiêu của 8363.
- Mục tiêu ngày lăn: (mục tiêu tháng − lũy kế tới hết hôm trước) ÷ số ngày còn lại (tính cả hôm nay)
  -> bán dư/thiếu hôm nay thì ngày sau tự giảm/tăng.
- Chia theo ca: sáng 40% / chiều 60%, mỗi ca 3 bạn chia đều (thu ngân và fresh cùng ca như nhau, không ghi tên).
- Lệnh: "MTKM 14536" / "MTKM MX" / "MTKM Mỹ Xuyên". Gửi file POS 76 của 14536 vào chat -> bot trả ngay bảng.
"""
import calendar
import re
import unicodedata
from datetime import datetime, timedelta

import openpyxl

import storage

MA_ST_CHINH = "8363"
CAU_HINH = {
    "14536": {
        "ten": "MỸ XUYÊN",
        "bi_danh": ["14536", "mx", "my xuyen"],
        "muc_tieu_nam_thang": 10_600_000,
        "so_ban_moi_ca": 3,
        "ty_le_sang": 0.4,
    },
}
NHOM_NAM = "Nấm Các Loại"
_COT_CAN = ("Ngày xuất", "Mã siêu thị", "Nhóm hàng", "Thành tiền phải thu khách hàng (chưa VAT)")


def _conn():
    conn = storage._connect()
    conn.execute("""CREATE TABLE IF NOT EXISTS mtkm_st_phu (
        ma_st TEXT, ngay TEXT, ten_st TEXT, nam REAL, gio TEXT, PRIMARY KEY (ma_st, ngay))""")
    return conn


def _ma(v):
    s = str(v).strip()
    return s[:-2] if s.endswith(".0") else s


def _khong_dau(s):
    s = unicodedata.normalize("NFD", str(s or "").lower().replace("đ", "d"))
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


def _ngay_str(v):
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    s = str(v or "").strip()
    for f in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s[:10], f).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def doc_file(path):
    """Đọc file POS 76. Trả về (theo_st, cac_ma):
      theo_st = {ma_st phụ: {"ten_st": .., "ngay": {yyyy-mm-dd: tiền nấm}}}  (chỉ siêu thị trong CAU_HINH)
      cac_ma  = tập mã siêu thị có trong file; None nếu không phải file POS 76."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        it = ws.iter_rows(values_only=True)
        try:
            header = [str(c).strip() if c is not None else "" for c in next(it)]
        except StopIteration:
            return {}, None
        if not all(c in header for c in _COT_CAN):
            return {}, None
        i_ng, i_st, i_nh, i_tt = (header.index(c) for c in _COT_CAN)
        i_ten = header.index("Tên siêu thị") if "Tên siêu thị" in header else None
        theo_st, cac_ma = {}, set()
        for row in it:
            if i_st >= len(row) or row[i_st] in (None, ""):
                continue
            ma = _ma(row[i_st])
            cac_ma.add(ma)
            if ma not in CAU_HINH:
                continue
            ngay = _ngay_str(row[i_ng])
            if not ngay:
                continue
            o = theo_st.setdefault(ma, {"ten_st": "", "ngay": {}})
            if not o["ten_st"] and i_ten is not None and row[i_ten]:
                o["ten_st"] = str(row[i_ten]).strip()
            o["ngay"].setdefault(ngay, 0.0)
            if str(row[i_nh] or "").strip() == NHOM_NAM:
                try:
                    o["ngay"][ngay] += float(row[i_tt] or 0)
                except (TypeError, ValueError):
                    pass
        return theo_st, cac_ma
    finally:
        wb.close()


def luu(theo_st, gio):
    conn = _conn()
    with conn:
        for ma, o in theo_st.items():
            for ngay, nam in o["ngay"].items():
                conn.execute("INSERT INTO mtkm_st_phu (ma_st, ngay, ten_st, nam, gio) VALUES (?, ?, ?, ?, ?) "
                             "ON CONFLICT(ma_st, ngay) DO UPDATE SET ten_st=excluded.ten_st, nam=excluded.nam, "
                             "gio=excluded.gio", (ma, ngay, o["ten_st"], nam, gio))
    conn.close()


def tom_tat(theo_st):
    ds = []
    for ma, o in theo_st.items():
        ngay = sorted(o["ngay"])
        ds.append(f"{ma} {CAU_HINH[ma]['ten']} ({len(ngay)} ngày "
                  f"{datetime.strptime(ngay[0], '%Y-%m-%d').strftime('%d/%m')}–"
                  f"{datetime.strptime(ngay[-1], '%Y-%m-%d').strftime('%d/%m')})")
    return "✅ Đã cập nhật MTKM nấm siêu thị " + ", ".join(ds) + "."


def nhan_dien_lenh(text):
    """'MTKM 14536' / 'MTKM MX' / 'MTKM Mỹ Xuyên' -> mã siêu thị phụ; không phải thì None."""
    t = re.sub(r"\s+", " ", _khong_dau(text)).strip()
    m = re.match(r"^(mtkm|muc tieu khuyen mai)\s+(.+)$", t)
    if not m:
        return None
    du = m.group(2).strip()
    for ma, ch in CAU_HINH.items():
        if du in ch["bi_danh"]:
            return ma
    return None


def tinh(ma_st):
    """Số cho thẻ: ngày báo cáo = ngày mới nhất đã có số của siêu thị đó."""
    ch = CAU_HINH[ma_st]
    conn = _conn()
    try:
        rows = conn.execute("SELECT ngay, nam, gio, ten_st FROM mtkm_st_phu WHERE ma_st = ? ORDER BY ngay",
                            (ma_st,)).fetchall()
    finally:
        conn.close()
    if not rows:
        return None
    ngay_bc, nam_bc, gio, _ten = rows[-1]
    d = datetime.strptime(ngay_bc, "%Y-%m-%d")
    dau = d.replace(day=1)
    so_ngay_thang = calendar.monthrange(d.year, d.month)[1]
    theo_ngay = {r[0]: r[1] for r in rows}
    truoc = [dau + timedelta(days=i) for i in range(d.day - 1)]
    lk_truoc = sum(theo_ngay.get(x.strftime("%Y-%m-%d"), 0.0) for x in truoc)
    thieu = [x.strftime("%d/%m") for x in truoc if x.strftime("%Y-%m-%d") not in theo_ngay]
    mt_thang = ch["muc_tieu_nam_thang"]
    con_ngay = so_ngay_thang - d.day + 1
    mt_ngay = max(mt_thang - lk_truoc, 0) / con_ngay
    sang = mt_ngay * ch["ty_le_sang"]
    return {
        "ma_st": ma_st, "ten": ch["ten"], "ngay": ngay_bc, "gio": gio or "", "mt_thang": mt_thang,
        "lk_truoc": lk_truoc, "hom_nay": nam_bc, "mt_ngay": mt_ngay, "con_ngay": con_ngay,
        "so_ngay_truoc": d.day - 1, "can_pct": (d.day - 1) / so_ngay_thang * 100,
        "sang": sang, "chieu": mt_ngay - sang, "so_ban": ch["so_ban_moi_ca"], "thieu": thieu,
        "ty_le_sang": ch["ty_le_sang"],
    }


def build_the(kq):
    import menu_bot as m
    from muc_tieu_flex import _bubble, _sec, _hang, _t, _o_tieu_de, _hop_canh_bao, NEN_TONG
    from flex_builder import GRAY, GREEN, BLACK, RED, MTKM_HL_BG, MTKM_GOLD_TEXT

    def k(v):
        return f"{v / 1000:,.0f}k".replace(",", ".")

    def tr(v):
        return f"{v / 1e6:.2f}tr".replace(".", ",")

    def mau(p, can):
        return GREEN if p >= can - 1e-9 else (m.VANG if p >= can * 0.8 else RED)

    d = datetime.strptime(kq["ngay"], "%Y-%m-%d")
    lk = kq["lk_truoc"] + kq["hom_nay"]
    p_ngay = kq["hom_nay"] / kq["mt_ngay"] * 100 if kq["mt_ngay"] else 100.0
    p_thang = lk / kq["mt_thang"] * 100 if kq["mt_thang"] else 0.0
    can = kq["can_pct"]
    body = []
    if kq["thieu"]:
        body.append(_hop_canh_bao([f"⚠ Chưa có số ngày {', '.join(kq['thieu'][:8])}"
                                   f"{'…' if len(kq['thieu']) > 8 else ''} – đang tính là 0, anh gửi thêm file "
                                   "POS 76 các ngày này"]))
    w = (30, 24, 24, 22)
    r1 = [_hang([_o_tieu_de(["Mục"], w[0], "start"), _o_tieu_de(["Mục tiêu"], w[1], "end"),
                 _o_tieu_de(["Đã bán"], w[2], "end"), _o_tieu_de(["% đạt"], w[3], "end")], MTKM_HL_BG, "4px"),
          _hang([_t(f"Nấm · ngày {d.strftime('%d/%m')}", "xs", MTKM_GOLD_TEXT, True, flex=w[0]),
                 _t(k(kq["mt_ngay"]), "sm", GRAY, False, "end", w[1]),
                 _t(k(kq["hom_nay"]), "sm", BLACK, True, "end", w[2]),
                 _t(f"{p_ngay:.0f}%", "sm", mau(p_ngay, 100), True, "end", w[3])], None, "4px"),
          _hang([_t("Nấm · lũy kế tháng", "xs", MTKM_GOLD_TEXT, True, flex=w[0]),
                 _t(tr(kq["mt_thang"]), "sm", GRAY, False, "end", w[1]),
                 _t(tr(lk), "sm", BLACK, True, "end", w[2]),
                 _t(f"{p_thang:.0f}%", "sm", mau(p_thang, can), True, "end", w[3])], NEN_TONG, "4px")]
    con_thieu = max(kq["mt_thang"] - kq["lk_truoc"], 0)
    body += [_sec("1. MTKM NGÀY · SIÊU THỊ"), m._bang(r1),
             m._ghi_chu(f"Mục tiêu tháng {tr(kq['mt_thang'])} · còn thiếu {tr(con_thieu)} chia đều "
                        f"{kq['con_ngay']} ngày còn lại → ngày {d.strftime('%d/%m')} cần {k(kq['mt_ngay'])} · "
                        f"tiến độ cần {can:.0f}%")]
    n = kq["so_ban"]
    ps = round(kq["ty_le_sang"] * 100)
    w2 = (30, 24, 24, 22)
    r2 = [_hang([_o_tieu_de(["Ca"], w2[0], "start"), _o_tieu_de(["Mục tiêu", "cả ca"], w2[1], "end"),
                 _o_tieu_de(["Thu ngân", "mỗi bạn"], w2[2], "end"), _o_tieu_de(["Fresh", "mỗi bạn"], w2[3], "end")],
                MTKM_HL_BG, "4px")]
    for ten, v in ((f"Sáng ({ps}%)", kq["sang"]), (f"Chiều ({100 - ps}%)", kq["chieu"])):
        r2.append(_hang([_t(ten, "xs", MTKM_GOLD_TEXT, True, flex=w2[0]), _t(k(v), "sm", GRAY, False, "end", w2[1]),
                         _t(k(v / n), "sm", BLACK, True, "end", w2[2]), _t(k(v / n), "sm", BLACK, True, "end", w2[3])],
                        None, "4px"))
    body += [_sec("2. MỤC TIÊU NGÀY · THU NGÂN / FRESH"), m._bang(r2),
             m._ghi_chu(f"Mỗi ca {n} bạn, chia đều mục tiêu ca · bán dư/thiếu hôm nay thì ngày mai tự giảm/tăng")]
    # nhận xét
    nx = f"Lũy kế {p_thang:.0f}% (cần {can:.0f}%)"
    if kq["so_ngay_truoc"]:
        tb = kq["lk_truoc"] / kq["so_ngay_truoc"]
        lech = (kq["mt_ngay"] / tb * 100 - 100) if tb else 0
        nx += (f": {kq['so_ngay_truoc']} ngày TB {k(tb)}/ngày, từ nay cần {k(kq['mt_ngay'])}/ngày "
               f"({'+' if lech >= 0 else '−'}{abs(lech):.0f}%)")
    gio = kq["gio"]
    if kq["hom_nay"] >= kq["mt_ngay"]:
        nx += f". Ngày {d.strftime('%d/%m')} đã đạt mục tiêu ({k(kq['hom_nay'])})."
    elif gio and gio < "14:00" and kq["hom_nay"] < kq["sang"]:
        nx += (f". Tới {gio} bán {k(kq['hom_nay'])}, ca sáng cần {k(kq['sang'])} – còn thiếu "
               f"{k(kq['sang'] - kq['hom_nay'])} trước khi đổi ca.")
    else:
        nx += (f". {'Tới ' + gio + ' b' if gio else 'B'}án {k(kq['hom_nay'])}/{k(kq['mt_ngay'])} – còn thiếu "
               f"{k(kq['mt_ngay'] - kq['hom_nay'])}.")
    body.append(m._nhan_xet(nx))
    phu = f"Ngày {d.strftime('%d/%m/%Y')}" + (f" · cập nhật {gio}" if gio else "") + " · nguồn POS 76"
    return _bubble(f"MTKM · {kq['ma_st']} {kq['ten']}", [phu], body)
