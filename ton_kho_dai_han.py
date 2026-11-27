"""
TỒN KHO DÀI HẠN (TKDH) — 27/11/2026, anh Quí.

- Nguồn danh sách: file công ty "20261126_TỒN KHO DÀI HẠN.xlsx" (sheet "Chi tiết ST"), lọc siêu thị 8363 -> 20 mã.
  Mỗi mã có: tồn ST chốt 26/11, SL cần duy trì (cần giữ), SL cần xử lý.
- Anh Quí: KHÔNG chạy theo tuần (tuần 1 nước, tuần 2 hóa phẩm…) — chạy TẤT CẢ 20 mã cùng lúc cho tới khi xử lý hết
  phần cần xử lý (tồn về mức "cần giữ").
- (27/11/2026 18:46) Tồn mốc lấy từ POS 322 thật (không tính hàng đi đường). Anh Quí: hàng này KHÔNG điều chuyển,
  chỉ bán tại siêu thị -> mỗi ngày chỉ cần nạp POS 76, bot trừ số bán từ 28/11.
  (Coke Sleek: file công ty ghi tồn 4.320 lon nhưng POS 322 thật chỉ 335 lon -> dùng số POS.)
- Tồn hôm nay:
    * Nếu sau này có file POS 322 (BC tồn theo model) nạp từ 28/11 trở đi -> lấy tồn trong file đó, rồi trừ tiếp số
      bán POS 76 của các ngày SAU ngày file tồn.
    * Nếu không -> tồn mốc 27/11 trừ số bán POS 76 từ 28/11.
- Còn phải xử lý = max(tồn hôm nay − cần giữ, 0). Mã nào về 0 là XONG.
- Lệnh: "TKDH" / "TỒN KHO DÀI HẠN".
"""
from datetime import datetime, timedelta

import storage

MA_ST = "8363"
NGAY_CHOT = "2026-11-27"      # tồn mốc POS 322 lúc 18:46 27/11
GIO_CHOT = "18:46"
TINH_TU = "2026-11-28"        # bắt đầu trừ số bán POS 76

# model: (tên ngắn, ngành, trạng thái, quy cách thùng, tồn mốc 27/11 (POS 322), cần giữ (file công ty), cần xử lý = tồn − cần giữ)
DANH_SACH = {
    "1607002279": ("Coke Sleek lon 330ml", "nuoc", "KD", 24, 335, 89, 246),
    "1607002399": ("Nutri Boost dâu 297ml", "nuoc", "KD", 24, 1628, 48, 1580),
    "1607002175": ("C2 chanh 500ml", "nuoc", "KD", 24, 1033, 95, 938),
    "1607002398": ("Nutri Boost cam 297ml", "nuoc", "KD", 24, 493, 11, 482),
    "2604001002": ("Nước mắm cá hồi D-Shining 500ml", "thuc_pham", "Xả tồn", 12, 150, 7, 143),
    "1802001090": ("Sốt thịt nướng ngũ vị Barona", "thuc_pham", "KD", 80, 109, 22, 87),
    "1709000610": ("Cá nục sốt cà Lilly 155g", "thuc_pham", "KD", 100, 143, 97, 46),
    "1712000577": ("Xốt bò kho Barona", "thuc_pham", "KD", 80, 71, 30, 41),
    "1607000580": ("Cá hộp 3 Cô Gái 190g", "thuc_pham", "KD", 48, 42, 9, 33),
    "1812002489": ("Knorr canh chua 30g", "thuc_pham", "KD", 60, 68, 39, 29),
    "2307003356": ("Cốt bún riêu Ông Chà Và", "thuc_pham", "KD", 60, 40, 16, 24),
    "2307003361": ("Sốt xá xíu Ông Chà Và", "thuc_pham", "KD", 60, 44, 20, 24),
    "1703001335": ("Sốt sườn chua ngọt Barona", "thuc_pham", "KD", 80, 47, 25, 22),
    "2307003355": ("Cốt hủ tiếu Nam Vang Ông Chà Và", "thuc_pham", "KD", 60, 38, 16, 22),
    "1806000066": ("Tắm gội Carrie Junior cherry", "my_pham", "Ngưng KD", 24, 203, 39, 164),
    "1607000568": ("BVS Laurier hằng ngày 20m", "my_pham", "KD", 48, 82, 28, 54),
    "1607004723": ("BVS Laurier siêu mỏng cánh 20m", "my_pham", "KD", 45, 45, 15, 30),
    "2605001355": ("Bàn chải Colgate Slim Soft 1+1", "my_pham", "Xả tồn", 36, 36, 7, 29),
    "2309017023": ("Nước giặt Lix ngàn hoa 3.5kg", "hoa_pham", "KD", 4, 36, 7, 29),
    "2001000002": ("Xả vải IZI Home 2.4L", "hoa_pham", "KD", 6, 31, 7, 24),
}
NGANH = [("nuoc", "THỨC UỐNG GIẢI KHÁT"), ("thuc_pham", "THỰC PHẨM - GIA VỊ"),
         ("my_pham", "MỸ PHẨM"), ("hoa_pham", "HÓA PHẨM")]
QUY_CACH = {m: v[3] for m, v in DANH_SACH.items()}   # excel_reader dùng để quy THÙNG -> đơn vị lẻ


def he_so_don_vi(model, don_vi):
    """1 dòng bán/tồn theo THÙNG -> nhân quy cách; còn lại tính 1."""
    return QUY_CACH.get(model, 1) if "THÙNG" in str(don_vi or "").upper() else 1


def _conn():
    conn = storage._connect()
    conn.execute("""CREATE TABLE IF NOT EXISTS tkdh_ton (
        ngay TEXT, model TEXT, ton REAL, gio TEXT, PRIMARY KEY (ngay, model))""")
    return conn


def luu_ton(stock_payload, ngay_str, gio=None):
    """Gọi khi bot nhận file POS 322 (BC tồn theo model): lưu tồn của 20 mã TKDH theo ngày.
    Mã không có trong file: nếu là file tồn ĐẦY ĐỦ cả siêu thị (> 500 dòng) -> coi như tồn 0 (đã hết);
    file lọc ít dòng -> không ghi mã đó (giữ cách tính cũ, tránh báo "xong" sai)."""
    rows = stock_payload.get("rows") or []
    if not rows:
        return 0
    file_day_du = len(rows) > 500
    ton = {m: 0.0 for m in DANH_SACH} if file_day_du else {}
    for r in rows:
        m = str(r.get("model") or "").strip()
        if m.endswith(".0"):
            m = m[:-2]
        if m in DANH_SACH:
            ton[m] = ton.get(m, 0.0) + float(r.get("ton_kho") or 0) * he_so_don_vi(m, r.get("don_vi"))
    if not ton:
        return 0
    conn = _conn()
    with conn:
        conn.execute("DELETE FROM tkdh_ton WHERE ngay = ?", (ngay_str,))
        conn.executemany("INSERT INTO tkdh_ton (ngay, model, ton, gio) VALUES (?, ?, ?, ?)",
                         [(ngay_str, m, v, gio) for m, v in ton.items()])
    conn.close()
    return len(ton)


def _ton_moi_nhat():
    conn = _conn()
    row = conn.execute("SELECT MAX(ngay) FROM tkdh_ton WHERE ngay >= ?", (TINH_TU,)).fetchone()
    if not row or not row[0]:
        conn.close()
        return None, None, {}
    ngay = row[0]
    rs = conn.execute("SELECT model, ton, gio FROM tkdh_ton WHERE ngay = ?", (ngay,)).fetchall()
    conn.close()
    return ngay, (rs[0][2] if rs else None), {m: t for m, t, _g in rs}


def _d(s):
    return datetime.strptime(s, "%Y-%m-%d").date()


def tinh(hom_nay_str):
    reports = storage.get_category_reports_range(TINH_TU, hom_nay_str) or {}
    ban_ngay = {}
    thieu = []
    d = _d(TINH_TU)
    while d <= _d(hom_nay_str):
        ns = d.strftime("%Y-%m-%d")
        p = reports.get(ns)
        if p is not None and "tkdh_ban" in p:
            ban_ngay[ns] = p.get("tkdh_ban") or {}
        elif d < _d(hom_nay_str):
            thieu.append(d.strftime("%d/%m"))
        d += timedelta(days=1)
    so_den = max(ban_ngay) if ban_ngay else None

    ngay_ton, gio_ton, ton_file = _ton_moi_nhat()
    dong = []
    for m, (ten, nganh, tt, qc, ton_dau, can_giu, can_xl) in DANH_SACH.items():
        if ngay_ton and m in ton_file:
            goc = ton_file[m]
            ban_sau = sum(b.get(m, 0) for n, b in ban_ngay.items() if n > ngay_ton)
        else:
            goc = ton_dau
            ban_sau = sum(b.get(m, 0) for b in ban_ngay.values())
        ban_tong = sum(b.get(m, 0) for b in ban_ngay.values())
        ton_nay = max(goc - ban_sau, 0)
        con = max(ton_nay - can_giu, 0)
        da = max(can_xl - con, 0)
        dong.append({"model": m, "ten": ten, "nganh": nganh, "tt": tt, "qc": qc, "ton_dau": ton_dau,
                     "can_giu": can_giu, "can_xl": can_xl, "ton_nay": ton_nay, "con": con, "da": da,
                     "ban": ban_tong, "pct": (da / can_xl * 100) if can_xl else 100.0, "xong": con <= 0})
    tong_can = sum(x["can_xl"] for x in dong)
    tong_da = sum(x["da"] for x in dong)
    return {"ngay": hom_nay_str, "dong": dong, "so_xong": sum(1 for x in dong if x["xong"]),
            "tong_can": tong_can, "tong_da": tong_da, "thieu": thieu, "so_den": so_den,
            "ngay_ton": ngay_ton, "gio_ton": gio_ton}


def build_the(kq):
    import menu_bot as mb
    from muc_tieu_flex import _bubble, _sec, _hang, _t, _o_tieu_de, _hop_canh_bao, NEN_TONG
    from flex_builder import GRAY, GREEN, BLACK, RED, MTKM_HL_BG, MTKM_GOLD_TEXT

    def so(v):
        return f"{v:,.0f}".replace(",", ".")

    def mau(p):
        return GREEN if p >= 100 - 1e-9 else (mb.VANG if p >= 50 else RED)

    body = []
    if kq["thieu"]:
        body.append(_hop_canh_bao([f"⚠ Chưa có số bán POS 76 ngày {', '.join(kq['thieu'][:8])}"
                                   f"{'…' if len(kq['thieu']) > 8 else ''} – đang tính là chưa bán, anh nạp thêm "
                                   "file POS 76 các ngày này"]))
    p_tong = kq["tong_da"] / kq["tong_can"] * 100 if kq["tong_can"] else 0
    w0 = (40, 30, 30)
    body += [_sec("TỔNG 20 MÃ · CHẠY TẤT CẢ ĐẾN HẾT TỒN"),
             mb._bang([_hang([_o_tieu_de(["Đã xử lý xong"], w0[0], "start"), _o_tieu_de(["Cần xử lý"], w0[1], "end"),
                              _o_tieu_de(["Đã xử lý"], w0[2], "end")], MTKM_HL_BG, "4px"),
                       _hang([_t(f"{kq['so_xong']}/{len(kq['dong'])} mã", "sm", BLACK, True, flex=w0[0]),
                              _t(so(kq["tong_can"]), "sm", GRAY, False, "end", w0[1]),
                              _t(f"{so(kq['tong_da'])} ({p_tong:.0f}%)", "sm", mau(p_tong), True, "end", w0[2])],
                             NEN_TONG, "4px")])]
    w = (40, 20, 20, 20)
    body2 = []
    for nk, ten_ng in NGANH:
        ds = sorted([x for x in kq["dong"] if x["nganh"] == nk and not x["xong"]], key=lambda x: -x["con"])
        if not ds:
            continue
        rows = [_hang([_o_tieu_de(["Sản phẩm"], w[0], "start"), _o_tieu_de(["Cần XL"], w[1], "end"),
                       _o_tieu_de(["Đã XL"], w[2], "end"), _o_tieu_de(["Còn"], w[3], "end")], MTKM_HL_BG, "4px")]
        for x in ds:
            nhan = x["ten"] + (f" ({x['tt']})" if x["tt"] != "KD" else "")
            rows.append(_hang([_t(nhan, "xs", MTKM_GOLD_TEXT, True, flex=w[0]),
                               _t(so(x["can_xl"]), "xs", GRAY, False, "end", w[1]),
                               _t(so(x["da"]), "xs", mau(x["pct"]), True, "end", w[2]),
                               _t(so(x["con"]), "xs", BLACK, True, "end", w[3])], None, "4px"))
        (body if nk in ("nuoc", "thuc_pham") else body2).extend([_sec(f"{ten_ng} · {len(ds)} mã"), mb._bang(rows)])
    # thẻ 1: tổng + nước + thực phẩm; thẻ 2: mỹ phẩm, hóa phẩm, ghi chú
    xong = [x for x in kq["dong"] if x["xong"]]
    if xong:
        body2.append(mb._ghi_chu("✅ Đã xử lý xong: " + " · ".join(x["ten"] for x in xong)))
    # nhận xét
    nuoc = [x for x in kq["dong"] if x["nganh"] == "nuoc"]
    can_nuoc = sum(x["can_xl"] for x in nuoc)
    nx = (f"4 mã nước chiếm {can_nuoc / kq['tong_can'] * 100:.0f}% lượng cần xử lý "
          f"({so(sum(x['con'] for x in nuoc))} chai/lon còn lại) – chỉ bán tại siêu thị nên phải chưng ụ lớn ngay "
          "quầy thu ngân + mời từng khách (C2 chanh còn cộng luôn vào mục tiêu C2). "
          "Các mã gia vị, mỹ phẩm nhỏ: chưng ụ + mời tại quầy là hết.") if kq["tong_can"] else ""
    if nx:
        body2.append(mb._nhan_xet(nx))
    nguon = []
    if kq["ngay_ton"]:
        nguon.append(f"tồn POS 322 ngày {_d(kq['ngay_ton']).strftime('%d/%m')}"
                     + (f" {kq['gio_ton']}" if kq.get("gio_ton") else ""))
    else:
        nguon.append(f"tồn mốc POS 322 {GIO_CHOT} {_d(NGAY_CHOT).strftime('%d/%m')} trừ số bán")
    if kq["so_den"]:
        nguon.append(f"bán POS 76 đến {_d(kq['so_den']).strftime('%d/%m')}")
    body2.append(mb._ghi_chu("Cần XL = phần phải xử lý để tồn về mức cần giữ (file công ty) · Còn = còn phải xử lý · "
                            "Nguồn: " + " · ".join(nguon) + ". Mỗi ngày chỉ cần nạp POS 76."))
    phu = f"Ngày {_d(kq['ngay']).strftime('%d/%m/%Y')} · 20 mã · không chia tuần, chạy đến hết tồn"
    # (27/11/2026) 1 thẻ vượt giới hạn 30KB của LINE -> tách 2 thẻ (vuốt ngang)
    return [_bubble("TỒN KHO DÀI HẠN · 8363 (1/2)", [phu], body),
            _bubble("TỒN KHO DÀI HẠN · 8363 (2/2)", [phu], body2)]
