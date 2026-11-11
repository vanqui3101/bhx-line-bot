"""
menu_bot.py - (09/11/2026) MENU NÚT BẤM trong chat riêng với bot + các thẻ mới.

5 nút (gõ đúng chữ này trong nhóm cũng ra thẻ):
  1) "DOANH THU LŨY KẾ"  : doanh thu siêu thị (Fresh/FMCG) lũy kế tới HẾT HÔM QUA + dự kiến cả tháng,
                           kèm lũy kế các mục tiêu KM từ đầu tháng
  2) "LN FRESH"          : lợi nhuận 7 ngành Fresh theo file POS 3933 mới nhất anh gửi / bỏ Drive
  3) "MTKM"              : (thẻ có sẵn) mục tiêu ngày cả siêu thị
  4) "MỤC TIÊU TỪNG BẠN" : mục tiêu hôm nay của từng bạn thu ngân + fresh (theo bài phân line)
  5) "KẾT QUẢ 7 NGÀY"    : 7 ngày đã CHỐT SỐ gần nhất, theo vị trí, từng bạn: ngày đạt / chưa đạt
                           (% đạt ngày trên 80% mới tính ĐẠT; từ 80% trở xuống hoặc không báo = chưa đạt)
Lệnh "CÀI MENU" (chat riêng với bot): bot gắn menu cho đúng người gõ.
"""
import calendar
import json
import os
import re
import unicodedata
from datetime import datetime, timedelta

import storage
import mtkm_tracker
import muc_tieu_ngay
import thuong_freshfmcg
from muc_tieu_ngay import MUC_THU_NGAN, MUC_FRESH, pct_txt, _pct
from muc_tieu_flex import (_bubble, _sec, _hang, _t, _o_tieu_de, _hop_canh_bao, _mau_pct, NEN_TONG, NEN_XANH,
                           NEN_DO, VANG)
from flex_builder import RED, GRAY, GREEN, BLACK, MTKM_HL_BG, MTKM_GOLD_TEXT

NGUONG_DAT = 80  # % đạt ngày phải TRÊN mức này mới tính là ngày đạt
NEN_XAM = "#F2F2F2"


def _d(s):
    return datetime.strptime(s, "%Y-%m-%d").date()


def _s(d):
    return d.strftime("%Y-%m-%d")


def _bo_dau(s):
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D").lower()


def _tr(v, dau=False):
    """Tiền -> triệu: dưới 100tr 2 số lẻ, từ 100tr 1 số lẻ. dau=True thì có dấu +/−."""
    v = v or 0
    s = f"{abs(v) / 1e6:,.{1 if abs(v) >= 1e8 else 2}f}".replace(",", "_").replace(".", ",").replace("_", ".")
    if dau:
        return ("+" if v > 0 else "−" if v < 0 else "") + s
    return ("−" if v < 0 else "") + s


def _so(v, dv):
    if dv == "tr":
        return _tr(v)
    if dv == "kg":
        return f"{v or 0:.1f}".replace(".", ",")
    return f"{v or 0:.0f}"


def _bang(rows):
    return {"type": "box", "layout": "vertical", "paddingStart": "8px", "paddingEnd": "8px", "contents": rows}


def _ghi_chu(text):
    return {"type": "box", "layout": "vertical", "paddingAll": "10px", "paddingStart": "12px", "paddingEnd": "12px",
            "contents": [_t(text, "xxs", GRAY)]}


def _nhan_xet(text):
    return {"type": "box", "layout": "vertical", "paddingAll": "10px", "paddingStart": "12px", "paddingEnd": "12px",
            "contents": [{"type": "text", "text": "x", "wrap": True, "size": "sm", "contents": [
                {"type": "span", "text": "Nhận xét: ", "weight": "bold", "color": MTKM_GOLD_TEXT},
                {"type": "span", "text": text, "color": "#333333"}]}]}


# ---------------------------------------------------------------------------
# 1. DOANH THU LŨY KẾ (tới hết hôm qua)
# ---------------------------------------------------------------------------
_FRESH_THAP = {x.lower() for x in thuong_freshfmcg.FRESH_NGANH_HANG}


def _fresh_fmcg(payload):
    """Tách doanh thu 1 ngày (file POS 76 đã lưu) thành Fresh / FMCG theo cột Ngành hàng.
    So tên ngành KHÔNG phân biệt hoa/thường (file T11 ghi "Rau củ các loại", trước ghi "Rau Củ Các Loại")."""
    nh = (payload or {}).get("nganh_hang") or {}
    items = nh.get("items") if isinstance(nh, dict) and "items" in nh else \
        [{"ten": k, "thanh_tien": v} for k, v in (nh.items() if isinstance(nh, dict) else [])]
    fresh = sum(i.get("thanh_tien") or 0 for i in items if str(i.get("ten", "")).strip().lower() in _FRESH_THAP)
    tong = sum(i.get("thanh_tien") or 0 for i in items)
    return {"fresh_dt": fresh, "fmcg_dt": tong - fresh}


def tinh_doanh_thu(hom_nay_str):
    hq = _d(hom_nay_str) - timedelta(days=1)
    dau = hq.replace(day=1)
    so_ngay_thang = calendar.monthrange(hq.year, hq.month)[1]
    dau_truoc = (dau - timedelta(days=1)).replace(day=1)
    rows = {n: _fresh_fmcg(p) for n, p in storage.get_category_reports_range(_s(dau_truoc), _s(hq)).items()}
    ds = [dau + timedelta(days=i) for i in range(hq.day)]
    co = [x for x in ds if _s(x) in rows]
    if not co:
        return None
    so_ngay = len(co)
    # cùng kỳ tháng trước: chỉ so những ngày (theo số ngày trong tháng) có dữ liệu ở CẢ 2 tháng
    n_truoc = calendar.monthrange(dau_truoc.year, dau_truoc.month)[1]
    cap = [(x, dau_truoc.replace(day=x.day)) for x in co if x.day <= n_truoc and _s(dau_truoc.replace(day=x.day)) in rows]

    def _nhom(ten, khoa, base):
        lk = sum(rows[_s(x)][khoa] or 0 for x in co) if khoa else None
        return {"ten": ten, "khoa": khoa, "base": base, "lk": lk}

    nhom = [_nhom("Fresh", "fresh_dt", thuong_freshfmcg.BASE_FRESH),
            _nhom("FMCG", "fmcg_dt", thuong_freshfmcg.BASE_FMCG)]
    nhom.append({"ten": "Tổng", "khoa": None, "base": sum(n["base"] for n in nhom), "lk": sum(n["lk"] for n in nhom)})
    for n in nhom:
        khoas = [n["khoa"]] if n["khoa"] else ["fresh_dt", "fmcg_dt"]
        n["tb"] = n["lk"] / so_ngay
        n["du_kien"] = n["tb"] * so_ngay_thang
        n["pct_base"] = n["du_kien"] / n["base"] * 100 if n["base"] else None
        if cap:
            nay = sum(rows[_s(a)][k] or 0 for a, _b in cap for k in khoas)
            truoc = sum(rows[_s(b)][k] or 0 for _a, b in cap for k in khoas)
            n["pct_ck"] = (nay / truoc * 100 - 100) if truoc else None
        else:
            n["pct_ck"] = None
    kq_km = mtkm_tracker.tinh_mtkm(_s(hq))
    return {"hq": _s(hq), "dau": _s(dau), "so_ngay_thang": so_ngay_thang, "so_ngay": so_ngay,
            "thieu": [x.strftime("%d/%m") for x in ds if _s(x) not in rows], "nhom": nhom,
            "ck": (cap[0][1].strftime("%d"), cap[-1][1].strftime("%d/%m")) if cap else None, "km": kq_km}


def _mau_tien_do(p, can):
    if p is None:
        return GRAY
    return GREEN if p >= can - 1e-9 else (VANG if p >= can * 0.8 else RED)


def build_the_doanh_thu(kq):
    hq, dau = _d(kq["hq"]), _d(kq["dau"])
    body = []
    if kq["thieu"]:
        body.append(_hop_canh_bao([f"⚠ Chưa có file POS ngày {', '.join(kq['thieu'])} – lũy kế đang thấp hơn thực tế, "
                                   "dự kiến tháng tính theo TB các ngày đã có"]))
    # --- bảng doanh thu ---
    w = (20, 20, 18, 22, 16)
    rows = [_hang([_o_tieu_de(["Nhóm"], w[0], "start"), _o_tieu_de(["Lũy kế"], w[1], "end"),
                   _o_tieu_de(["TB", "/ngày"], w[2], "end"), _o_tieu_de(["Dự kiến", "tháng"], w[3], "end"),
                   _o_tieu_de(["So", "base"], w[4], "end")], MTKM_HL_BG, "4px")]
    for n in kq["nhom"]:
        tong = n["khoa"] is None
        rows.append(_hang([_t(n["ten"], "xs", MTKM_GOLD_TEXT, True, flex=w[0]),
                           _t(_tr(n["lk"]), "sm", BLACK, True, "end", w[1]),
                           _t(_tr(n["tb"]), "sm", GRAY, False, "end", w[2]),
                           _t(_tr(n["du_kien"]), "sm", BLACK, tong, "end", w[3]),
                           _t(pct_txt(n["pct_base"]), "sm", _mau_pct(n["pct_base"]), True, "end", w[4])],
                          NEN_TONG if tong else None, "4px"))
    body += [_sec("DOANH THU SIÊU THỊ (chưa VAT, triệu)"), _bang(rows)]
    if kq["ck"]:
        ck = " · ".join(f"{n['ten']} {('+' if n['pct_ck'] >= 0 else '−')}{abs(n['pct_ck']):.1f}%".replace(".", ",")
                        for n in kq["nhom"] if n["pct_ck"] is not None)
        body.append(_ghi_chu(f"So cùng kỳ tháng trước ({kq['ck'][0]}–{kq['ck'][1]}): {ck}"))
    # --- bảng lũy kế mục tiêu KM ---
    km = kq["km"]
    w2 = (30, 18, 18, 16, 18)
    r2 = [_hang([_o_tieu_de(["Mục"], w2[0], "start"), _o_tieu_de(["MT", "tháng"], w2[1], "end"),
                 _o_tieu_de(["Lũy kế"], w2[2], "end"), _o_tieu_de(["% đạt"], w2[3], "end"),
                 _o_tieu_de(["Còn", "thiếu"], w2[4], "end")], MTKM_HL_BG, "4px")]
    cham = []
    for d in km["dong"]:
        dv = d["don_vi"]
        mt = d["muc_tieu"]
        gd = d.get("tien_do_rieng") is not None
        if gd:  # thi đua theo tuần: mục tiêu = mục tiêu của tuần
            mt = d["luy_ke"] / d["pct_thang"] * 100 if d["pct_thang"] else 0
        can = d["tien_do_rieng"] if gd else km["pct_tien_do"]
        p = d["pct_thang"]
        if not d.get("la_tong") and p < can * 0.8:
            cham.append((p / can if can else 1, d["ten"], p, can))
        ten = d["ten"] + (" *" if gd else "")
        r2.append(_hang([_t(ten, "xs", MTKM_GOLD_TEXT, True, flex=w2[0]),
                         _t(_so(mt, dv), "sm", GRAY, False, "end", w2[1]),
                         _t(_so(d["luy_ke"], dv), "sm", BLACK, True, "end", w2[2]),
                         _t(pct_txt(p), "sm", _mau_tien_do(p, can), True, "end", w2[3]),
                         _t(_so(max(mt - d["luy_ke"], 0), dv), "sm", BLACK, False, "end", w2[4])],
                        NEN_TONG if d.get("la_tong") else None, "4px"))
    body += [_sec(f"LŨY KẾ MỤC TIÊU KM · cần {km['pct_tien_do']:.1f}%".replace(".", ",")), _bang(r2),
             _ghi_chu("Màu: xanh = theo kịp tiến độ · vàng = hơi chậm · đỏ = chậm nhiều · * thi đua theo tuần "
                      "(% và MT tính theo tuần) · đơn vị: triệu, trứng vỉ, heo kg, C2 lốc, 888 túi, kem đánh răng tuýp · "
                      "So base = dự kiến tháng ÷ base thưởng (Fresh 373tr · FMCG 1.056tr)")]
    tong = kq["nhom"][-1]
    nx = f"Dự kiến cả tháng {_tr(tong['du_kien'])}tr ({pct_txt(tong['pct_base'])} base)"
    if tong["pct_ck"] is not None:
        nx += f", {'cao' if tong['pct_ck'] >= 0 else 'thấp'} hơn cùng kỳ {abs(tong['pct_ck']):.1f}%".replace(".", ",")
    if cham:
        cham.sort()
        nx += "; chậm nhiều: " + ", ".join(f"{t} {pct_txt(p)} (cần {can:.0f}%)" for _r, t, p, can in cham[:3])
    body.append(_nhan_xet(nx + "."))
    return _bubble(f"DOANH THU · THÁNG {hq.month}",
                   [f"Lũy kế {dau.strftime('%d/%m')} – {hq.strftime('%d/%m')} (hết ngày hôm qua) · "
                    f"{kq['so_ngay']}/{kq['so_ngay_thang']} ngày có số"], body)


# ---------------------------------------------------------------------------
# 2. LN FRESH (file POS 3933 "Báo cáo hiệu quả bán Fresh siêu thị")
# ---------------------------------------------------------------------------
TEN_NGAN_LN = [("thit", "Thịt"), ("thuy hai san", "Thủy hải sản"), ("rau da lat", "Rau Đà Lạt"),
               ("rau dia phuong", "Rau địa phương"), ("trai cay ngoai", "Trái cây ngoại"),
               ("trai cay noi", "Trái cây nội"), ("trung", "Trứng")]


def _doc_header_ln(ws):
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True)):
        h = [_bo_dau(c).strip() if c is not None else "" for c in row]
        if "nganh hang" in h and "ln luy ke" in h:
            return i + 1, h
    return None, None


def la_file_ln_fresh(path):
    from openpyxl import load_workbook
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        ok = _doc_header_ln(wb.active)[0] is not None
        wb.close()
        return ok
    except Exception:
        return False


def _so_thuc(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("%", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def doc_ln_fresh(path, ten_file=""):
    """Trả về {"thang", "gio_xuat", "nganh": [{ten, dt, ti_le, ln, tb, chenh}], "tong": {...}}."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    dong_hd, h = _doc_header_ln(ws)
    if dong_hd is None:
        wb.close()
        raise ValueError("Không thấy cột 'Ngành hàng' / 'LN lũy kế' trong file POS 3933")

    def cot(*tu):
        for i, x in enumerate(h):  # trùng khớp nguyên tên cột trước (tránh "Mã ngành hàng" khi tìm "Ngành hàng")
            if len(tu) == 1 and x == tu[0]:
                return i
        for i, x in enumerate(h):
            if all(t in x for t in tu):
                return i
        return None

    c = {"thang": cot("thang"), "nganh": cot("nganh hang"), "dt": cot("dt fresh"), "ti_le": cot("ti le"),
         "ln": cot("ln luy ke"), "tb": cot("ln tb"), "chenh": cot("chenh lech")}
    nganh, thang = [], None
    for row in ws.iter_rows(min_row=dong_hd + 1, values_only=True):
        if c["nganh"] is None or c["nganh"] >= len(row) or not row[c["nganh"]]:
            continue
        g = lambda k: _so_thuc(row[c[k]]) if c[k] is not None and c[k] < len(row) else None
        ten_goc = str(row[c["nganh"]]).strip()
        kd = _bo_dau(ten_goc)
        ten = next((n for k, n in TEN_NGAN_LN if kd.startswith(k)), ten_goc)
        if thang is None and c["thang"] is not None and row[c["thang"]]:
            thang = str(row[c["thang"]]).strip()
        ln, tb = g("ln") or 0, g("tb") or 0
        chenh = g("chenh")
        nganh.append({"ten": ten, "dt": g("dt") or 0, "ti_le": g("ti_le"), "ln": ln, "tb": tb,
                      "chenh": chenh if chenh is not None else ln - tb})
    wb.close()
    if not nganh:
        raise ValueError("File POS 3933 không có dòng ngành hàng nào")
    m = re.search(r"(20\d{2})(\d{2})(\d{2})_(\d{2})(\d{2})\d{2}", ten_file or "")
    gio_xuat = f"{m.group(1)}-{m.group(2)}-{m.group(3)} {m.group(4)}:{m.group(5)}" if m else None
    tong = {k: sum(n[k] for n in nganh) for k in ("dt", "ln", "tb", "chenh")}
    return {"thang": thang, "gio_xuat": gio_xuat, "nganh": nganh, "tong": tong}


def _conn_ln():
    conn = storage._connect()
    conn.execute("CREATE TABLE IF NOT EXISTS ln_fresh (thang TEXT PRIMARY KEY, gio_xuat TEXT, gio_luu TEXT, "
                 "data_json TEXT NOT NULL)")
    return conn


def luu_ln_fresh(payload, gio_luu):
    """Lưu số LN mới nhất của tháng; file cũ hơn (giờ xuất sớm hơn) không đè file mới."""
    thang = payload.get("thang") or datetime.now().strftime("%Y%m")
    conn = _conn_ln()
    cu = conn.execute("SELECT gio_xuat FROM ln_fresh WHERE thang = ?", (thang,)).fetchone()
    if cu and cu[0] and payload.get("gio_xuat") and payload["gio_xuat"] < cu[0]:
        conn.close()
        return False
    with conn:
        conn.execute("INSERT INTO ln_fresh (thang, gio_xuat, gio_luu, data_json) VALUES (?, ?, ?, ?) "
                     "ON CONFLICT(thang) DO UPDATE SET gio_xuat=excluded.gio_xuat, gio_luu=excluded.gio_luu, "
                     "data_json=excluded.data_json",
                     (thang, payload.get("gio_xuat"), gio_luu, json.dumps(payload, ensure_ascii=False)))
    conn.close()
    return True


def lay_ln_fresh_moi_nhat():
    conn = _conn_ln()
    r = conn.execute("SELECT data_json, gio_luu FROM ln_fresh ORDER BY thang DESC LIMIT 1").fetchone()
    conn.close()
    if not r:
        return None
    p = json.loads(r[0])
    p["gio_luu"] = r[1]
    return p


def build_the_ln_fresh(p):
    w = (28, 18, 16, 19, 19)
    rows = [_hang([_o_tieu_de(["Ngành"], w[0], "start"), _o_tieu_de(["DT"], w[1], "end"),
                   _o_tieu_de(["DT/GV"], w[2], "end"), _o_tieu_de(["LN", "lũy kế"], w[3], "end"),
                   _o_tieu_de(["So TB", "3 tháng"], w[4], "end")], MTKM_HL_BG, "4px")]
    for n in p["nganh"]:
        tl = n["ti_le"]
        rows.append(_hang([_t(n["ten"], "xs", MTKM_GOLD_TEXT, True, flex=w[0]),
                           _t(_tr(n["dt"]), "sm", GRAY, False, "end", w[1]),
                           _t(f"{tl:.0f}%" if tl is not None else "–", "sm",
                              RED if tl is not None and tl < 100 else BLACK, False, "end", w[2]),
                           _t(_tr(n["ln"], True), "sm", GREEN if n["ln"] >= 0 else RED, True, "end", w[3]),
                           _t(_tr(n["chenh"], True), "sm", GREEN if n["chenh"] >= 0 else RED, True, "end", w[4])],
                          pad="4px"))
    t = p["tong"]
    rows.append(_hang([_t("TỔNG", "xs", MTKM_GOLD_TEXT, True, flex=w[0]), _t(_tr(t["dt"]), "sm", BLACK, True, "end", w[1]),
                       _t("", "sm", GRAY, False, "end", w[2]),
                       _t(_tr(t["ln"], True), "sm", GREEN if t["ln"] >= 0 else RED, True, "end", w[3]),
                       _t(_tr(t["chenh"], True), "sm", GREEN if t["chenh"] >= 0 else RED, True, "end", w[4])],
                      NEN_TONG, "4px"))
    duoi = sorted((n for n in p["nganh"] if n["chenh"] < 0), key=lambda n: n["chenh"])
    nx = (f"Tổng {'ĐẠT' if t['chenh'] > 0 else 'CHƯA ĐẠT'} tăng trưởng (chênh {_tr(t['chenh'], True)}tr so TB 3 tháng, "
          f"LN {_tr(t['ln'], True)}tr). ")
    if duoi:
        x = duoi[0]
        nx += (f"{len(duoi)}/{len(p['nganh'])} ngành dưới nền: " + ", ".join(f"{n['ten']} {_tr(n['chenh'], True)}"
                                                                         for n in duoi) + ". ")
        if x["ti_le"] is not None and x["ti_le"] < 100:
            nx += f"Kéo xuống nhiều nhất là {x['ten']}, DT/GV chỉ {x['ti_le']:.1f}%".replace(".", ",") + \
                  " – kiểm tra giảm giá, hủy, mất mát."
    else:
        nx += "Cả 7 ngành đều trên nền."
    gio = p.get("gio_xuat")
    dong_phu = [(f"Lũy kế tới lúc xuất file POS 3933 ({gio[11:]} ngày {gio[8:10]}/{gio[5:7]})" if gio
                 else "Lũy kế theo file POS 3933 mới nhất") + " · đơn vị: triệu",
                "Đạt khi chênh lệch so TB 3 tháng trước > 0"]
    thang = p.get("thang") or ""
    return _bubble(f"LN FRESH · THÁNG {int(thang[4:]) if len(thang) == 6 else ''}".strip(" ·"),
                   dong_phu, [_sec("THEO NGÀNH"), _bang(rows), _nhan_xet(nx.strip())])


# ---------------------------------------------------------------------------
# 4. MỤC TIÊU TỪNG BẠN HÔM NAY (thu ngân + fresh, theo bài phân line)
# ---------------------------------------------------------------------------
def build_the_muc_tieu_hom_nay(hom_nay_str):
    """Trả về (bubble, None) hoặc (None, câu báo lỗi)."""
    import muc_tieu_flex
    ch = muc_tieu_ngay.tinh_cua_hang(hom_nay_str)
    if ch is None:
        return None, "Tháng này chưa cài mục tiêu ngày anh ạ."
    pl = storage.get_phan_line(hom_nay_str)
    if not pl:
        return None, ("Hôm nay chưa có bài phân line nên chưa chia mục tiêu từng bạn anh ạ – "
                      "anh đăng bài phân line là bot chia ngay.")
    tb = muc_tieu_ngay.chia_tung_ban(ch, pl)
    return muc_tieu_flex.build_the_muc_tieu(hom_nay_str, ch, tb), None


# ---------------------------------------------------------------------------
# 5. KẾT QUẢ 7 NGÀY GẦN NHẤT — từng vị trí, từng bạn
# ---------------------------------------------------------------------------
def tinh_7_ngay(den_ngay_str, so_ngay=7):
    snaps = storage.get_chot_so_gan_nhat(den_ngay_str, so_ngay)
    kq = {"ngays": [n for n, _ in snaps]}
    for vai_tro, thu_tu in (("thu_ngan", MUC_THU_NGAN), ("fresh", MUC_FRESH)):
        nguoi = {}
        for ngay, data in snaps:
            for d in data.get(vai_tro) or []:
                mt = {k: v for k, v in (d.get("mt") or {}).items() if v and k in thu_tu}
                if not mt:  # báo ngoài line (không có mục tiêu) -> không tính
                    continue
                ps = [p for p in (_pct((((d.get("bc") or {}).get(k)) or 0), v) for k, v in mt.items()) if p is not None]
                pct = (sum(ps) / len(ps) if ps else 0) if d.get("da_bao") else 0
                x = nguoi.setdefault(d["ten"], {"ten": d["ten"], "ngay": {}})
                x["ngay"][ngay] = {"pct": pct, "da_bao": bool(d.get("da_bao")), "dat": pct > NGUONG_DAT}
        ds = []
        for x in nguoi.values():
            lam = len(x["ngay"])
            x["lam"] = lam
            x["dat"] = sum(1 for v in x["ngay"].values() if v["dat"])
            x["ti_le"] = x["dat"] / lam * 100 if lam else 0
            x["tb"] = sum(v["pct"] for v in x["ngay"].values()) / lam if lam else 0
            x["khong_bao"] = sum(1 for v in x["ngay"].values() if not v["da_bao"])
            ds.append(x)
        ds.sort(key=lambda x: (-x["ti_le"], -x["tb"], x["ten"]))
        kq[vai_tro] = ds
    return kq


def _chip(text, nen, mau, dam=True):
    """Ô 1 ngày (gọn tối đa: thẻ LINE giới hạn ~30KB, đông người x 7 ngày rất dễ vượt)."""
    o = {"type": "text", "text": text, "size": "xxs", "align": "center", "color": mau}
    if dam:
        o["weight"] = "bold"
    return {"type": "box", "layout": "vertical", "backgroundColor": nen, "cornerRadius": "3px", "contents": [o]}


def _o_ngay(v):
    if v is None:
        return _chip("·", NEN_XAM, "#BBBBBB", False)
    if not v["da_bao"]:
        return _chip("KB", NEN_DO, RED)
    return _chip(f"{v['pct']:.0f}", NEN_XANH if v["dat"] else NEN_DO, GREEN if v["dat"] else RED)


def _dong_7(cells, nen=None):
    o = {"type": "box", "layout": "horizontal", "spacing": "2px", "paddingAll": "3px", "contents": cells}
    if nen:
        o["backgroundColor"] = nen
    return o


def build_the_7_ngay(kq):
    ngays = kq["ngays"]
    body = []
    nx = []
    for vai_tro, tieu_de in (("thu_ngan", "THU NGÂN"), ("fresh", "FRESH")):
        ds = kq[vai_tro]
        if not ds:
            continue
        # ô ngày để flex mặc định (=1) cho nhẹ thẻ; cột Tên rộng gấp đôi, Đạt / % TB gấp 1,3
        rows = [_dong_7([_t("Tên", "xxs", MTKM_GOLD_TEXT, True, flex=20, wrap=False)] +
                        [_t(n[8:10], "xxs", MTKM_GOLD_TEXT, True, "center", 10, wrap=False) for n in ngays] +
                        [_t("Đạt", "xxs", MTKM_GOLD_TEXT, True, "end", 13, wrap=False),
                         _t("% TB", "xxs", MTKM_GOLD_TEXT, True, "end", 13, wrap=False)], MTKM_HL_BG)]
        for x in ds:
            mau_dat = GREEN if x["ti_le"] >= 80 else (VANG if x["ti_le"] >= 50 else RED)
            rows.append(_dong_7([_t(x["ten"], "xs", MTKM_GOLD_TEXT, True, flex=20, wrap=False)] +
                                [dict(_o_ngay(x["ngay"].get(n)), flex=10) for n in ngays] +
                                [_t(f"{x['dat']}/{x['lam']}", "sm", mau_dat, True, "end", 13, wrap=False),
                                 _t(pct_txt(x["tb"]), "xs", _mau_pct(x["tb"]), False, "end", 13, wrap=False)]))
        body += [_sec(tieu_de), _bang(rows)]
        cao, thap = ds[0], ds[-1]
        s = f"{tieu_de.capitalize()}: đều nhất {cao['ten']} ({cao['dat']}/{cao['lam']} ngày đạt)"
        if len(ds) > 1:
            s += f", yếu nhất {thap['ten']} ({thap['dat']}/{thap['lam']})"
        kb = [f"{x['ten']} {x['khong_bao']} ngày" for x in ds if x["khong_bao"]]
        if kb:
            s += "; không báo: " + ", ".join(kb)
        nx.append(s)
    if not body:
        return None
    body.append(_ghi_chu("Ô = % đạt mục tiêu ngày · xanh: trên 80% (đạt) · đỏ: từ 80% trở xuống · KB: không báo "
                         "(tính chưa đạt) · ô xám: không làm vị trí này · Đạt = số ngày đạt / số ngày làm · "
                         "% TB = trung bình % đạt các ngày làm"))
    body.append(_nhan_xet(". ".join(nx) + "."))
    d1, d2 = _d(ngays[0]), _d(ngays[-1])
    tieu_de = f"KẾT QUẢ {len(ngays)} NGÀY GẦN NHẤT"
    phu = [f"{d1.strftime('%d/%m')} – {d2.strftime('%d/%m')} · chỉ tính ngày đã CHỐT SỐ",
           "Từng vị trí · từng bạn · % đạt ngày trên 80% mới tính đạt"]
    the = _bubble(tieu_de, phu, body)
    if len(json.dumps(the, ensure_ascii=False, separators=(",", ":")).encode()) <= 28000:
        return the
    # đông người: tách THU NGÂN / FRESH thành 2 thẻ trượt ngang (mỗi thẻ LINE tối đa ~30KB)
    ds_the, khoi = [], []
    for o in body:
        if o.get("backgroundColor") and o.get("contents") and o["contents"][0].get("text") in ("THU NGÂN", "FRESH") and khoi:
            ds_the.append(khoi)
            khoi = []
        khoi.append(o)
    ds_the.append(khoi)
    return {"type": "carousel", "contents": [_bubble(tieu_de, phu, k) for k in ds_the]}


# ---------------------------------------------------------------------------
# LỆNH & MENU
# ---------------------------------------------------------------------------
LENH = [
    ("doanh_thu", re.compile(r"\s*(doanh\s*thu\s*luy\s*ke|dtlk)\s*")),
    ("ln_fresh", re.compile(r"\s*(ln\s*fresh|loi\s*nhuan\s*fresh)\s*")),
    ("muc_tieu_tung_ban", re.compile(r"\s*(muc\s*tieu\s*tung\s*ban|mttb)\s*")),
    ("bay_ngay", re.compile(r"\s*(ket\s*qua\s*7\s*ngay|7\s*ngay)\s*")),
    ("cai_menu", re.compile(r"\s*cai\s*menu\s*")),
    ("menu", re.compile(r"\s*menu\s*")),
]


def nhan_lenh(text_kd):
    for ten, pat in LENH:
        if pat.fullmatch(text_kd or ""):
            return ten
    return None


# (10/11/2026) THẺ MENU trong NHÓM: ai gõ "MENU" bot gửi thẻ 5 nút; nút nằm trong tin nhắn nên ai
# trong nhóm kéo lên bấm lúc nào cũng được (bấm = gửi đúng chữ lệnh, bot trả bảng vào nhóm).
# (10/11/2026) Nút bấm dùng POSTBACK: bấm vào KHÔNG hiện chữ lệnh trong khung chat, bot chỉ trả bảng.
NUT_MENU = [  # (dòng 1, dòng 2, mã lệnh, màu nền)
    ("Doanh thu", "lũy kế", "doanh_thu", "#0FA958"),
    ("LN", "Fresh", "ln_fresh", "#169C4B"),
    ("Mục tiêu", "ngày", "mtkm", "#12A36B"),
    ("Mục tiêu", "từng bạn", "muc_tieu_tung_ban", "#0B9A5E"),
    ("Kết quả", "7 ngày", "bay_ngay", "#119E55"),
]
LENH_POSTBACK = {x[2] for x in NUT_MENU}


def lenh_tu_postback(data):
    """'menu=doanh_thu' -> 'doanh_thu' (None nếu không phải nút menu)."""
    if (data or "").startswith("menu="):
        lenh = data[5:]
        return lenh if lenh in LENH_POSTBACK else None
    return None


def _nut(d1, d2, lenh, mau):
    return {"type": "box", "layout": "vertical", "flex": 1, "backgroundColor": mau, "cornerRadius": "12px",
            "paddingTop": "14px", "paddingBottom": "14px", "paddingStart": "6px", "paddingEnd": "6px",
            "action": {"type": "postback", "label": (d1 + " " + d2)[:20], "data": "menu=" + lenh},
            "contents": [_t(d1, "md", "#FFFFFF", True, "center", wrap=False),
                         _t(d2, "md", "#FFFFFF", True, "center", wrap=False)]}


def build_the_menu():
    """(11/11/2026) Không tiêu đề, không dòng chữ nhỏ. Bố cục: 2 ô trái – 2 ô phải, 1 ô ở giữa bên dưới."""
    n = [_nut(*x) for x in NUT_MENU]  # 0 Doanh thu, 1 LN Fresh, 2 Mục tiêu ngày, 3 Từng bạn, 4 Kết quả 7 ngày
    hang = lambda ds: {"type": "box", "layout": "horizontal", "spacing": "8px", "contents": ds}
    dem = {"type": "filler"}
    giua = dict(n[4], flex=2)
    return {"type": "bubble", "size": "mega",
            "body": {"type": "box", "layout": "vertical", "spacing": "8px", "paddingAll": "12px",
                     "backgroundColor": "#F2FBF5",
                     "contents": [hang([n[0], n[2]]), hang([n[1], n[3]]),
                                  {"type": "box", "layout": "horizontal", "spacing": "8px",
                                   "contents": [dict(dem, flex=1), giua, dict(dem, flex=1)]}]}}


# Ô bấm trên ảnh menu 2500x1686: hàng trên 3 ô, hàng dưới 2 ô (khớp ảnh richmenu.jpg)
MENU_O = [  # (ô, nhãn, mã lệnh) — postback, bấm không hiện chữ trong khung chat
    ((0, 0, 833, 843), "Doanh thu lũy kế", "doanh_thu"),
    ((833, 0, 834, 843), "LN Fresh", "ln_fresh"),
    ((1667, 0, 833, 843), "Mục tiêu ngày", "mtkm"),
    ((0, 843, 1250, 843), "Mục tiêu từng bạn", "muc_tieu_tung_ban"),
    ((1250, 843, 1250, 843), "Kết quả 7 ngày", "bay_ngay"),
]
TEN_MENU = "TROLY menu 5 o"  # gõ CÀI MENU lại thì menu cũ cùng tên bị xoá, gắn menu mới
ANH_MENU = os.path.join(os.path.dirname(os.path.abspath(__file__)), "richmenu.jpg")


def cai_menu_cho(user_id, messaging_api, blob_api):
    """Tạo menu (xoá menu cũ cùng tên), tải ảnh, gắn cho đúng user_id. Trả về rich_menu_id."""
    from linebot.v3.messaging import RichMenuRequest, RichMenuSize, RichMenuArea, RichMenuBounds, PostbackAction
    try:
        for m in (messaging_api.get_rich_menu_list().richmenus or []):
            if m.name == TEN_MENU:
                messaging_api.delete_rich_menu(m.rich_menu_id)
    except Exception:
        import traceback
        traceback.print_exc()
    areas = [RichMenuArea(bounds=RichMenuBounds(x=x, y=y, width=w, height=h),
                          action=PostbackAction(label=nhan[:20], data="menu=" + lenh))
             for (x, y, w, h), nhan, lenh in MENU_O]
    rm = messaging_api.create_rich_menu(rich_menu_request=RichMenuRequest(
        size=RichMenuSize(width=2500, height=1686), selected=True, name=TEN_MENU,
        chat_bar_text="Menu", areas=areas))
    with open(ANH_MENU, "rb") as f:
        blob_api.set_rich_menu_image(rich_menu_id=rm.rich_menu_id, body=bytearray(f.read()),
                                     _headers={"Content-Type": "image/jpeg"})
    messaging_api.link_rich_menu_id_to_user(user_id, rm.rich_menu_id)
    return rm.rich_menu_id


# ---------------------------------------------------------------------------
# (11/11/2026) LỊCH SỬ BẤM MENU — ai bấm ô nào, ngày giờ nào (chỉ anh Quí xem trong chat riêng)
#   "LỊCH SỬ MENU"         -> hôm nay
#   "LỊCH SỬ MENU dd/mm"   -> 1 ngày
#   "LỊCH SỬ MENU 7 NGÀY"  -> tổng 7 ngày theo từng người
# ---------------------------------------------------------------------------
TEN_NUT = {"doanh_thu": "Doanh thu lũy kế", "ln_fresh": "LN Fresh", "mtkm": "Mục tiêu ngày",
           "muc_tieu_tung_ban": "Mục tiêu từng bạn", "bay_ngay": "Kết quả 7 ngày", "menu": "Mở MENU"}
_LENH_LICH_SU = re.compile(r"\s*lich\s*su\s*(?:bam\s*)?menu(?:\s+(?:(\d{1,2})\s*/\s*(\d{1,2})|(7)\s*ngay))?\s*")


def _conn_log():
    conn = storage._connect()
    conn.execute("CREATE TABLE IF NOT EXISTS menu_log (id INTEGER PRIMARY KEY AUTOINCREMENT, ngay TEXT NOT NULL, "
                 "gio TEXT, user_id TEXT, ten TEXT, nut TEXT, kieu TEXT, noi TEXT)")
    return conn


def ghi_bam_menu(ngay, gio, user_id, ten, nut, kieu, noi):
    """kieu: 'bấm' (bấm ô) | 'gõ' (gõ chữ lệnh); noi: 'Nhóm' | 'Riêng'."""
    conn = _conn_log()
    with conn:
        conn.execute("INSERT INTO menu_log (ngay, gio, user_id, ten, nut, kieu, noi) VALUES (?, ?, ?, ?, ?, ?, ?)",
                     (ngay, gio, user_id, ten, nut, kieu, noi))
    conn.close()


def lay_lich_su(tu, den):
    conn = _conn_log()
    rows = conn.execute("SELECT ngay, gio, user_id, ten, nut, kieu, noi FROM menu_log WHERE ngay BETWEEN ? AND ? "
                        "ORDER BY ngay, gio, id", (tu, den)).fetchall()
    conn.close()
    return [dict(zip(("ngay", "gio", "user_id", "ten", "nut", "kieu", "noi"), r)) for r in rows]


def doc_lenh_lich_su(text_kd, hom_nay_str):
    """Trả về (tu, den, kieu) với kieu 'ngay' | '7ngay'; None nếu không phải lệnh."""
    m = _LENH_LICH_SU.fullmatch(text_kd or "")
    if not m:
        return None
    hn = _d(hom_nay_str)
    if m.group(3):
        return _s(hn - timedelta(days=6)), _s(hn), "7ngay"
    if m.group(1):
        d = _s(hn.replace(month=int(m.group(2)), day=int(m.group(1))))
        return d, d, "ngay"
    return _s(hn), _s(hn), "ngay"


def build_the_lich_su_ngay(ngay, rows, toi_da=45):
    dn = _d(ngay)
    tieu_de = f"LỊCH SỬ BẤM MENU · {dn.strftime('%d/%m')}"
    if not rows:
        return _bubble(tieu_de, ["Chỉ anh xem được trong chat riêng"],
                       [_ghi_chu("Ngày này chưa ai bấm menu.")])
    nguoi = {}
    for r in rows:
        nguoi[r["ten"]] = nguoi.get(r["ten"], 0) + 1
    w = (13, 27, 38, 14)
    ds = [_hang([_o_tieu_de(["Giờ"], w[0], "start"), _o_tieu_de(["Ai bấm"], w[1], "start"),
                 _o_tieu_de(["Ô"], w[2], "start"), _o_tieu_de(["Ở"], w[3], "end")], MTKM_HL_BG, "4px")]
    hien = rows[-toi_da:]
    for r in hien:
        ds.append(_hang([_t(r["gio"][:5], "xs", GRAY, False, flex=w[0], wrap=False),
                         _t(r["ten"], "xs", MTKM_GOLD_TEXT, True, flex=w[1], wrap=False),
                         _t(TEN_NUT.get(r["nut"], r["nut"]) + (" (gõ)" if r["kieu"] == "gõ" else ""), "xs", BLACK,
                            flex=w[2], wrap=False),
                         _t(r["noi"], "xs", GRAY, False, "end", w[3], wrap=False)], pad="3px"))
    body = [_sec(f"{len(rows)} LẦN · {len(nguoi)} NGƯỜI"), _bang(ds)]
    if len(rows) > len(hien):
        body.append(_ghi_chu(f"Chỉ hiện {toi_da} lần gần nhất trong ngày."))
    top = sorted(nguoi.items(), key=lambda x: -x[1])
    body.append(_nhan_xet(", ".join(f"{t} {n} lần" for t, n in top[:6]) + "."))
    return _bubble(tieu_de, ["Ai bấm ô nào, mấy giờ · chỉ anh xem được trong chat riêng"], body)


def build_the_lich_su_7_ngay(tu, den, rows):
    d1, d2 = _d(tu), _d(den)
    tieu_de = f"LỊCH SỬ BẤM MENU · {d1.strftime('%d/%m')} – {d2.strftime('%d/%m')}"
    if not rows:
        return _bubble(tieu_de, ["Chỉ anh xem được trong chat riêng"], [_ghi_chu("7 ngày qua chưa ai bấm menu.")])
    nguoi = {}
    for r in rows:
        x = nguoi.setdefault(r["ten"], {"lan": 0, "ngay": set(), "nut": {}, "cuoi": ""})
        x["lan"] += 1
        x["ngay"].add(r["ngay"])
        x["nut"][r["nut"]] = x["nut"].get(r["nut"], 0) + 1
        x["cuoi"] = max(x["cuoi"], r["ngay"] + " " + (r["gio"] or ""))
    w = (25, 11, 11, 31, 22)
    ds = [_hang([_o_tieu_de(["Ai bấm"], w[0], "start"), _o_tieu_de(["Lần"], w[1], "end"),
                 _o_tieu_de(["Số", "ngày"], w[2], "end"), _o_tieu_de(["Bấm nhiều nhất"], w[3], "start"),
                 _o_tieu_de(["Lần cuối"], w[4], "end")], MTKM_HL_BG, "4px")]
    for ten, x in sorted(nguoi.items(), key=lambda kv: -kv[1]["lan"]):
        nut = max(x["nut"].items(), key=lambda kv: kv[1])[0]
        cuoi = x["cuoi"]
        ds.append(_hang([_t(ten, "xs", MTKM_GOLD_TEXT, True, flex=w[0], wrap=False),
                         _t(str(x["lan"]), "sm", BLACK, True, "end", w[1]),
                         _t(str(len(x["ngay"])), "sm", GRAY, False, "end", w[2]),
                         _t(TEN_NUT.get(nut, nut), "xs", BLACK, flex=w[3], wrap=False),
                         _t(f"{cuoi[8:10]}/{cuoi[5:7]} {cuoi[11:16]}", "xs", GRAY, False, "end", w[4], wrap=False)],
                        pad="3px"))
    dem_nut = {}
    for r in rows:
        dem_nut[r["nut"]] = dem_nut.get(r["nut"], 0) + 1
    nx = "Ô được bấm nhiều nhất: " + ", ".join(f"{TEN_NUT.get(k, k)} {v} lần"
                                              for k, v in sorted(dem_nut.items(), key=lambda kv: -kv[1])[:3]) + "."
    return _bubble(tieu_de, [f"{len(rows)} lần bấm · {len(nguoi)} người · chỉ anh xem được trong chat riêng"],
                   [_sec("THEO TỪNG NGƯỜI"), _bang(ds), _nhan_xet(nx)])
