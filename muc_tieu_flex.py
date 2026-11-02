"""
muc_tieu_flex.py - Vẽ 2 thẻ Flex (01/11/2026):
  - build_the_muc_tieu(...)  : MỤC TIÊU NGÀY TỪNG BẠN (gửi ngay khi anh đăng bài phân line)
  - build_the_chot_so(...)   : CHỐT SỐ cuối ngày (lệnh "CHỐT SỐ")
Số liệu tính sẵn ở muc_tieu_ngay.py; file này CHỈ vẽ thẻ.
"""
import json
from datetime import datetime

import storage
from flex_builder import (YELLOW, RED, GRAY, GREEN, DIVIDER, PAGE_BG, BLACK,
                          MTKM_GOLD_TEXT, MTKM_GOLD_SUB, MTKM_HL_BG, MTKM_SEC_BG)
from muc_tieu_ngay import MUC, MUC_THU_NGAN, MUC_FRESH, fmt_so, fmt_co_don_vi, ca_txt

VANG = "#C27C0E"
NEN_TONG = "#FFF8EC"
NEN_XANH = "#EAF6EE"
CHU_XANH = "#1E6B3A"
NEN_DO = "#FDECEA"


def _ngay_hien(ngay_str):
    return datetime.strptime(ngay_str, "%Y-%m-%d").strftime("%d/%m/%Y")


def _mau_pct(p):
    if p is None:
        return GRAY
    return GREEN if p >= 100 else (VANG if p >= 80 else RED)


def _t(text, size="sm", color=BLACK, bold=False, align="start", flex=None, wrap=True, margin=None):
    """Text gọn: chỉ ghi thuộc tính khác mặc định (thẻ LINE giới hạn dung lượng ~30KB/bubble)."""
    o = {"type": "text", "text": str(text) if str(text).strip() else "–", "size": size, "color": color}
    if bold:
        o["weight"] = "bold"
    if align != "start":
        o["align"] = align
    if wrap:
        o["wrap"] = True
    if flex is not None:
        o["flex"] = flex
    if margin:
        o["margin"] = margin
    return o


def _sec(title):
    return {"type": "box", "layout": "vertical", "backgroundColor": MTKM_SEC_BG, "paddingAll": "6px",
            "paddingStart": "12px", "margin": "md",
            "contents": [_t(title, "sm", "#FFFFFF", True)]}


def _hang(cells, bg=None, pad="6px"):
    o = {"type": "box", "layout": "horizontal", "paddingAll": pad, "paddingStart": "10px",
         "paddingEnd": "10px", "contents": cells}
    if bg:
        o["backgroundColor"] = bg
    return o


def _o2(tren, duoi, flex, mau=BLACK, dam=True, align="center", size="sm"):
    """Ô 2 dòng: số lớn + dòng nhỏ bên dưới (VD '1,50' / '/1,55')."""
    noi_dung = [_t(tren, size, mau, dam, align, wrap=False)]
    if duoi:
        noi_dung.append(_t(duoi, "xxs", GRAY, False, align, wrap=False))
    return {"type": "box", "layout": "vertical", "flex": flex, "contents": noi_dung}


def _hop_canh_bao(dong, nen=NEN_DO, mau=RED):
    return {"type": "box", "layout": "vertical", "backgroundColor": nen, "paddingAll": "8px",
            "paddingStart": "12px", "paddingEnd": "12px", "margin": "sm",
            "contents": [_t(x, "xs", mau, False) for x in dong]}


def _bubble(tieu_de, dong_phu, body):
    return {
        "type": "bubble", "size": "giga",
        "header": {"type": "box", "layout": "vertical", "backgroundColor": YELLOW, "paddingAll": "14px",
                   "contents": [_t(tieu_de, "lg", MTKM_GOLD_TEXT, True)] +
                               [_t(x, "xs", MTKM_GOLD_SUB, margin="xs") for x in dong_phu]},
        "body": {"type": "box", "layout": "vertical", "paddingAll": "0px", "paddingBottom": "12px",
                 "backgroundColor": PAGE_BG, "contents": body},
    }


def _dong_muc_tieu_cua_hang(ch):
    m = ch["muc"]
    d1 = " · ".join(f"{MUC[k]['ten']} {fmt_co_don_vi(m[k]['mt_ngay'], MUC[k]['don_vi'])}" for k in MUC_THU_NGAN)
    d2 = " · ".join(f"{MUC[k]['ten']} {fmt_co_don_vi(m[k]['mt_ngay'], MUC[k]['don_vi'])}" for k in MUC_FRESH)
    return [d1, d2]


# ---------------------------------------------------------------------------
# THẺ 1: MỤC TIÊU NGÀY TỪNG BẠN
# ---------------------------------------------------------------------------
def build_the_muc_tieu(ngay_str, ch, tb):
    """ch = muc_tieu_ngay.tinh_cua_hang(ngay); tb = muc_tieu_ngay.chia_tung_ban(ch, phan_line)."""
    body = [_sec("MỤC TIÊU NGÀY CỬA HÀNG")]
    body.append({"type": "box", "layout": "vertical", "paddingAll": "8px", "paddingStart": "12px",
                 "paddingEnd": "12px", "contents": [_t(x, "sm", MTKM_GOLD_TEXT, True, margin="xs")
                                                    for x in _dong_muc_tieu_cua_hang(ch)]})
    chua_dk = []
    for vai_tro, tieu_de, ds_muc in (("thu_ngan", "THU NGÂN", MUC_THU_NGAN), ("fresh", "FRESH", MUC_FRESH)):
        nguoi = tb.get(vai_tro) or []
        if not nguoi:
            continue
        body.append(_sec(tieu_de))
        nguoi = sorted(nguoi, key=lambda n: (0 if n["ca"][0] == "sang" else 1, len(n["ca"])))
        for j, n in enumerate(nguoi):
            mt = n["mt"]
            if vai_tro == "thu_ngan":
                phai = fmt_so(mt.get("sk", 0) + mt.get("dm", 0), "tr") + "tr"
                chi_tiet = (f"Sữa khô {fmt_so(mt.get('sk'), 'tr')} · Đông mát {fmt_so(mt.get('dm'), 'tr')} · "
                            f"C2 {fmt_so(mt.get('c2'), 'lốc')} lốc · 888: {fmt_so(mt.get('ng888'), 'túi')} túi")
            else:
                phai = f"Cá {fmt_so(mt.get('ca'), 'tr')}tr"
                chi_tiet = (f"Cá {fmt_so(mt.get('ca'), 'tr')}tr · Trứng {fmt_so(mt.get('trung'), 'vỉ')} vỉ · "
                            f"Heo C.P {fmt_so(mt.get('heo'), 'kg')}kg")
            if j:
                body.append({"type": "separator", "color": DIVIDER})
            body.append({"type": "box", "layout": "vertical", "paddingAll": "7px", "paddingStart": "12px",
                         "paddingEnd": "12px", "contents": [
                             {"type": "box", "layout": "horizontal", "contents": [
                                 {"type": "text", "text": n["ten"], "flex": 6, "wrap": True, "contents": [
                                     {"type": "span", "text": n["ten"], "weight": "bold", "color": MTKM_GOLD_TEXT, "size": "md"},
                                     {"type": "span", "text": f" · {ca_txt(n['ca'])}", "color": GRAY, "size": "xs"}]},
                                 _t(phai, "sm", RED, True, "end", 4)]},
                             _t(chi_tiet, "xs", "#555555", margin="xs")]})
            try:
                if not storage.get_user_id_da_dang_ky(n["ten"]) and n["ten"] not in chua_dk:
                    chua_dk.append(n["ten"])
            except Exception:
                pass
    fmcg = tb.get("fmcg") or {}
    ds_fmcg = [f"{t} ({'sáng' if ca == 'sang' else 'chiều'})" for ca in ("sang", "chieu") for t in fmcg.get(ca, [])]
    if ds_fmcg:
        body.append({"type": "box", "layout": "vertical", "paddingAll": "8px", "paddingStart": "12px",
                     "contents": [_t("FMCG: " + " · ".join(ds_fmcg) + " – không chia mục tiêu bán", "xs", GRAY)]})
    canh_bao = []
    if not (tb.get("thu_ngan") or tb.get("fresh")):
        canh_bao.append("Bot chưa đọc được tên bạn nào trong bài phân line – anh kiểm tra lại tiêu đề THU NGÂN / FRESH và tên giúp em")
    for t in chua_dk:
        canh_bao.append(f"{t} chưa đăng ký LINE với bot → {t} gõ \"DK {t}\" 1 lần để bot nhận báo cáo cuối ngày")
    if ch.get("ngay_thieu"):
        canh_bao.append(f"Chưa có file POS ngày {', '.join(ch['ngay_thieu'])} → mục tiêu ngày có thể lệch")
    if canh_bao:
        body.append(_hop_canh_bao(["⚠ " + x for x in canh_bao]))
    body.append({"type": "box", "layout": "vertical", "paddingStart": "12px", "paddingEnd": "12px", "paddingTop": "8px",
                 "contents": [_t("Cuối ngày (chậm nhất 21h45) các bạn nhắn số bán được vào nhóm · "
                                 "anh nạp file chốt rồi gõ CHỐT SỐ", "xxs", GRAY)]})
    return _bubble(f"MỤC TIÊU NGÀY · {_ngay_hien(ngay_str)}",
                   [f"Ngày {ch['ngay_thu']}/{ch['so_ngay']} · đơn vị: triệu",
                    "Ca sáng 40% · ca chiều 60% · chia đều trong ca · đã dồn phần thiếu/dư"],
                   body)


# ---------------------------------------------------------------------------
# THẺ 2: CHỐT SỐ CUỐI NGÀY
# ---------------------------------------------------------------------------
def _bang_cua_hang(ch):
    rows = [_hang([_t("Nhóm hàng", "xxs", MTKM_GOLD_TEXT, True, flex=34),
                   _t("MT ngày", "xxs", MTKM_GOLD_TEXT, True, "end", 22),
                   _t("POS", "xxs", MTKM_GOLD_TEXT, True, "end", 22),
                   _t("% đạt", "xxs", MTKM_GOLD_TEXT, True, "end", 22)], MTKM_HL_BG)]
    for i, k in enumerate(MUC_THU_NGAN + MUC_FRESH):
        m = ch["muc"][k]
        dv = MUC[k]["don_vi"]
        ten = MUC[k]["ten"] + ("" if dv == "tr" else f" ({dv})")
        phu = None
        if k == "trung" and m.get("tien") is not None:
            phu = f"tiền trứng {fmt_so(m['tien'], 'tr')}tr"
        elif k == "heo" and m.get("da_mo"):
            phu = f"trong đó da mỡ {fmt_so(m['da_mo'], 'kg')}kg"
        if phu:
            ten_cell = {"type": "box", "layout": "vertical", "flex": 34, "contents": [
                _t(ten, "sm", MTKM_GOLD_TEXT, True), _t(phu, "xxs", GRAY)]}
        else:
            ten_cell = _t(ten, "sm", MTKM_GOLD_TEXT, True, flex=34)
        p = m["pct"]
        rows.append({"type": "separator", "color": DIVIDER})
        rows.append(_hang([ten_cell,
                           _t(fmt_so(m["mt_ngay"], dv), "sm", BLACK, False, "end", 22),
                           _t(fmt_so(m["pos"], dv) if m["pos"] is not None else "–", "sm", BLACK, True, "end", 22),
                           _t((f"{p:.0f}%" + (" ✓" if p >= 100 else "")) if p is not None else "–", "sm",
                              _mau_pct(p), True, "end", 22)]))
    return {"type": "box", "layout": "vertical", "contents": rows}


def _bang_tung_ban(bang, ds_muc, ch):
    tieu_de_cot = {"sk": "Sữa khô", "dm": "Đông mát", "c2": "C2", "ng888": "888",
                   "ca": "Cá", "trung": "Trứng", "heo": "Heo C.P"}
    co = {"ten": 22, "muc": max(int(60 / len(ds_muc)), 14), "pct": 18}
    rows = [_hang([_t("Tên", "xxs", MTKM_GOLD_TEXT, True, flex=co["ten"])] +
                  [_t(tieu_de_cot[k], "xxs", MTKM_GOLD_TEXT, True, "center", co["muc"]) for k in ds_muc] +
                  [_t("% đạt", "xxs", MTKM_GOLD_TEXT, True, "end", co["pct"])], MTKM_HL_BG)]
    for d in bang["dong"]:
        cells = [{"type": "box", "layout": "vertical", "flex": co["ten"], "contents": [
            _t(d["ten"], "sm", MTKM_GOLD_TEXT, True, wrap=False),
            _t(d["ca"] + (f" · {d['gio']}" if d.get("gio") else ""), "xxs", GRAY, wrap=False)]}]
        for k in ds_muc:
            dv = MUC[k]["don_vi"]
            mt = d["mt"].get(k)
            if not d["da_bao"]:
                cells.append(_o2("–", f"/{fmt_so(mt, dv)}" if mt else "", co["muc"], GRAY, False))
                continue
            bc = d["bc"].get(k, 0)
            p = (bc / mt * 100) if mt else None
            cells.append(_o2(fmt_so(bc, dv), f"/{fmt_so(mt, dv)}" if mt else "", co["muc"], _mau_pct(p)))
        if d["da_bao"] and d["pct"] is not None:
            cells.append(_t(f"{d['pct']:.0f}%", "sm", _mau_pct(d["pct"]), True, "end", co["pct"]))
        else:
            cells.append(_t("chưa báo" if not d["da_bao"] else "–", "xxs", RED if not d["da_bao"] else GRAY,
                            True, "end", co["pct"]))
        rows.append({"type": "separator", "color": DIVIDER})
        rows.append(_hang(cells, pad="5px"))
    # tổng các bạn báo / POS / có tư vấn
    def _hang_tong(nhan, gia_tri, dam=False, mau=None):
        return _hang([_t(nhan, "xs", MTKM_GOLD_TEXT, True, flex=co["ten"])] +
                     [_t(gia_tri(k), "xs", (mau(k) if mau else BLACK), dam, "center", co["muc"]) for k in ds_muc] +
                     [{"type": "filler", "flex": co["pct"]}], NEN_TONG, pad="5px")
    rows.append({"type": "separator", "color": DIVIDER})
    rows.append(_hang_tong("Các bạn báo", lambda k: fmt_so(bang["tong_bc"][k], MUC[k]["don_vi"])))
    rows.append(_hang_tong("POS", lambda k: fmt_so(ch["muc"][k]["pos"], MUC[k]["don_vi"])
                           if ch["muc"][k]["pos"] is not None else "–", True))
    rows.append(_hang_tong("Có tư vấn",
                           lambda k: (f"{bang['co_tu_van'][k]:.0f}%" if bang["co_tu_van"][k] is not None else "–")
                           if bang["co_tu_van"][k] is None or bang["co_tu_van"][k] <= 100.5 else "⚠ vượt",
                           True, lambda k: RED if (bang["co_tu_van"][k] or 0) > 100.5 else _mau_pct(bang["co_tu_van"][k])))
    return {"type": "box", "layout": "vertical", "contents": rows}


def build_the_chot_so(ngay_str, kq):
    ch = kq["ch"]
    # Phần 1: mục tiêu ngày cửa hàng + MT ngày mai
    dau = [_sec("MỤC TIÊU NGÀY CỬA HÀNG"), _bang_cua_hang(ch)]
    mai = [f"{MUC[k]['ten']} {fmt_co_don_vi(ch['muc'][k]['mt_mai'], MUC[k]['don_vi'])}"
           for k in MUC_THU_NGAN + MUC_FRESH if ch["muc"][k]["mt_mai"] is not None]
    if mai:
        dau.append(_hop_canh_bao(["MT ngày mai (đã dồn phần thiếu/dư): " + " · ".join(mai)], NEN_XANH, CHU_XANH))
    # Phần 2: từng bạn
    tung_ban = []
    for vai_tro, tieu_de, ds_muc in (("thu_ngan", "TỪNG BẠN · THU NGÂN", MUC_THU_NGAN),
                                    ("fresh", "TỪNG BẠN · FRESH", MUC_FRESH)):
        bang = kq["bang"].get(vai_tro)
        if bang and bang["dong"]:
            tung_ban += [_sec(tieu_de), _bang_tung_ban(bang, ds_muc, ch)]
    # Phần 3: cảnh báo + nhận xét + chú thích
    canh_bao = []
    for b in kq["bao_vuot"]:
        dv = MUC[b["muc"]]["don_vi"]
        canh_bao.append(f"⚠ {MUC[b['muc']]['ten']}: các bạn báo {fmt_co_don_vi(b['tong'], dv)} nhưng POS chỉ "
                        f"{fmt_co_don_vi(b['pos'], dv)} → " +
                        (f"{', '.join(b['ten'])} báo cao hơn POS" if b["chac"] else f"kiểm tra {', '.join(b['ten'])}"))
    if kq["chua_bao"]:
        canh_bao.append("⚠ Chưa báo cáo: " + ", ".join(kq["chua_bao"]))
    if not kq["co_phan_line"]:
        canh_bao.append("⚠ Hôm nay chưa có bài phân line → chưa có mục tiêu từng bạn")
    elif not kq["so_nguoi"]:
        canh_bao.append("⚠ Bot chưa đọc được tên bạn nào trong bài phân line hôm nay → anh đăng lại bài "
                        "phân line rồi gõ CHỐT SỐ")
    if not ch["co_pos"]:
        canh_bao.append("⚠ Chưa có file POS 76 của hôm nay → anh nạp file chốt rồi gõ lại CHỐT SỐ")
    if ch.get("ngay_thieu"):
        canh_bao.append(f"⚠ Thiếu file POS ngày {', '.join(ch['ngay_thieu'])} → mục tiêu ngày có thể lệch")
    cuoi = [_hop_canh_bao(canh_bao)] if canh_bao else []
    cuoi.append({"type": "box", "layout": "vertical", "paddingAll": "10px", "paddingStart": "12px",
                 "paddingEnd": "12px", "contents": [{"type": "text", "text": "Nhận xét", "wrap": True, "size": "sm",
                                                     "contents": [
                     {"type": "span", "text": "Nhận xét: ", "weight": "bold", "color": MTKM_GOLD_TEXT},
                     {"type": "span", "text": kq["nhan_xet"], "color": "#333333"}]}]})
    cuoi.append({"type": "box", "layout": "vertical", "paddingStart": "12px", "paddingEnd": "12px",
                 "contents": [_t("% đạt từng bạn = trung bình % các mục được giao · xanh ≥100% · vàng 80–99% · "
                                 "đỏ <80% · Có tư vấn = các bạn báo / POS", "xxs", GRAY)]})
    tieu_de = f"CHỐT SỐ · {_ngay_hien(ngay_str)}"
    dong_phu = [f"Ngày {ch['ngay_thu']}/{ch['so_ngay']} · {kq['so_bao']}/{kq['so_nguoi']} bạn đã báo · đơn vị: triệu"]
    bubble = _bubble(tieu_de, dong_phu, dau + tung_ban + cuoi)
    # LINE giới hạn ~30KB/bubble: đông người quá thì tách phần TỪNG BẠN sang thẻ thứ 2 (vuốt ngang)
    if tung_ban and len(json.dumps(bubble, ensure_ascii=False, separators=(",", ":")).encode()) > 27000:
        return {"type": "carousel", "contents": [_bubble(tieu_de, dong_phu, dau + cuoi),
                                                 _bubble("TỪNG BẠN · " + _ngay_hien(ngay_str), dong_phu, tung_ban)]}
    return bubble
