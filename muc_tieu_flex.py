"""
muc_tieu_flex.py - Vẽ 2 thẻ Flex (01/11/2026):
  - build_the_muc_tieu(...)  : MỤC TIÊU NGÀY TỪNG BẠN (gửi ngay khi anh đăng bài phân line)
  - build_the_chot_so(...)   : CHỐT SỐ cuối ngày (lệnh "CHỐT SỐ")
Số liệu tính sẵn ở muc_tieu_ngay.py; file này CHỈ vẽ thẻ.
"""
import json
from datetime import datetime

import storage
from flex_builder import (YELLOW, RED, GRAY, GREEN, PAGE_BG, BLACK, VIEN_O,
                          MTKM_GOLD_TEXT, MTKM_GOLD_SUB, MTKM_HL_BG, MTKM_SEC_BG)
from muc_tieu_ngay import MUC, fmt_so, fmt_co_don_vi, ca_txt, pct_txt

VANG = "#C27C0E"
NHAN_THE = {"sk": "Sữa khô", "celano": "Celano", "dm": "Đông mát", "hmp": "HMP", "c2": "C2", "ng888": "888",
            "ca": "Cá", "trung": "Trứng", "heo": "Heo C.P"}
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
    o = {"type": "text", "text": str(text) if str(text).strip() else "–", "size": size}
    if color != BLACK:  # chữ đen là mặc định của LINE -> bỏ bớt cho nhẹ thẻ
        o["color"] = color
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


def _hang(cells, bg=None, pad="5px"):
    """(02/11/2026) Hàng bảng KẺ Ô như Excel: mỗi ô là 1 box có viền mỏng, số không dính nhau."""
    o_list = []
    for c in cells:
        if c.get("type") == "box":
            o = dict(c)
        else:
            c = dict(c)
            fl = c.pop("flex", None)
            if len(c.get("text", "")) <= 6:  # số ngắn (VD "11,28") không cần xuống dòng -> bỏ "wrap" cho nhẹ thẻ
                c.pop("wrap", None)
            o = {"type": "box", "layout": "vertical", "contents": [c]}
            if fl is not None:
                o["flex"] = fl
        o.update({"borderWidth": "light", "borderColor": VIEN_O, "paddingAll": pad})
        if len(o["contents"]) == 1:  # ô 1 dòng: canh giữa theo chiều dọc cho thẳng hàng với ô 2 dòng bên cạnh
            o["justifyContent"] = "center"
        if bg and "backgroundColor" not in o:
            o["backgroundColor"] = bg
        o_list.append(o)
    return {"type": "box", "layout": "horizontal", "contents": o_list}


def _o2(tren, duoi, flex, mau=BLACK, dam=True, align="end", size="sm"):
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


def _ten_muc(k):
    dv = MUC[k]["don_vi"]
    return MUC[k]["ten"] + ("" if dv == "tr" else f" ({dv})")


def _o_tieu_de(dong, flex, align="center"):
    """Ô tiêu đề cột: dòng 1 đậm, các dòng sau chữ nhỏ xám (VD tên bạn / ca / giờ báo)."""
    noi_dung = [_t(dong[0], "xs", MTKM_GOLD_TEXT, True, align)]  # cho xuống dòng: ngày đông người cột hẹp vẫn đủ chữ
    noi_dung += [_t(x, "xxs", GRAY, False, align) for x in dong[1:] if x]
    return {"type": "box", "layout": "vertical", "flex": flex, "contents": noi_dung}


def _bang_muc_tieu_vai_tro(ch, ds_muc, nguoi, vai_tro):
    """(02/11/2026) Bảng kẻ ô: hàng = mục, cột = Cả cửa hàng + từng bạn (đọc dọc theo tên mình)."""
    w_muc, w_ch, w_ng = 30, 22, 20
    cs = "sm" if len(nguoi) <= 4 else "xs"  # từ 5 bạn trở lên: chữ số nhỏ lại 1 cỡ để không bị cắt
    rows = [_hang([_o_tieu_de(["Mục"], w_muc, "start"), _o_tieu_de(["MT NGÀY"], w_ch, "end")] +
                  [_o_tieu_de([n["ten"], ca_txt(n["ca"])], w_ng, "end") for n in nguoi], MTKM_HL_BG, "4px")]
    for k in ds_muc:
        dv = MUC[k]["don_vi"]
        rows.append(_hang([_t(_ten_muc(k), "xs", MTKM_GOLD_TEXT, True, flex=w_muc),
                           _t(fmt_so(ch["muc"][k]["mt_ngay"], dv), cs, GRAY, False, "end", w_ch)] +
                          [_t(fmt_so(n["mt"][k], dv) if n["mt"].get(k) else "–", cs,
                              BLACK if n["mt"].get(k) else GRAY, bool(n["mt"].get(k)), "end", w_ng)
                           for n in nguoi], pad="4px"))
    if vai_tro == "thu_ngan":
        tien = [k for k in ds_muc if MUC[k]["don_vi"] == "tr"]
        tong_ch = sum(ch["muc"][k]["mt_ngay"] or 0 for k in tien)
        rows.append(_hang([_t("Tổng tiền", "xs", MTKM_GOLD_TEXT, True, flex=w_muc),
                           _t(fmt_so(tong_ch, "tr"), cs, GRAY, True, "end", w_ch)] +
                          [_t(fmt_so(sum(n["mt"].get(k, 0) for k in tien), "tr"), cs, RED, True, "end", w_ng)
                           for n in nguoi], NEN_TONG, "4px"))
    return {"type": "box", "layout": "vertical", "paddingStart": "8px", "paddingEnd": "8px", "contents": rows}


# ---------------------------------------------------------------------------
# THẺ 1: MỤC TIÊU NGÀY TỪNG BẠN
# ---------------------------------------------------------------------------
def build_the_muc_tieu(ngay_str, ch, tb):
    """ch = muc_tieu_ngay.tinh_cua_hang(ngay); tb = muc_tieu_ngay.chia_tung_ban(ch, phan_line)."""
    body = []
    chua_dk = []
    for vai_tro, tieu_de, ds_muc in (("thu_ngan", "THU NGÂN", ch["thu_ngan"]), ("fresh", "FRESH", ch["fresh"])):
        nguoi = sorted(tb.get(vai_tro) or [], key=lambda n: (0 if n["ca"][0] == "sang" else 1, len(n["ca"])))
        body += [_sec(tieu_de), _bang_muc_tieu_vai_tro(ch, ds_muc, nguoi, vai_tro)]
        for n in nguoi:
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
                   _t("THỰC TẾ", "xxs", MTKM_GOLD_TEXT, True, "end", 22),
                   _t("% đạt", "xxs", MTKM_GOLD_TEXT, True, "end", 22)], MTKM_HL_BG)]
    for i, k in enumerate(ch["thu_ngan"] + ch["fresh"]):
        m = ch["muc"][k]
        dv = MUC[k]["don_vi"]
        ten = MUC[k]["ten"] + ("" if dv == "tr" else f" ({dv})")
        phu = None
        if k == "trung" and m.get("tien") is not None:
            phu = f"tiền trứng {fmt_so(m['tien'], 'tr')}tr"
        elif k == "heo" and m.get("da_mo"):
            phu = f"trong đó da mỡ {fmt_so(m['da_mo'], 'kg')}kg"
        elif m.get("chu_ky"):
            ck = m["chu_ky"]
            mt_ck = f"{ck['muc_tieu_tien'] / 1_000_000:.3f}".rstrip("0").rstrip(".").replace(".", ",")
            phu = (f"thi đua {ck['tu'][8:10]}–{ck['den'][8:10]}/{ck['den'][5:7]}: "
                   f"{fmt_so(ck['tien'], 'tr')}/{mt_ck}tr")
        if phu:
            ten_cell = {"type": "box", "layout": "vertical", "flex": 34, "contents": [
                _t(ten, "sm", MTKM_GOLD_TEXT, True), _t(phu, "xxs", GRAY)]}
        else:
            ten_cell = _t(ten, "sm", MTKM_GOLD_TEXT, True, flex=34)
        p = m["pct"]
        rows.append(_hang([ten_cell,
                           _t(fmt_so(m["mt_ngay"], dv), "sm", BLACK, False, "end", 22),
                           _t(fmt_so(m["pos"], dv) if m["pos"] is not None else "–", "sm", BLACK, True, "end", 22),
                           _t(pct_txt(p) if p is not None else "–", "sm",
                              _mau_pct(p), True, "end", 22)]))
    return {"type": "box", "layout": "vertical", "paddingStart": "8px", "paddingEnd": "8px", "contents": rows}


def _bang_tung_ban(bang, ds_muc, ch):
    """(02/11/2026) 2 bảng kẻ ô:
       1) hàng = mục, cột = từng bạn: số báo (to) / mục tiêu (nhỏ), màu theo % · hàng cuối = % đạt
       2) đối chiếu: Các bạn báo | THỰC TẾ (POS) | Có tư vấn"""
    # cùng thứ tự với thẻ mục tiêu sáng: ca sáng -> cả ngày -> ca chiều -> báo ngoài line
    thu_tu = {"sáng": 0, "cả ngày": 1, "chiều": 2}
    dong = sorted(bang["dong"], key=lambda d: thu_tu.get(d["ca"], 3))
    w_muc, w_ng = 30, 22
    cs = "sm" if len(dong) <= 4 else "xs"
    rows = [_hang([_o_tieu_de(["Mục"], w_muc, "start")] +
                  [_o_tieu_de([d["ten"], d["ca"], (f"báo {d['gio']}" if d.get("gio") else "")], w_ng, "end") for d in dong],
                  MTKM_HL_BG, "4px")]
    for k in ds_muc:
        dv = MUC[k]["don_vi"]
        cells = [_t(_ten_muc(k), "xs", MTKM_GOLD_TEXT, True, flex=w_muc)]
        for d in dong:
            mt = d["mt"].get(k)
            if not d["da_bao"]:
                cells.append(_o2("–", f"/{fmt_so(mt, dv)}" if mt else "", w_ng, GRAY, False))
                continue
            bc = d["bc"].get(k, 0)
            if not mt and not bc:
                cells.append(_o2("–", "", w_ng, GRAY, False))
                continue
            p = (bc / mt * 100) if mt else None
            cells.append(_o2(fmt_so(bc, dv), f"/{fmt_so(mt, dv)}" if mt else "", w_ng, _mau_pct(p), size=cs))
        rows.append(_hang(cells, pad="4px"))
    hang_pct = [_t("% đạt", "xs", MTKM_GOLD_TEXT, True, flex=w_muc)]
    for d in dong:
        if d["da_bao"] and d["pct"] is not None:
            hang_pct.append(_t(pct_txt(d["pct"]), cs, _mau_pct(d["pct"]), True, "end", w_ng))
        else:
            hang_pct.append(_t("chưa báo" if not d["da_bao"] else "–", "xxs", RED if not d["da_bao"] else GRAY,
                               True, "end", w_ng))
    rows.append(_hang(hang_pct, NEN_TONG, "4px"))
    bang1 = {"type": "box", "layout": "vertical", "paddingStart": "8px", "paddingEnd": "8px", "contents": rows}

    # mục chưa giao cho ai và cũng không ai báo (VD HMP trước ngày áp dụng) -> để "–" cho gọn
    bo_trong = {k: not any(d["mt"].get(k) for d in dong) and not bang["tong_bc"][k] for k in ds_muc}
    w2 = (30, 23, 23, 24)
    rows2 = [_hang([_o_tieu_de(["Đối chiếu"], w2[0], "start"), _o_tieu_de(["Các bạn", "báo"], w2[1], "end"),
                    _o_tieu_de(["THỰC TẾ"], w2[2], "end"), _o_tieu_de(["Có tư vấn", "báo/thực tế"], w2[3], "end")],
                   MTKM_HL_BG, "4px")]
    for k in ds_muc:
        dv = MUC[k]["don_vi"]
        pos = ch["muc"][k]["pos"]
        ctv = bang["co_tu_van"][k]
        if bo_trong[k]:
            bao_txt, ctv_txt, ctv_mau = "–", "–", GRAY
        else:
            bao_txt = fmt_so(bang["tong_bc"][k], dv)
            if ctv is not None and ctv > 100.5:
                ctv_txt, ctv_mau = "⚠ BC KHỐNG", RED
            else:
                ctv_txt, ctv_mau = pct_txt(ctv), _mau_pct(ctv)
        rows2.append(_hang([_t(_ten_muc(k), "xs", MTKM_GOLD_TEXT, True, flex=w2[0]),
                            _t(bao_txt, "sm", BLACK, False, "end", w2[1]),
                            _t(fmt_so(pos, dv) if pos is not None else "–", "sm", BLACK, True, "end", w2[2]),
                            _t(ctv_txt, "sm", ctv_mau, True, "end", w2[3])], pad="4px"))
    bang2 = {"type": "box", "layout": "vertical", "paddingStart": "8px", "paddingEnd": "8px",
             "paddingTop": "8px", "contents": rows2}
    return [bang1, bang2]


def build_the_chot_so(ngay_str, kq):
    ch = kq["ch"]
    # Phần 1: mục tiêu ngày cửa hàng + MT ngày mai
    dau = [_sec("MỤC TIÊU NGÀY CỬA HÀNG"), _bang_cua_hang(ch)]
    mai = [f"{MUC[k]['ten']} {fmt_co_don_vi(ch['muc'][k]['mt_mai'], MUC[k]['don_vi'])}"
           for k in ch["thu_ngan"] + ch["fresh"] if ch["muc"][k]["mt_mai"] is not None]
    if mai:
        dau.append(_hop_canh_bao(["MT ngày mai (đã dồn phần thiếu/dư): " + " · ".join(mai)], NEN_XANH, CHU_XANH))
    # Phần 2: từng bạn
    phan_vai_tro = []  # [(tên vai trò, [khối])]
    for vai_tro, tieu_de, ds_muc in (("thu_ngan", "THU NGÂN", ch["thu_ngan"]),
                                    ("fresh", "FRESH", ch["fresh"])):
        bang = kq["bang"].get(vai_tro)
        if bang and bang["dong"]:
            phan_vai_tro.append((tieu_de, [_sec("TỪNG BẠN · " + tieu_de)] + _bang_tung_ban(bang, ds_muc, ch)))
    tung_ban = [x for _, khoi in phan_vai_tro for x in khoi]
    # Phần 3: cảnh báo + nhận xét + chú thích
    canh_bao = []
    for b in kq["bao_vuot"]:
        dv = MUC[b["muc"]]["don_vi"]
        canh_bao.append(f"⚠ BC KHỐNG {MUC[b['muc']]['ten']}: các bạn báo {fmt_co_don_vi(b['tong'], dv)} nhưng thực tế "
                        f"chỉ {fmt_co_don_vi(b['pos'], dv)} → " +
                        (f"{', '.join(b['ten'])}" if b["chac"] else f"kiểm tra {', '.join(b['ten'])}"))
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
                                 "đỏ <80% · Có tư vấn = các bạn báo / thực tế (POS)", "xxs", GRAY)]})
    tieu_de = f"CHỐT SỐ · {_ngay_hien(ngay_str)}"
    dong_phu = [f"Ngày {ch['ngay_thu']}/{ch['so_ngay']} · {kq['so_bao']}/{kq['so_nguoi']} bạn đã báo · đơn vị: triệu"]
    bubble = _bubble(tieu_de, dong_phu, dau + tung_ban + cuoi)
    # LINE giới hạn ~30KB/bubble: đông người quá thì tách phần TỪNG BẠN sang thẻ thứ 2 (vuốt ngang)
    if not tung_ban or _kb(bubble) <= 27000:
        return bubble
    ds = [_bubble(tieu_de, dong_phu, dau + cuoi)]
    gop = _bubble("TỪNG BẠN · " + _ngay_hien(ngay_str), dong_phu, tung_ban)
    if _kb(gop) <= 27000:
        ds.append(gop)
    else:  # đông người: mỗi vai trò 1 thẻ riêng (vuốt ngang)
        ds += [_bubble(f"{ten} · {_ngay_hien(ngay_str)}", dong_phu, khoi) for ten, khoi in phan_vai_tro]
    return {"type": "carousel", "contents": ds}


def _kb(o):
    return len(json.dumps(o, ensure_ascii=False, separators=(",", ":")).encode())
