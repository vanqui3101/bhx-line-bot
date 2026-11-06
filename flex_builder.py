"""
flex_builder.py - Dựng nội dung LINE Flex Message.
Theme: nền trắng/xám nhạt, header vàng nhạt (gold), chữ nhãn màu đen, số liệu màu đỏ.
Có 2 loại thẻ:
- build_flex_message(...) -> Báo cáo DOANH THU (Offline / Online / Tổng / Bill TB)
- build_category_flex_message(...) -> Báo cáo NGÀNH HÀNG (Nấm / C2 các loại / Nước giặt 888)
  [02/10/2026] Riêng thẻ NGÀNH HÀNG đổi theme sang "xanh nước biển" theo yêu
  cầu anh Quí — xem khối màu MTKM_* ngay phía trên build_category_flex_message.
  Các thẻ khác (DT, THƯỞNG, Thủy hải sản, Hủy tồn+MMKK, DTDK...) GIỮ NGUYÊN
  theme vàng/đỏ cũ.
"""
from datetime import datetime, timedelta
import math

# Bảng màu (08/09/2026: đã làm DỊU màu lại theo yêu cầu — trước đây nền quá
# xanh đậm + vàng chóe, giờ chuyển về nền trắng/xám rất nhạt, header vàng gold
# nhạt hơn, các khối nền phụ chỉ còn xám nhạt thay vì xanh dương bão hòa).
YELLOW = "#F6D365"  # header — vàng gold dịu (trước: #FFEE00 quá chói)
BLACK = "#1A1A1A"
RED = "#C0392B"  # số liệu — đỏ trầm hơn (trước: #D62020)
GRAY = "#666666"
GRAY_LIGHT = "#555555"
DIVIDER = "#E0E0E0"  # trước: #A8CCE0 (xanh dương)
ROW_BG = "#F5F5F5"  # trước: #D7ECFA (xanh dương đậm)
PAGE_BG = "#FFFFFF"  # nền thẻ — trắng (trước: #C8E6F9 xanh dương bão hòa)
GREEN = "#1E8246"

def _fmt_money(n):
    return f"{n:,.0f}".replace(",", ".")

def _fmt_int(n):
    return f"{n:,.0f}".replace(",", ".")

def _fmt_date_short(ngay):
    try:
        d = datetime.strptime(ngay, "%Y-%m-%d")
        return d.strftime("%d/%m")
    except Exception:
        return ngay or "—"

def _fmt_date_display(ngay):
    try:
        d = datetime.strptime(ngay, "%Y-%m-%d")
        return d.strftime("%d/%m/%Y")
    except Exception:
        return ngay

def _total_dt(rec):
    return (rec.get("dt_offline") or 0) + (rec.get("dt_online") or 0)

def _offline_dt(rec):
    return rec.get("dt_offline") or 0

def _online_dt(rec):
    return rec.get("dt_online") or 0

def _total_bill(rec):
    return (rec.get("bill_offline") or 0) + (rec.get("bill_online") or 0)

def _avg_bill_value(records):
    total_dt = sum(_total_dt(r) for r in records)
    total_bill = sum(_total_bill(r) for r in records)
    return (total_dt / total_bill) if total_bill else 0

def _pct_change(new, old):
    if not old:
        return None
    return (new - old) / old * 100

def _arrow(pct):
    if pct is None:
        return ""
    return "▲" if pct >= 0 else "▼"

def _delta_text(now, prev, pct):
    if prev is None:
        return "Chưa có dữ liệu tháng trước để so sánh"
    delta = now - prev
    sign = "+" if delta >= 0 else "-"
    arrow = _arrow(pct)
    return f"{sign}{_fmt_money(abs(delta))} đ ({arrow}{f'{abs(pct):.1f}'.replace('.', ',')}%)"

def _metric_block(label, now_val, prev_val, big=False):
    pct = _pct_change(now_val, prev_val) if prev_val is not None else None
    delta_str = _delta_text(now_val, prev_val, pct)
    return {
        "type": "box",
        "layout": "vertical",
        "margin": "lg",
        "contents": [
            {"type": "text", "text": label, "size": "xs", "weight": "bold", "color": BLACK},
            {
                "type": "box",
                "layout": "horizontal",
                "margin": "xs",
                "contents": [
                    {
                        "type": "text",
                        "text": f"{_fmt_money(now_val)} đ",
                        "size": "xl" if big else "lg",
                        "weight": "bold",
                        "color": RED,
                        "flex": 6,
                    },
                    {
                        "type": "text",
                        "text": f"{_fmt_money(prev_val)} đ" if prev_val is not None else "—",
                        "size": "sm",
                        "color": GRAY,
                        "align": "end",
                        "gravity": "bottom",
                        "flex": 5,
                    },
                ],
            },
            {
                "type": "text",
                "text": delta_str,
                "size": "xs",
                "weight": "bold",
                "color": RED,
                "align": "end",
                "margin": "xs",
                "wrap": True,
            },
        ],
    }

def build_flex_message(latest_date, latest_records, prev_date, prev_records, gio_now=None, gio_prev=None,
                        nganh_hang_breakdown=None, nganh_hang_ngay=None):
    total_now = sum(_total_dt(r) for r in latest_records)
    offline_now = sum(_offline_dt(r) for r in latest_records)
    online_now = sum(_online_dt(r) for r in latest_records)
    avg_bill_now = _avg_bill_value(latest_records)

    if prev_records:
        total_prev = sum(_total_dt(r) for r in prev_records)
        offline_prev = sum(_offline_dt(r) for r in prev_records)
        online_prev = sum(_online_dt(r) for r in prev_records)
        avg_bill_prev = _avg_bill_value(prev_records)
    else:
        total_prev = offline_prev = online_prev = avg_bill_prev = None

    ten_st = latest_records[0].get("ten_st") if latest_records else None
    subtitle_line2 = _fmt_date_display(latest_date)
    if prev_date:
        subtitle_line2 += f" · so với {_fmt_date_display(prev_date)} (cùng ngày tháng trước)"

    note_text = "Chưa có dữ liệu cùng ngày tháng trước để so sánh." if not prev_records else None

    body_contents = [
        {
            "type": "box",
            "layout": "vertical",
            "backgroundColor": ROW_BG,
            "cornerRadius": "12px",
            "paddingAll": "16px",
            "contents": [
                _metric_block("DOANH THU OFFLINE", offline_now, offline_prev),
                {"type": "separator", "margin": "lg", "color": DIVIDER},
                _metric_block("DOANH THU ONLINE", online_now, online_prev),
                {"type": "separator", "margin": "lg", "color": DIVIDER},
                _metric_block("TỔNG DOANH THU", total_now, total_prev, big=True),
                {"type": "separator", "margin": "lg", "color": DIVIDER},
                _metric_block("GIÁ TRỊ BILL TRUNG BÌNH", avg_bill_now, avg_bill_prev),
            ],
        },
    ]
    if note_text:
        body_contents.append(
            {"type": "text", "text": note_text, "size": "xxs", "color": GRAY, "margin": "md", "wrap": True}
        )

    if nganh_hang_breakdown:
        body_contents.append({"type": "separator", "margin": "xl", "color": DIVIDER})
        header_line = "DOANH THU THEO NGÀNH HÀNG"
        if nganh_hang_ngay:
            header_line += f" ({_fmt_date_display(nganh_hang_ngay)})"
        body_contents.append(
            {"type": "text", "text": header_line, "size": "sm", "weight": "bold", "color": BLACK, "margin": "xl"}
        )
        body_contents.append(
            {"type": "text", "text": "(tính từ file doanh thu chi tiết)", "size": "xxs", "color": GRAY, "margin": "xs"}
        )
        nh_rows = []
        for it in nganh_hang_breakdown:
            nh_rows.append({
                "type": "box",
                "layout": "horizontal",
                "margin": "sm",
                "contents": [
                    {"type": "text", "text": it["ten"], "size": "xs", "color": BLACK, "flex": 6, "wrap": True},
                    {"type": "text", "text": f"{_fmt_money(it['thanh_tien'])} đ", "size": "xs", "weight": "bold", "color": RED, "flex": 5, "align": "end"},
                ],
            })
        body_contents.extend(nh_rows)
        tong_nh = sum(it["thanh_tien"] for it in nganh_hang_breakdown)
        body_contents.append({"type": "separator", "margin": "md", "color": DIVIDER})
        body_contents.append({
            "type": "box",
            "layout": "horizontal",
            "margin": "sm",
            "contents": [
                {"type": "text", "text": "TỔNG", "size": "xs", "weight": "bold", "color": BLACK, "flex": 6},
                {"type": "text", "text": f"{_fmt_money(tong_nh)} đ", "size": "xs", "weight": "bold", "color": RED, "flex": 5, "align": "end"},
            ],
        })

    contents = {
        "type": "bubble",
        "size": "giga",
        "header": {
            "type": "box",
            "layout": "vertical",
            "backgroundColor": YELLOW,
            "paddingAll": "16px",
            "contents": [
                {"type": "text", "text": "BÁO CÁO DOANH THU", "color": BLACK, "weight": "bold", "size": "lg"},
                {"type": "text", "text": ten_st or "", "color": BLACK, "size": "sm", "margin": "sm", "wrap": True},
                {"type": "text", "text": subtitle_line2, "color": "#3D3200", "size": "xs", "margin": "xs", "wrap": True},
            ],
        },
        "body": {
            "type": "box",
            "layout": "vertical",
            "paddingAll": "16px",
            "backgroundColor": PAGE_BG,
            "contents": body_contents,
        },
    }
    return contents

# ---------------------------------------------------------------------------
# BÁO CÁO NGÀNH HÀNG (Nấm / C2 các loại / Nước giặt 888)
# [02/10/2026] Bỏ BÁNH TRUNG THU (hết mùa), thêm NƯỚC GIẶT 888 (kèm dòng dự
# kiến đến hết tháng). Đổi theme riêng cho thẻ này sang "xanh nước biển".
# ---------------------------------------------------------------------------
TABLE_HEAD_BG = "#EFEFEF"  # trước: #AAD2F0 (xanh dương)
ROW_ALT_BG = "#FAFAFA"  # trước: #DEF0FB (xanh dương nhạt)

# Màu riêng cho thẻ NGÀNH HÀNG (MTKM) — "xanh nước biển" theo yêu cầu anh
# Quí 02/10/2026. CHỈ dùng trong build_category_flex_message, không ảnh
# hưởng theme vàng/đỏ của các thẻ báo cáo khác.
MTKM_HEADER_BG = "#0B3C66"
MTKM_SUBTITLE = "#CFE3F2"
MTKM_VALUE = "#0B3C66"

def _table_header_row():
    return {
        "type": "box",
        "layout": "horizontal",
        "backgroundColor": TABLE_HEAD_BG,
        "paddingAll": "8px",
        "margin": "sm",
        "contents": [
            {"type": "text", "text": "Sản phẩm", "size": "md", "weight": "bold", "color": BLACK, "flex": 6},
            {"type": "text", "text": "SL bán", "size": "md", "weight": "bold", "color": BLACK, "flex": 3},
            {"type": "text", "text": "% bán/nhập", "size": "md", "weight": "bold", "color": BLACK, "flex": 4, "align": "end"},
        ],
    }

def _table_data_row(ten, qty_text, pct_value, alt_bg):
    pct_text = f"{pct_value:.1f}%" if pct_value is not None else "—"
    row = {
        "type": "box",
        "layout": "horizontal",
        "paddingAll": "8px",
        "contents": [
            {"type": "text", "text": ten, "size": "md", "color": BLACK, "flex": 6, "wrap": True},
            {"type": "text", "text": qty_text, "size": "md", "weight": "bold", "color": RED, "flex": 3, "wrap": True},
            {"type": "text", "text": pct_text, "size": "md", "weight": "bold", "color": RED, "flex": 4, "align": "end"},
        ],
    }
    if alt_bg:
        row["backgroundColor"] = ROW_ALT_BG
    return row

def _category_section_table(title, note, items, empty_note):
    """Mục có bảng sản phẩm 3 cột (Sản phẩm / SL bán / % bán nhập) — dùng
    cho báo cáo Doanh thu thủy hải sản. KHÔNG dùng cho MTKM nữa (MTKM dùng
    _mtkm_section_table bên dưới, chỉ 2 cột)."""
    contents = [{"type": "text", "text": title, "size": "xl", "weight": "bold", "color": BLACK}]
    if note:
        contents.append({"type": "text", "text": note, "size": "lg", "color": GRAY, "margin": "xs", "wrap": True})
    contents.append(_table_header_row())
    if not items:
        contents.append({"type": "text", "text": empty_note, "size": "lg", "color": GRAY, "margin": "sm", "wrap": True})
    else:
        for i, (ten, qty_text, pct_value) in enumerate(items):
            contents.append(_table_data_row(ten, qty_text, pct_value, alt_bg=(i % 2 == 1)))
    return {
        "type": "box",
        "layout": "vertical",
        "margin": "lg",
        "contents": contents,
    }

def _category_section_simple(title, total_label, total_value_text, value_color=RED):
    """Mục chỉ có 1 dòng tổng (Nấm)."""
    return {
        "type": "box",
        "layout": "vertical",
        "margin": "lg",
        "contents": [
            {"type": "text", "text": title, "size": "xl", "weight": "bold", "color": BLACK},
            {
                "type": "text",
                "text": f"{total_label}: {total_value_text}",
                "size": "lg",
                "weight": "bold",
                "color": value_color,
                "align": "end",
                "margin": "md",
            },
        ],
    }

# ---- Helper riêng cho MTKM: bảng 2 cột (Sản phẩm / SL bán), không có cột
# % bán/nhập — và dòng "dự kiến đến hết tháng" cho Nước giặt 888. ----

def _mtkm_table_header_row():
    return {
        "type": "box",
        "layout": "horizontal",
        "backgroundColor": TABLE_HEAD_BG,
        "paddingAll": "8px",
        "margin": "sm",
        "contents": [
            {"type": "text", "text": "Sản phẩm", "size": "md", "weight": "bold", "color": BLACK, "flex": 7},
            {"type": "text", "text": "SL bán", "size": "md", "weight": "bold", "color": BLACK, "flex": 3, "align": "end"},
        ],
    }

def _mtkm_table_data_row(ten, qty_text, alt_bg):
    row = {
        "type": "box",
        "layout": "horizontal",
        "paddingAll": "8px",
        "contents": [
            {"type": "text", "text": ten, "size": "md", "color": BLACK, "flex": 7, "wrap": True},
            {"type": "text", "text": qty_text, "size": "md", "weight": "bold", "color": MTKM_VALUE, "flex": 3, "align": "end", "wrap": True},
        ],
    }
    if alt_bg:
        row["backgroundColor"] = ROW_ALT_BG
    return row

def _mtkm_section_table(title, note, items, empty_note):
    """items: list[(ten, qty_text)] — dùng cho C2 CÁC LOẠI / NƯỚC GIẶT 888,
    chỉ 2 cột (Sản phẩm, SL bán), không còn cột % bán/nhập."""
    contents = [{"type": "text", "text": title, "size": "xl", "weight": "bold", "color": BLACK}]
    if note:
        contents.append({"type": "text", "text": note, "size": "sm", "color": GRAY, "margin": "xs", "wrap": True})
    contents.append(_mtkm_table_header_row())
    if not items:
        contents.append({"type": "text", "text": empty_note, "size": "lg", "color": GRAY, "margin": "sm", "wrap": True})
    else:
        for i, (ten, qty_text) in enumerate(items):
            contents.append(_mtkm_table_data_row(ten, qty_text, alt_bg=(i % 2 == 1)))
    return {
        "type": "box",
        "layout": "vertical",
        "margin": "lg",
        "contents": contents,
    }

def _mtkm_forecast_row(label, value_text):
    return {
        "type": "box",
        "layout": "horizontal",
        "margin": "sm",
        "paddingTop": "4px",
        "contents": [
            {"type": "text", "text": label, "size": "sm", "color": GRAY, "flex": 7, "wrap": True},
            {"type": "text", "text": value_text, "size": "md", "weight": "bold", "color": MTKM_VALUE, "flex": 3, "align": "end", "wrap": True},
        ],
    }

# ---------------------------------------------------------------------------
# THẺ MTKM BẢN MỚI (05/10/2026) — 3 khối gọn: NẤM / THI ĐUA TUẦN / NƯỚC GIẶT 888
# Có LŨY KẾ + MỤC TIÊU NGÀY (phần thiếu tự dồn sang các ngày sau). Số liệu
# tính sẵn ở mtkm_tracker.tinh_mtkm(); hàm dưới đây CHỈ vẽ thẻ.
# Màu theo đúng bản demo anh Quí duyệt: header vàng gold, nền trắng, khối xám nhạt.
# ---------------------------------------------------------------------------
MTKM_GOLD_TEXT = "#412402"
MTKM_GOLD_SUB = "#633806"
MTKM_BAR = "#BA7517"
MTKM_HL_BG = "#FAEEDA"
MTKM_BOX_BG = "#F5F5F5"


def _tr(n, so_le=1):
    """Đồng -> 'x,y tr'."""
    return f"{(n or 0) / 1_000_000:.{so_le}f}".replace(".", ",") + " tr"


def _mtkm_row(label, value_text, value_color=BLACK, bold=False):
    return {
        "type": "box", "layout": "horizontal", "margin": "sm",
        "contents": [
            {"type": "text", "text": label, "size": "sm", "color": GRAY, "flex": 6, "wrap": True},
            {"type": "text", "text": value_text, "size": "sm", "color": value_color,
             "weight": "bold" if bold else "regular", "flex": 5, "align": "end", "wrap": True},
        ],
    }


def _mtkm_bar(pct):
    w = max(1, min(100, int(round(pct or 0))))
    return {
        "type": "box", "layout": "vertical", "margin": "md", "height": "6px",
        "backgroundColor": DIVIDER, "cornerRadius": "3px",
        "contents": [{
            "type": "box", "layout": "vertical", "width": f"{w}%", "height": "6px",
            "backgroundColor": MTKM_BAR, "cornerRadius": "3px",
            "contents": [{"type": "filler"}],
        }],
    }


def _mtkm_khoi(icon_title, sub_title, so, fmt, don_vi_con_lai, ghi_chu=None):
    """Vẽ 1 khối (Nấm / Tuần / 888). fmt: hàm định dạng số."""
    thieu = so["con_thieu_hom_nay"]
    thieu_text = f"-{fmt(thieu)}" if thieu > 0 else "Đã đủ"
    thieu_color = RED if thieu > 0 else GREEN

    o = [
        _mtkm_row("Hôm nay", f"{fmt(so['hom_nay'])} / {fmt(so['muc_tieu_hom_nay'])}", BLACK, True),
        _mtkm_row("Còn thiếu hôm nay", thieu_text, thieu_color, True),
        {"type": "separator", "margin": "md", "color": DIVIDER},
        _mtkm_row(don_vi_con_lai["luy_ke_label"], f"{fmt(so['luy_ke'])} / {fmt(so['muc_tieu'])}", BLACK),
        _mtkm_bar(so["pct"]),
        {
            "type": "box", "layout": "horizontal", "margin": "sm",
            "contents": [
                {"type": "text", "text": f"{so['pct']:.0f}%", "size": "xs", "color": GRAY, "flex": 3},
                {"type": "text",
                 "text": (f"Còn {fmt(so['con_lai'])} / {so['so_ngay_sau_hom_nay']} ngày"
                          if so["so_ngay_sau_hom_nay"] > 0 else "Ngày cuối kỳ"),
                 "size": "xs", "color": GRAY, "flex": 7, "align": "end", "wrap": True},
            ],
        },
    ]
    khoi = [
        {"type": "text", "text": icon_title, "size": "md", "weight": "bold", "color": BLACK, "wrap": True},
        {"type": "text", "text": sub_title, "size": "xs", "color": GRAY, "margin": "xs", "wrap": True},
        {"type": "box", "layout": "vertical", "margin": "md", "paddingAll": "10px",
         "backgroundColor": MTKM_BOX_BG, "cornerRadius": "8px", "contents": o},
    ]
    if so["mai_can"] is not None:
        khoi.append({
            "type": "box", "layout": "horizontal", "margin": "sm", "paddingAll": "8px",
            "backgroundColor": MTKM_HL_BG, "cornerRadius": "8px",
            "contents": [
                {"type": "text", "text": "Mai cần (nếu chốt số này)", "size": "sm", "color": MTKM_GOLD_SUB, "flex": 6, "wrap": True},
                {"type": "text", "text": fmt(so["mai_can"]), "size": "md", "weight": "bold",
                 "color": MTKM_GOLD_TEXT, "flex": 4, "align": "end"},
            ],
        })
    if ghi_chu:
        khoi.append({"type": "text", "text": ghi_chu, "size": "xs", "color": RED, "margin": "sm", "wrap": True})
    return {"type": "box", "layout": "vertical", "margin": "lg", "contents": khoi}


# (20/10/2026) THẺ MTKM DẠNG BẢNG DỌC — 2 phần HÔM NAY / LŨY KẾ THÁNG, cùng
# kiểu cột: Ngành | Mục tiêu | DT hiện tại | % đạt. Số liệu tính sẵn ở
# mtkm_tracker.tinh_mtkm(); hàm dưới đây CHỈ vẽ thẻ. Đơn vị: triệu đồng.
MTKM_SEC_BG = "#BA7517"
VIEN_O = "#CFCFCF"  # (02/11/2026) màu viền ô bảng kẻ ô
MTKM_TONG_BG = "#FFF8EC"


def _so_tr(n, so_le=2):
    return f"{(n or 0) / 1_000_000:.{so_le}f}".replace(".", ",")


def _mtkm_sec(title, sub=None):
    contents = [{"type": "text", "text": title, "size": "sm", "weight": "bold", "color": "#FFFFFF", "flex": 0}]
    if sub:
        contents.append({"type": "text", "text": sub, "size": "xs", "color": "#FFFFFF", "margin": "sm",
                         "gravity": "center", "wrap": True})
    return {"type": "box", "layout": "horizontal", "backgroundColor": MTKM_SEC_BG,
            "paddingAll": "6px", "paddingStart": "12px", "contents": contents}


def _mtkm_bang_hang(cot, header=False, bg=None, mau_pct=BLACK, dam=False):
    """1 hàng bảng: cot = [ngành, mục tiêu, DT hiện tại, % đạt]."""
    flex = [5, 3, 3, 3]
    o = []
    for i, txt in enumerate(cot):
        o.append({
            "type": "text", "text": txt, "flex": flex[i], "wrap": True,
            "size": "xs" if header else "sm",
            "align": "start" if i == 0 else "end",
            "weight": "bold" if (header or i == 0 or i == 3 or dam) else "regular",
            "color": MTKM_GOLD_TEXT if (header or i == 0) else (mau_pct if i == 3 else BLACK),
            "gravity": "center",
        })
    box = {"type": "box", "layout": "horizontal", "paddingAll": "7px", "paddingStart": "12px",
           "paddingEnd": "12px", "contents": o}
    if bg:
        box["backgroundColor"] = bg
    return box


def _mtkm_bang(title, sub, ten_cot_mt, dong, lay):
    """lay(s) -> (mục tiêu, DT, %, đạt?)"""
    rows = [_mtkm_sec(title, sub),
            _mtkm_bang_hang(["Ngành", ten_cot_mt, "DT hiện tại", "% đạt"], header=True, bg=MTKM_HL_BG)]
    for i, s in enumerate(dong):
        mt, dt, pct, dat = lay(s)
        pct_txt = f"{pct:.0f}%" if title.startswith("HÔM") else f"{pct:.1f}%".replace(".", ",")
        if i:
            rows.append({"type": "separator", "color": DIVIDER})
        rows.append(_mtkm_bang_hang(
            [s["ten"], _so_tr(mt, 2 if title.startswith("HÔM") else 1), _so_tr(dt, 2 if title.startswith("HÔM") else 1), pct_txt],
            bg=MTKM_TONG_BG if s.get("la_tong") else None,
            mau_pct=GREEN if dat else RED, dam=s.get("la_tong", False),
        ))
    return {"type": "box", "layout": "vertical", "contents": rows}


def _mtkm_so(v, don_vi, chi_tiet=False):
    """Định dạng số trên thẻ MTKM: tiền -> triệu (2 số lẻ), kg -> 1 số lẻ, túi -> số nguyên."""
    if don_vi == "tr":
        return f"{(v or 0) / 1_000_000:.2f}".replace(".", ",")
    if don_vi == "kg":
        return f"{(v or 0):.1f}".replace(".", ",")
    return f"{(v or 0):.0f}"


def _mtkm_o(text, flex, align="end", color=BLACK, bold=False, size="sm"):
    return {"type": "text", "text": text, "flex": flex, "size": size, "align": align, "color": color,
            "weight": "bold" if bold else "regular", "wrap": True, "gravity": "center"}


# (01/11/2026) THẺ MTKM THÁNG 11 — 1 bảng gọn (MT ngày | Hôm nay | % ngày | % tháng)
# + phần MỤC TIÊU TỪNG BẠN (đọc bài phân line, chia ca sáng 40% / chiều 60%).
# Số liệu tính sẵn ở mtkm_tracker.tinh_mtkm(); hàm dưới đây CHỈ vẽ thẻ.
def build_category_flex_message(ngay, ten_st, payload, kq=None, gio=None):
    kq = kq or {}
    dong = kq.get("dong") or []
    tien_do = kq.get("pct_tien_do", 0)
    body = []
    if dong:
        # (02/11/2026) Bảng KẺ Ô như Excel cho dễ nhìn: mỗi ô có viền, số không dính nhau
        cot = (32, 15, 16, 18, 19)

        def _o(noi_dung, flex, align="end", color=BLACK, bold=False, size="sm", bg=None):
            if isinstance(noi_dung, str):
                noi_dung = [{"type": "text", "text": noi_dung, "size": size, "align": align, "color": color,
                             "weight": "bold" if bold else "regular", "wrap": True}]
            o = {"type": "box", "layout": "vertical", "flex": flex, "justifyContent": "center",
                 "borderWidth": "light", "borderColor": VIEN_O, "paddingAll": "5px", "contents": noi_dung}
            if bg:
                o["backgroundColor"] = bg
            return o

        def _hang_o(cells):
            return {"type": "box", "layout": "horizontal", "contents": cells}

        rows = [_hang_o([_o("Nhóm hàng", cot[0], "start", MTKM_GOLD_TEXT, True, "xxs", MTKM_HL_BG),
                         _o("MT ngày", cot[1], "center", MTKM_GOLD_TEXT, True, "xxs", MTKM_HL_BG),
                         _o("Hôm nay", cot[2], "center", MTKM_GOLD_TEXT, True, "xxs", MTKM_HL_BG),
                         _o("% ngày", cot[3], "center", MTKM_GOLD_TEXT, True, "xxs", MTKM_HL_BG),
                         _o("% tháng", cot[4], "center", MTKM_GOLD_TEXT, True, "xxs", MTKM_HL_BG)])]
        da_ghi_nhom_dem = False
        for d in dong:
            if d["don_vi"] in ("loc", "tui") and not da_ghi_nhom_dem:
                da_ghi_nhom_dem = True
                rows.append(_hang_o([_o("ĐẾM SỐ LƯỢNG (lốc / túi)", 100, "start", "#FFFFFF", True, "xxs", MTKM_SEC_BG)]))
            dat_ngay = d["pct_ngay"] >= 100
            dat_thang = d["pct_thang"] >= (d["tien_do_rieng"] if d.get("tien_do_rieng") is not None else tien_do)
            bg = "#FFF8EC" if d["la_tong"] else None
            ten = [{"type": "text", "text": d["ten"], "size": "sm", "weight": "bold", "color": MTKM_GOLD_TEXT, "wrap": True}]
            if d.get("nhan_gd"):
                ten.append({"type": "text", "text": d["nhan_gd"], "size": "xxs", "color": GRAY, "wrap": True})
            pct_ngay = math.floor(d["pct_ngay"]) if d["pct_ngay"] < 100 else round(d["pct_ngay"])
            rows.append(_hang_o([
                _o(ten, cot[0], bg=bg),
                _o(_mtkm_so(d["mt_ngay"], d["don_vi"]), cot[1], bold=d["la_tong"], bg=bg),
                _o(_mtkm_so(d["hom_nay"], d["don_vi"]), cot[2], bold=True, bg=bg),
                _o(f"{pct_ngay:.0f}%", cot[3], color=GREEN if dat_ngay else RED,
                   bold=True, bg=bg),
                _o(f"{d['pct_thang']:.1f}%".replace(".", ","), cot[4],
                   color=GREEN if dat_thang else RED, bold=True, bg=bg),
            ]))
        body.append({"type": "box", "layout": "vertical", "paddingStart": "8px", "paddingEnd": "8px",
                     "paddingTop": "8px", "contents": rows})

        # ---- MỤC TIÊU TỪNG BẠN ----
        tung_ban = kq.get("tung_ban") or []
        if tung_ban:
            body.append({"type": "box", "layout": "vertical", "backgroundColor": MTKM_SEC_BG,
                         "paddingAll": "6px", "paddingStart": "12px", "margin": "md",
                         "contents": [{"type": "text", "text": "MỤC TIÊU TỪNG BẠN HÔM NAY", "size": "sm",
                                       "weight": "bold", "color": "#FFFFFF"}]})
            nhan = {"sk": "Sữa khô", "sm": "Sữa mát", "kem": "Kem", "dl": "Đông lạnh",
                    "ng888": "888", "ca_trung": "Cá + trứng", "heo": "Heo C.P"}
            for j, b in enumerate(tung_ban):
                mt = b["mt"]
                if b["loai"] == "thu_ngan":
                    tong = sum(v for k, v in mt.items() if k in ("sk", "sm", "kem", "dl"))
                    tong_txt = _mtkm_so(tong, "tr") + "tr"
                    chi_tiet = " · ".join(f"{nhan[k]} {_mtkm_so(mt[k], 'tr')}" for k in ("sk", "sm", "kem", "dl") if k in mt)
                    if "ng888" in mt:
                        chi_tiet += f" · 888: {int(mt['ng888'])} túi"
                else:
                    tong_txt = (_mtkm_so(mt.get("ca_trung", 0), "tr") + "tr + "
                                + _mtkm_so(mt.get("heo", 0), "kg") + "kg")
                    chi_tiet = (f"Cá + trứng {_mtkm_so(mt.get('ca_trung', 0), 'tr')}tr · "
                                f"Heo C.P {_mtkm_so(mt.get('heo', 0), 'kg')}kg")
                if j:
                    body.append({"type": "separator", "color": DIVIDER})
                body.append({
                    "type": "box", "layout": "vertical", "paddingAll": "7px",
                    "paddingStart": "12px", "paddingEnd": "12px",
                    "contents": [
                        {"type": "box", "layout": "horizontal", "contents": [
                            {"type": "text", "text": f"{b['ten']} · {b['vai_tro']}", "size": "sm", "weight": "bold",
                             "color": MTKM_GOLD_TEXT, "flex": 6, "wrap": True},
                            {"type": "text", "text": tong_txt, "size": "sm", "weight": "bold", "color": RED,
                             "flex": 4, "align": "end", "wrap": True}]},
                        {"type": "text", "text": chi_tiet, "size": "xs", "color": GRAY_LIGHT, "wrap": True, "margin": "xs"},
                    ],
                })
    else:
        body.append({"type": "text", "text": "Chưa có mục tiêu MTKM cho tháng này.", "size": "sm",
                     "color": GRAY, "margin": "md", "wrap": True})

    ghi = ["Đơn vị: triệu đồng (heo: kg, trứng: vỉ, Celano/C2: lốc, 888: túi) · MT ngày đã dồn phần thiếu · "
           "Celano: % là % của cả chu kỳ thi đua"]
    ghi_do = list(kq.get("ghi_chu") or [])
    if kq.get("ngay_thieu"):
        ghi_do.append(f"⚠ Chưa có dữ liệu ngày {', '.join(kq['ngay_thieu'])} (đang tính = 0)")
    ghi_box = [{"type": "text", "text": g, "size": "xxs", "color": GRAY, "wrap": True} for g in ghi]
    ghi_box += [{"type": "text", "text": g, "size": "xxs", "color": RED, "wrap": True, "margin": "xs"} for g in ghi_do]
    body.append({"type": "box", "layout": "vertical", "paddingStart": "12px", "paddingEnd": "12px",
                 "paddingTop": "8px", "contents": ghi_box})

    dong_ngay = _fmt_date_display(ngay) + (f" · cập nhật {gio}" if gio else "")
    tien_do_txt = (f"Ngày {kq.get('ngay_thu', '')}/{kq.get('so_ngay_thang', '')} · "
                   f"tiến độ tháng cần {tien_do:.1f}%".replace(".", ","))
    return {
        "type": "bubble",
        "size": "giga",
        "header": {
            "type": "box", "layout": "vertical", "backgroundColor": YELLOW, "paddingAll": "14px",
            "contents": [
                {"type": "text", "text": f"MTKM {ten_st or ''}".strip(), "color": MTKM_GOLD_TEXT, "weight": "bold", "size": "lg", "wrap": True},
                {"type": "text", "text": dong_ngay, "color": MTKM_GOLD_SUB, "size": "xs", "margin": "xs"},
                {"type": "text", "text": tien_do_txt, "color": MTKM_GOLD_SUB, "size": "xs", "margin": "xs"},
            ],
        },
        "body": {
            "type": "box", "layout": "vertical", "paddingAll": "0px", "paddingBottom": "12px",
            "backgroundColor": PAGE_BG, "contents": body,
        },
    }

# ---------------------------------------------------------------------------
# BÁO CÁO THƯỞNG (FRESH + FMCG) — lệnh "TD" / "THƯỞNG"
# ---------------------------------------------------------------------------
def _thuong_table_header():
    return {
        "type": "box", "layout": "horizontal", "backgroundColor": TABLE_HEAD_BG,
        "paddingAll": "6px", "margin": "sm",
        "contents": [
            {"type": "text", "text": "Chỉ tiêu", "size": "sm", "flex": 4},
            {"type": "text", "text": "Base", "size": "sm", "weight": "bold", "color": BLACK, "flex": 3, "align": "center"},
            {"type": "text", "text": "Thực tế", "size": "sm", "weight": "bold", "color": BLACK, "flex": 3, "align": "center"},
            {"type": "text", "text": "Dự kiến", "size": "sm", "weight": "bold", "color": BLACK, "flex": 3, "align": "center"},
        ],
    }

def _thuong_table_row(label, base_text, thuc_te_text, du_kien_text):
    return {
        "type": "box", "layout": "horizontal", "paddingAll": "6px",
        "contents": [
            {"type": "text", "text": label, "size": "sm", "color": BLACK, "flex": 4, "wrap": True},
            {"type": "text", "text": base_text, "size": "sm", "color": BLACK, "flex": 3, "align": "center", "wrap": True},
            {"type": "text", "text": thuc_te_text, "size": "sm", "weight": "bold", "color": RED, "flex": 3, "align": "center", "wrap": True},
            {"type": "text", "text": du_kien_text, "size": "sm", "weight": "bold", "color": GREEN, "flex": 3, "align": "center", "wrap": True},
        ],
    }

def _thuong_muc_box(label, gia_tri):
    return {
        "type": "box", "layout": "horizontal", "backgroundColor": TABLE_HEAD_BG,
        "paddingAll": "10px", "margin": "md", "cornerRadius": "8px",
        "contents": [
            {"type": "text", "text": label, "size": "sm", "weight": "bold", "color": BLACK, "flex": 6, "wrap": True},
            {"type": "text", "text": gia_tri, "size": "md", "weight": "bold", "color": GREEN, "flex": 4, "align": "end"},
        ],
    }

def _thuong_section_title(title, note):
    contents = [{"type": "text", "text": title, "size": "lg", "weight": "bold", "color": BLACK}]
    if note:
        contents.append({"type": "text", "text": note, "size": "xs", "color": GRAY, "margin": "xs", "wrap": True})
    return contents

def build_thuong_flex_message(ten_st, payload):
    ngay_bd = _fmt_date_display(payload["ngay_bat_dau"])
    ngay_kt = _fmt_date_display(payload["ngay_ket_thuc"])
    so_ngay = payload["so_ngay_da_qua"]
    so_ngay_thang = payload["so_ngay_ca_thang"]

    fresh = payload["fresh"]
    skdm = payload["skdm"]
    bianuoc = payload["bianuoc"]
    c2 = payload["c2"]

    body_contents = []

    # ---- FRESH ----
    body_contents.extend(_thuong_section_title("FRESH — Thịt heo/gà nhập khẩu", "So với TB tháng 5-6/2026"))
    body_contents.append(_thuong_table_header())
    body_contents.append(_thuong_table_row(
        "Sản lượng", f"{fresh['base_kg']:,.1f} kg".replace(",", "."),
        f"{fresh['thuc_te_kg']:,.1f} kg".replace(",", "."),
        f"{fresh['du_kien_kg']:,.1f} kg".replace(",", "."),
    ))
    body_contents.append(_thuong_table_row(
        "Tăng trưởng", "—", "—",
        f"{fresh['pct']:+.1f}%".replace(".", ","),
    ))
    body_contents.append(_thuong_muc_box(f"Mức thưởng: {fresh['muc']}", f"{_fmt_money(fresh['thuong_du_kien'])} đ"))
    body_contents.append({"type": "separator", "margin": "xl", "color": DIVIDER})

    # ---- SKDM ----
    body_contents.extend(_thuong_section_title("Sữa - Kem - Đông - Mát", "So với tháng 7/2026"))
    body_contents.append(_thuong_table_header())
    body_contents.append(_thuong_table_row(
        "Doanh thu", f"{_fmt_money(skdm['base_tien'])} đ",
        f"{_fmt_money(skdm['thuc_te_tien'])} đ",
        f"{_fmt_money(skdm['du_kien_tien'])} đ",
    ))
    body_contents.append(_thuong_table_row(
        "Tăng trưởng", "—", "—",
        f"{skdm['pct']:+.1f}%".replace(".", ","),
    ))
    body_contents.append(_thuong_muc_box(f"Mức thưởng: {skdm['muc']}", f"{_fmt_money(skdm['thuong_du_kien'])} đ"))
    body_contents.append({"type": "separator", "margin": "xl", "color": DIVIDER})

    # ---- BIA-NUOC ----
    body_contents.extend(_thuong_section_title("Bia - Nước", "So với tháng 6/2026 · Size ST < 2 tỷ"))
    body_contents.append(_thuong_table_header())
    body_contents.append(_thuong_table_row(
        "Doanh thu", f"{_fmt_money(bianuoc['base_tien'])} đ",
        f"{_fmt_money(bianuoc['thuc_te_tien'])} đ",
        f"{_fmt_money(bianuoc['du_kien_tien'])} đ",
    ))
    body_contents.append(_thuong_table_row(
        "Tăng trưởng", "—", "—",
        f"{bianuoc['pct']:+.1f}%".replace(".", ","),
    ))
    body_contents.append(_thuong_muc_box(f"Mức thưởng: {bianuoc['muc']}", f"{_fmt_money(bianuoc['thuong_du_kien'])} đ"))
    body_contents.append({"type": "separator", "margin": "xl", "color": DIVIDER})

    # ---- C2 ----
    body_contents.extend(_thuong_section_title("Trà C2 (theo sản phẩm)", "500đ/chai (Freeze/Tắc/Sâm Cúc) · 1.000đ/chai (Olong 1L)"))
    body_contents.append({
        "type": "box", "layout": "horizontal", "backgroundColor": TABLE_HEAD_BG,
        "paddingAll": "6px", "margin": "sm",
        "contents": [
            {"type": "text", "text": "Sản phẩm", "size": "sm", "weight": "bold", "color": BLACK, "flex": 5},
            {"type": "text", "text": "Thực tế", "size": "sm", "weight": "bold", "color": BLACK, "flex": 3, "align": "center"},
            {"type": "text", "text": "Dự kiến 31/8", "size": "sm", "weight": "bold", "color": BLACK, "flex": 4, "align": "center"},
        ],
    })
    for it in c2["items"]:
        tt_text = f"{it['chai_mtd']:,.0f} chai".replace(",", ".")
        dk_text = f"~{it['chai_proj']:,.0f} chai".replace(",", ".")
        body_contents.append({
            "type": "box", "layout": "horizontal", "paddingAll": "6px",
            "contents": [
                {"type": "text", "text": it["ten"], "size": "sm", "color": BLACK, "flex": 5, "wrap": True},
                {"type": "text", "text": tt_text, "size": "sm", "weight": "bold", "color": RED, "flex": 3, "align": "center", "wrap": True},
                {"type": "text", "text": dk_text, "size": "sm", "weight": "bold", "color": GREEN, "flex": 4, "align": "center", "wrap": True},
            ],
        })
    body_contents.append(_thuong_muc_box("Tổng thưởng C2 (dự kiến)", f"{_fmt_money(c2['tong_thuong_du_kien'])} đ"))
    body_contents.append({"type": "separator", "margin": "xl", "color": DIVIDER})

    # ---- TONG ----
    body_contents.append({
        "type": "box", "layout": "horizontal", "backgroundColor": GREEN, "cornerRadius": "10px",
        "paddingAll": "14px", "margin": "lg",
        "contents": [
            {"type": "text", "text": "TỔNG THƯỞNG DỰ KIẾN", "size": "sm", "weight": "bold", "color": "#FFFFFF", "flex": 6, "wrap": True},
            {"type": "text", "text": f"{_fmt_money(payload['tong_thuong_du_kien'])} đ", "size": "lg", "weight": "bold", "color": "#FFFFFF", "flex": 5, "align": "end"},
        ],
    })

    contents = {
        "type": "bubble",
        "size": "giga",
        "header": {
            "type": "box", "layout": "vertical", "backgroundColor": YELLOW, "paddingAll": "16px",
            "contents": [
                {"type": "text", "text": "BÁO CÁO THƯỞNG", "color": BLACK, "weight": "bold", "size": "md"},
                {"type": "text", "text": ten_st or "", "color": BLACK, "size": "lg", "margin": "sm", "wrap": True},
                {"type": "text", "text": f"{ngay_bd} - {ngay_kt} ({so_ngay}/{so_ngay_thang} ngày)", "color": "#3D3200", "size": "md", "margin": "xs"},
            ],
        },
        "body": {
            "type": "box", "layout": "vertical", "paddingAll": "16px", "backgroundColor": PAGE_BG,
            "contents": body_contents,
        },
    }
    return contents

# ---------------------------------------------------------------------------
# BÁO CÁO DOANH THU THỦY HẢI SẢN — dùng lại đúng layout "BÁO CÁO NGÀNH HÀNG"
# ---------------------------------------------------------------------------
def build_seafood_flex_message(ten_st, ngay_bd, ngay_kt, so_ngay, items, doanh_thu_tong, gia_tri_huy, du_kien_cuoi_thang):
    """items: list[(ten_sp, qty_text, pct_ban_nhap)] — dùng lại _category_section_table."""
    body_contents = [
        {
            "type": "box", "layout": "horizontal", "margin": "none",
            "contents": [
                {"type": "text", "text": "Doanh thu:", "size": "lg", "color": BLACK, "flex": 3},
                {"type": "text", "text": f"{_fmt_money(doanh_thu_tong)} đ", "size": "xl", "weight": "bold", "color": RED, "flex": 4, "align": "end"},
            ],
        },
        _category_section_table(
            "THỦY HẢI SẢN CÁC LOẠI",
            None,
            items,
            "Không có dữ liệu trong khoảng ngày này",
        ),
        {"type": "separator", "margin": "xl", "color": DIVIDER},
        _thuong_muc_box("Giá trị hủy/hao hụt (ước theo giá bán TB)", f"{_fmt_money(gia_tri_huy)} đ"),
        _thuong_muc_box("Dự kiến doanh thu cuối tháng", f"{_fmt_money(du_kien_cuoi_thang)} đ"),
    ]
    return {
        "type": "bubble",
        "size": "giga",
        "header": {
            "type": "box", "layout": "vertical", "backgroundColor": YELLOW, "paddingAll": "16px",
            "contents": [
                {"type": "text", "text": "DOANH THU THỦY HẢI SẢN", "color": BLACK, "weight": "bold", "size": "lg"},
                {"type": "text", "text": ten_st or "", "color": BLACK, "size": "sm", "margin": "sm", "wrap": True},
                {"type": "text", "text": f"{ngay_bd} - {ngay_kt} (lũy kế {so_ngay} ngày)", "color": "#3D3200", "size": "xs", "margin": "xs"},
            ],
        },
        "body": {
            "type": "box", "layout": "vertical", "paddingAll": "16px", "backgroundColor": PAGE_BG,
            "contents": body_contents,
        },
    }

# ---------------------------------------------------------------------------
# CÔNG VIỆC 1 — chi tiết HỦY TỒN + MMKK từng sản phẩm (đúng 1 ngày)
# ---------------------------------------------------------------------------
def _fresh_detail_section(title, grouped_items, empty_note):
    """grouped_items: list[(nganh_hang, [(ten_sp, qty_text), ...])], đã sắp theo
    Ngành hàng. Mỗi ngành hàng có tiêu đề nhỏ riêng, dưới là các dòng sản phẩm."""
    contents = [{"type": "text", "text": title, "size": "xl", "weight": "bold", "color": BLACK}]
    if not grouped_items:
        contents.append({"type": "text", "text": empty_note, "size": "md", "color": GRAY, "margin": "sm", "wrap": True})
        return {"type": "box", "layout": "vertical", "margin": "lg", "contents": contents}
    for nganh_hang, sp_list in grouped_items:
        contents.append({
            "type": "text", "text": nganh_hang, "size": "sm", "weight": "bold",
            "color": "#1E5C8C", "margin": "md",
        })
        for i, (ten_sp, qty_text) in enumerate(sp_list):
            row = {
                "type": "box", "layout": "horizontal", "paddingAll": "6px",
                "contents": [
                    {"type": "text", "text": ten_sp, "size": "sm", "color": BLACK, "flex": 6, "wrap": True},
                    {"type": "text", "text": qty_text, "size": "sm", "weight": "bold", "color": RED, "flex": 3, "align": "end", "wrap": True},
                ],
            }
            if i % 2 == 1:
                row["backgroundColor"] = ROW_ALT_BG
            contents.append(row)
    return {"type": "box", "layout": "vertical", "margin": "lg", "contents": contents}

def build_fresh_detail_flex_message(ten_st, ngay_display, huy_groups, mmkk_groups, so_huy, so_mmkk):
    body_contents = [
        _fresh_detail_section(f"HỦY TỒN ({so_huy} sản phẩm)", huy_groups, "Không có hủy tồn ngày này"),
        {"type": "separator", "margin": "xl", "color": DIVIDER},
        _fresh_detail_section(f"MẤT MÁT KIỂM KÊ ({so_mmkk} sản phẩm)", mmkk_groups, "Không có MMKK ngày này"),
    ]
    return {
        "type": "bubble",
        "size": "giga",
        "header": {
            "type": "box", "layout": "vertical", "backgroundColor": YELLOW, "paddingAll": "16px",
            "contents": [
                {"type": "text", "text": "HỦY TỒN + MMKK", "color": BLACK, "weight": "bold", "size": "lg"},
                {"type": "text", "text": ten_st or "", "color": BLACK, "size": "sm", "margin": "sm", "wrap": True},
                {"type": "text", "text": ngay_display, "color": "#3D3200", "size": "xs", "margin": "xs"},
            ],
        },
        "body": {
            "type": "box", "layout": "vertical", "paddingAll": "16px", "backgroundColor": PAGE_BG,
            "contents": body_contents,
        },
    }

# ---------------------------------------------------------------------------
# CÔNG VIỆC 2 — tổng HỦY TỒN + MMKK theo nhóm lớn (khoảng nhiều ngày)
# ---------------------------------------------------------------------------
def build_fresh_group_flex_message(ten_st, ngay_bd, ngay_kt, so_ngay, nhom_rows):
    """nhom_rows: list[(nhom, nhap_text, huy_text, mmkk_text, tong_text, pct_ban_text, pct_mmkk_text)]
    — 4 nhóm lớn + Trứng. (13/09/2026: thêm cột Nhập + 2 dòng tỉ lệ % Bán/Nhập
    và % MMKK/Nhập theo yêu cầu anh Quí — tách thành 2 dòng/nhóm thay vì nhồi
    hết vào 1 hàng ngang cho đỡ chật màn hình điện thoại.)"""
    table_rows = [{
        "type": "box", "layout": "horizontal", "backgroundColor": TABLE_HEAD_BG,
        "paddingAll": "8px", "margin": "sm",
        "contents": [
            {"type": "text", "text": "Nhóm", "size": "sm", "weight": "bold", "color": BLACK, "flex": 3},
            {"type": "text", "text": "Nhập", "size": "sm", "weight": "bold", "color": BLACK, "flex": 2, "align": "center"},
            {"type": "text", "text": "Hủy tồn", "size": "sm", "weight": "bold", "color": BLACK, "flex": 2, "align": "center"},
            {"type": "text", "text": "MMKK", "size": "sm", "weight": "bold", "color": BLACK, "flex": 2, "align": "center"},
            {"type": "text", "text": "Tổng", "size": "sm", "weight": "bold", "color": BLACK, "flex": 2, "align": "end"},
        ],
    }]
    for i, (nhom, nhap_text, huy_text, mmkk_text, tong_text, pct_ban_text, pct_mmkk_text) in enumerate(nhom_rows):
        bg = ROW_ALT_BG if i % 2 == 1 else None
        data_row = {
            "type": "box", "layout": "horizontal", "paddingAll": "8px",
            "contents": [
                {"type": "text", "text": nhom, "size": "sm", "color": BLACK, "flex": 3, "wrap": True},
                {"type": "text", "text": nhap_text, "size": "xs", "weight": "bold", "color": BLACK, "flex": 2, "align": "center", "wrap": True},
                {"type": "text", "text": huy_text, "size": "xs", "weight": "bold", "color": RED, "flex": 2, "align": "center", "wrap": True},
                {"type": "text", "text": mmkk_text, "size": "xs", "weight": "bold", "color": RED, "flex": 2, "align": "center", "wrap": True},
                {"type": "text", "text": tong_text, "size": "xs", "weight": "bold", "color": BLACK, "flex": 2, "align": "end", "wrap": True},
            ],
        }
        ratio_row = {
            "type": "box", "layout": "horizontal", "paddingAll": "8px", "paddingTop": "0px",
            "contents": [
                {"type": "text", "text": f"Bán/Nhập: {pct_ban_text}   MMKK/Nhập: {pct_mmkk_text}",
                 "size": "xxs", "color": GRAY_LIGHT, "flex": 1, "wrap": True},
            ],
        }
        if bg:
            data_row["backgroundColor"] = bg
            ratio_row["backgroundColor"] = bg
        table_rows.append(data_row)
        table_rows.append(ratio_row)
    return {
        "type": "bubble",
        "size": "giga",
        "header": {
            "type": "box", "layout": "vertical", "backgroundColor": YELLOW, "paddingAll": "16px",
            "contents": [
                {"type": "text", "text": "TỔNG HỦY TỒN + MMKK THEO NHÓM", "color": BLACK, "weight": "bold", "size": "md", "wrap": True},
                {"type": "text", "text": ten_st or "", "color": BLACK, "size": "sm", "margin": "sm", "wrap": True},
                {"type": "text", "text": f"{ngay_bd} - {ngay_kt} ({so_ngay} ngày)", "color": "#3D3200", "size": "xs", "margin": "xs"},
            ],
        },
        "body": {
            "type": "box", "layout": "vertical", "paddingAll": "16px", "backgroundColor": PAGE_BG,
            "contents": table_rows,
        },
    }

# ---------------------------------------------------------------------------
# LỆNH "DTDK" - Doanh Thu Dự Kiến (theo tháng + tiến độ target năm)
# ---------------------------------------------------------------------------
def _progress_bar(pct):
    """Thanh tiến độ dạng box ngang (giả lập progress bar bằng flex ratio)."""
    p = max(0, min(100, round(pct)))
    remain = 100 - p
    bar_contents = []
    if p > 0:
        bar_contents.append({
            "type": "box", "layout": "vertical", "flex": p,
            "backgroundColor": GREEN, "cornerRadius": "6px",
            "contents": [{"type": "filler"}],
        })
    if remain > 0:
        bar_contents.append({
            "type": "box", "layout": "vertical", "flex": remain,
            "backgroundColor": DIVIDER, "cornerRadius": "6px",
            "contents": [{"type": "filler"}],
        })
    if not bar_contents:
        bar_contents = [{
            "type": "box", "layout": "vertical", "flex": 1,
            "backgroundColor": DIVIDER, "contents": [{"type": "filler"}],
        }]
    return {
        "type": "box", "layout": "horizontal", "height": "10px",
        "spacing": "none", "margin": "sm", "contents": bar_contents,
    }

def _target_section(label, target_value, remain_value, pct):
    return {
        "type": "box", "layout": "vertical", "backgroundColor": ROW_BG,
        "cornerRadius": "8px", "paddingAll": "10px", "margin": "sm",
        "contents": [
            {"type": "box", "layout": "horizontal", "contents": [
                {"type": "text", "text": label, "size": "xs", "color": GRAY_LIGHT, "flex": 6, "wrap": True},
                {"type": "text", "text": f"{_fmt_money(target_value)} đ", "size": "xs", "weight": "bold", "color": BLACK, "flex": 6, "align": "end", "wrap": True},
            ]},
            {"type": "box", "layout": "horizontal", "margin": "xs", "contents": [
                {"type": "text", "text": "Còn thiếu", "size": "xs", "color": GRAY_LIGHT, "flex": 6},
                {"type": "text", "text": f"{_fmt_money(remain_value)} đ", "size": "xs", "weight": "bold", "color": RED, "flex": 6, "align": "end", "wrap": True},
            ]},
            _progress_bar(pct),
            {"type": "text", "text": f"{pct:.2f}% hoàn thành", "size": "sm", "weight": "bold", "color": GREEN, "align": "end", "margin": "xs"},
        ],
    }

def build_dtdk_flex_message(ten_st, thang_hien_tai_label, tong_hien_tai, tb_ngay,
                             du_kien_thang, month_rows, grand, target_nam, target_tang_them):
    """month_rows: list[(label, tong_thang)] — các tháng ĐÃ QUA (không gồm tháng hiện tại)."""
    table_rows = [{
        "type": "box", "layout": "horizontal", "backgroundColor": TABLE_HEAD_BG,
        "paddingAll": "8px", "margin": "sm",
        "contents": [
            {"type": "text", "text": "Tháng", "size": "sm", "weight": "bold", "color": BLACK, "flex": 5},
            {"type": "text", "text": "Doanh thu", "size": "sm", "weight": "bold", "color": BLACK, "flex": 6, "align": "end"},
        ],
    }]
    for i, (label, val) in enumerate(month_rows):
        row = {
            "type": "box", "layout": "horizontal", "paddingAll": "8px",
            "contents": [
                {"type": "text", "text": label, "size": "sm", "color": BLACK, "flex": 5},
                {"type": "text", "text": f"{_fmt_money(val)} đ", "size": "sm", "weight": "bold", "color": RED, "flex": 6, "align": "end", "wrap": True},
            ],
        }
        if i % 2 == 1:
            row["backgroundColor"] = ROW_ALT_BG
        table_rows.append(row)
    table_rows.append({
        "type": "box", "layout": "horizontal", "backgroundColor": YELLOW,
        "paddingAll": "8px", "margin": "sm", "cornerRadius": "4px",
        "contents": [
            {"type": "text", "text": "TỔNG", "size": "sm", "weight": "bold", "color": BLACK, "flex": 5},
            {"type": "text", "text": f"{_fmt_money(grand)} đ", "size": "sm", "weight": "bold", "color": BLACK, "flex": 6, "align": "end", "wrap": True},
        ],
    })

    if tong_hien_tai > 0:
        thang_hien_tai_box = {
            "type": "box", "layout": "vertical", "backgroundColor": ROW_BG,
            "cornerRadius": "8px", "paddingAll": "10px", "margin": "sm",
            "contents": [
                {"type": "box", "layout": "horizontal", "contents": [
                    {"type": "text", "text": "Đã có", "size": "xs", "color": GRAY_LIGHT, "flex": 5},
                    {"type": "text", "text": f"{_fmt_money(tong_hien_tai)} đ", "size": "xs", "weight": "bold", "color": BLACK, "flex": 7, "align": "end", "wrap": True},
                ]},
                {"type": "box", "layout": "horizontal", "margin": "xs", "contents": [
                    {"type": "text", "text": "TB/ngày", "size": "xs", "color": GRAY_LIGHT, "flex": 5},
                    {"type": "text", "text": f"{_fmt_money(tb_ngay)} đ", "size": "xs", "weight": "bold", "color": BLACK, "flex": 7, "align": "end", "wrap": True},
                ]},
                {"type": "box", "layout": "horizontal", "margin": "xs", "contents": [
                    {"type": "text", "text": "Dự kiến cả tháng", "size": "xs", "color": GRAY_LIGHT, "flex": 5},
                    {"type": "text", "text": f"{_fmt_money(du_kien_thang)} đ", "size": "xs", "weight": "bold", "color": GREEN, "flex": 7, "align": "end", "wrap": True},
                ]},
            ],
        }
    else:
        thang_hien_tai_box = {
            "type": "box", "layout": "vertical", "backgroundColor": ROW_BG,
            "cornerRadius": "8px", "paddingAll": "10px", "margin": "sm",
            "contents": [{"type": "text", "text": "Chưa có dữ liệu tháng này", "size": "xs", "color": GRAY_LIGHT}],
        }

    pct_goc = (grand / target_nam * 100) if target_nam else 0
    pct_tang = (grand / target_tang_them * 100) if target_tang_them else 0
    remain_goc = target_nam - grand
    remain_tang = target_tang_them - grand

    body_contents = (
        [
            {"type": "text", "text": f"🗓️ {thang_hien_tai_label}", "size": "sm", "weight": "bold", "color": BLACK},
            thang_hien_tai_box,
            {"type": "separator", "margin": "lg", "color": DIVIDER},
            {"type": "text", "text": "📅 Tổng theo tháng", "size": "sm", "weight": "bold", "color": BLACK, "margin": "lg"},
        ]
        + table_rows
        + [
            {"type": "text", "text": "🎯 Tiến độ target 2026", "size": "sm", "weight": "bold", "color": BLACK, "margin": "lg"},
            _target_section("Target năm", target_nam, remain_goc, pct_goc),
            _target_section("Mục tiêu tăng thêm (+15%)", target_tang_them, remain_tang, pct_tang),
        ]
    )

    return {
        "type": "bubble",
        "size": "giga",
        "header": {
            "type": "box", "layout": "vertical", "backgroundColor": YELLOW, "paddingAll": "16px",
            "contents": [
                {"type": "text", "text": "DOANH THU THEO THÁNG", "color": BLACK, "weight": "bold", "size": "md", "wrap": True},
                {"type": "text", "text": ten_st or "", "color": BLACK, "size": "sm", "margin": "sm", "wrap": True},
                {"type": "text", "text": "Offline + Online, chưa VAT", "color": "#3D3200", "size": "xs", "margin": "xs"},
            ],
        },
        "body": {
            "type": "box", "layout": "vertical", "paddingAll": "16px", "backgroundColor": PAGE_BG,
            "contents": body_contents,
        },
    }
