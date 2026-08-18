"""
flex_builder.py - Dựng nội dung LINE Flex Message.

Theme: nền trắng, header vàng, chữ nhãn màu đen, số liệu màu đỏ.

Có 2 loại thẻ:
- build_flex_message(...)          -> Báo cáo DOANH THU (Offline / Online / Tổng / Bill TB)
- build_category_flex_message(...) -> Báo cáo NGÀNH HÀNG (Nấm / Bánh trung thu / Trà C2)
"""

from datetime import datetime
import math

YELLOW = "#FFEE00"
BLACK = "#1A1A1A"
RED = "#D62020"
GRAY = "#666666"
GRAY_LIGHT = "#555555"
DIVIDER = "#A8CCE0"
ROW_BG = "#D7ECFA"
PAGE_BG = "#C8E6F9"
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
# BÁO CÁO NGÀNH HÀNG (Nấm / Bánh trung thu / Trà C2)
# ---------------------------------------------------------------------------

TABLE_HEAD_BG = "#AAD2F0"
ROW_ALT_BG = "#DEF0FB"


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
    """Mục có bảng sản phẩm (Bánh trung thu / Trà C2) — không có dòng Tổng."""
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


def _category_section_simple(title, total_label, total_value_text):
    """Mục chỉ có 1 dòng tổng (Nấm)."""
    return {
        "type": "box",
        "layout": "vertical",
        "margin": "lg",
        "contents": [
            {"type": "text", "text": title, "size": "xl", "weight": "bold", "color": BLACK},
            {
                "type": "text",
                "text": f"{total_label}:  {total_value_text}",
                "size": "lg",
                "weight": "bold",
                "color": RED,
                "align": "end",
                "margin": "md",
            },
        ],
    }


def build_category_flex_message(ngay, ten_st, payload):
    nam_dt = payload["nam"]["doanh_thu"]

    btt = payload["banh_trung_thu"]
    btt_items = [
        (it["ten"], f"{_fmt_int(it['sl'])} cái", it.get("pct_ban_nhap"))
        for it in btt["items"]
    ]

    c2 = payload["c2"]
    c2_items = [
        (it["ten"], f"{math.ceil(it['chai'] / 24)} thùng", it.get("pct_ban_nhap"))
        for it in c2["items"]
    ]

    sections = [
        _category_section_simple("NẤM", "Doanh thu", f"{_fmt_money(nam_dt)} đ"),
        _category_section_table(
            "BÁNH TRUNG THU",
            "(chỉ tính hàng bán, không tính tặng)",
            btt_items,
            "Không có dữ liệu bán trong ngày",
        ),
        _category_section_table(
            "TRÀ C2",
            "(quy đổi 1 thùng = 24 chai, 1 lốc = 6 chai)",
            c2_items,
            "Không có dữ liệu bán trong ngày",
        ),
    ]

    body_contents = []
    for i, sec in enumerate(sections):
        body_contents.append(sec)
        if i < len(sections) - 1:
            body_contents.append({"type": "separator", "margin": "lg", "color": DIVIDER})

    contents = {
        "type": "bubble",
        "size": "giga",
        "header": {
            "type": "box",
            "layout": "vertical",
            "backgroundColor": YELLOW,
            "paddingAll": "16px",
            "contents": [
                {"type": "text", "text": "BÁO CÁO NGÀNH HÀNG", "color": BLACK, "weight": "bold", "size": "lg"},
                {"type": "text", "text": ten_st or "", "color": BLACK, "size": "sm", "margin": "sm", "wrap": True},
                {"type": "text", "text": _fmt_date_display(ngay), "color": "#3D3200", "size": "xs", "margin": "xs"},
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
