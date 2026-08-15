"""
flex_builder.py - Dựng nội dung LINE Flex Message.

Theme: nền trắng, header vàng, chữ nhãn màu đen, số liệu màu đỏ.

Có 2 loại thẻ:
- build_flex_message(...)          -> Báo cáo DOANH THU (Offline / Online / Tổng / Bill TB)
- build_category_flex_message(...) -> Báo cáo NGÀNH HÀNG (Nấm / Bánh trung thu / Trà C2)
"""

from datetime import datetime

YELLOW = "#FFEE00"
BLACK = "#1A1A1A"
RED = "#D62020"
GRAY = "#666666"
GRAY_LIGHT = "#555555"
DIVIDER = "#A8CCE0"
ROW_BG = "#D7ECFA"
PAGE_BG = "#C8E6F9"


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
    return f"{sign}{_fmt_money(abs(delta))} đ ({arrow}{abs(pct):.1f}%)"


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


def build_flex_message(latest_date, latest_records, prev_date, prev_records, gio_now=None, gio_prev=None):
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

def _category_item_row(ten, qty_text, thanh_tien):
    return {
        "type": "box",
        "layout": "vertical",
        "margin": "md",
        "contents": [
            {"type": "text", "text": f"• {ten}", "size": "xs", "color": BLACK, "wrap": True},
            {
                "type": "text",
                "text": f"{qty_text}  —  {_fmt_money(thanh_tien)} đ",
                "size": "xs",
                "weight": "bold",
                "color": RED,
                "align": "end",
            },
        ],
    }


def _category_section(title, note, items, total_label, total_value_text):
    contents = [{"type": "text", "text": title, "size": "md", "weight": "bold", "color": BLACK}]
    if note:
        contents.append({"type": "text", "text": note, "size": "xxs", "color": GRAY, "margin": "xs", "wrap": True})
    for it in items:
        contents.append(_category_item_row(*it))
    if not items and not note:
        pass
    contents.append({
        "type": "text",
        "text": f"{total_label}:  {total_value_text}",
        "size": "sm",
        "weight": "bold",
        "color": RED,
        "align": "end",
        "margin": "md",
    })
    return {
        "type": "box",
        "layout": "vertical",
        "margin": "lg",
        "contents": contents,
    }


def build_category_flex_message(ngay, ten_st, payload):
    nam_dt = payload["nam"]["doanh_thu"]

    btt = payload["banh_trung_thu"]
    btt_items = [(it["ten"], f"{_fmt_int(it['sl'])} cái", it["thanh_tien"]) for it in btt["items"]]

    c2 = payload["c2"]
    c2_items = [(it["ten"], f"{_fmt_int(it['chai'])} chai", it["thanh_tien"]) for it in c2["items"]]

    sections = [
        _category_section("NẤM", None, [], "Doanh thu", f"{_fmt_money(nam_dt)} đ"),
        _category_section(
            "BÁNH TRUNG THU",
            "(chỉ tính hàng bán, không tính tặng)",
            btt_items,
            f"Tổng {_fmt_int(btt['tong_sl'])} cái",
            f"{_fmt_money(btt['tong_tien'])} đ",
        ),
        _category_section(
            "TRÀ C2",
            "(quy đổi 1 thùng = 24 chai, 1 lốc = 6 chai)",
            c2_items,
            f"Tổng {_fmt_int(c2['tong_chai'])} chai",
            f"{_fmt_money(c2['tong_tien'])} đ",
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
