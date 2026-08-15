"""
flex_builder.py - Dựng nội dung LINE Flex Message cho báo cáo doanh thu.
Bố cục 2 cột: Tháng này | Tháng trước, kèm chênh lệch tiền + %.
"""

from datetime import datetime

MAX_ROWS = 10  # số siêu thị hiển thị trong bảng chi tiết

NAVY = "#1F4E78"
GRAY = "#999999"
GRAY_LIGHT = "#888888"
DARK = "#333333"
GREEN = "#2E7D32"
RED = "#C62828"


def _fmt_money(n):
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


def _pct_text(pct):
    if pct is None:
        return "—", GRAY
    arrow = "▲" if pct >= 0 else "▼"
    color = GREEN if pct >= 0 else RED
    return f"{arrow}{abs(pct):.1f}%", color


def _delta_text(now, prev, pct):
    if prev is None:
        return "Chưa có dữ liệu tháng trước để so sánh", GRAY
    delta = now - prev
    sign = "+" if delta >= 0 else "-"
    color = GREEN if delta >= 0 else RED
    pct_text, _ = _pct_text(pct)
    return f"Chênh lệch: {sign}{_fmt_money(abs(delta))} đ ({pct_text})", color


def _metric_compare_block(label, now_val, prev_val, is_last=False):
    pct = _pct_change(now_val, prev_val) if prev_val is not None else None
    delta_str, delta_color = _delta_text(now_val, prev_val, pct)

    block = {
        "type": "box",
        "layout": "vertical",
        "margin": "lg" if not is_last else "lg",
        "contents": [
            {"type": "text", "text": label, "size": "xs", "color": GRAY_LIGHT},
            {
                "type": "box",
                "layout": "horizontal",
                "margin": "xs",
                "contents": [
                    {
                        "type": "text",
                        "text": f"{_fmt_money(now_val)} đ",
                        "size": "md",
                        "weight": "bold",
                        "color": NAVY,
                        "flex": 5,
                    },
                    {
                        "type": "text",
                        "text": f"{_fmt_money(prev_val)} đ" if prev_val is not None else "—",
                        "size": "sm",
                        "color": GRAY,
                        "align": "end",
                        "flex": 5,
                    },
                ],
            },
            {
                "type": "text",
                "text": delta_str,
                "size": "xxs",
                "color": delta_color,
                "align": "end",
                "margin": "xs",
                "wrap": True,
            },
        ],
    }
    return block


def build_flex_message(latest_date, latest_records, prev_date, prev_records, gio_now=None, gio_prev=None):
    total_now = sum(_total_dt(r) for r in latest_records)
    online_now = sum(_online_dt(r) for r in latest_records)
    avg_bill_now = _avg_bill_value(latest_records)

    if prev_records:
        total_prev = sum(_total_dt(r) for r in prev_records)
        online_prev = sum(_online_dt(r) for r in prev_records)
        avg_bill_prev = _avg_bill_value(prev_records)
    else:
        total_prev = online_prev = avg_bill_prev = None

    prev_by_store = {r["ma_st"]: _total_dt(r) for r in prev_records} if prev_records else {}

    rows_sorted = sorted(latest_records, key=lambda r: -_total_dt(r))[:MAX_ROWS]

    store_rows = []
    for rec in rows_sorted:
        dt_now = _total_dt(rec)
        dt_prev = prev_by_store.get(rec["ma_st"])
        row_pct = _pct_change(dt_now, dt_prev) if dt_prev else None
        row_pct_text, row_pct_color = _pct_text(row_pct)
        store_rows.append({
            "type": "box",
            "layout": "horizontal",
            "margin": "sm",
            "contents": [
                {"type": "text", "text": (rec.get("ten_st") or rec.get("ma_st"))[:18], "size": "xs", "flex": 5, "wrap": True, "color": DARK},
                {"type": "text", "text": _fmt_money(dt_now), "size": "xs", "flex": 4, "align": "end", "color": DARK},
                {"type": "text", "text": _fmt_money(dt_prev) if dt_prev else "—", "size": "xs", "flex": 4, "align": "end", "color": GRAY},
                {"type": "text", "text": row_pct_text, "size": "xxs", "flex": 3, "align": "end", "color": row_pct_color},
            ],
        })

    # Nhãn cột thời gian: ưu tiên hiện kèm giờ cập nhật nếu có
    label_now = _fmt_date_short(latest_date)
    if gio_now:
        label_now += f" ({gio_now})"
    label_prev = _fmt_date_short(prev_date) if prev_date else "—"
    if gio_prev:
        label_prev += f" ({gio_prev})"

    subtitle = f"{len(latest_records)} siêu thị"
    if prev_date:
        subtitle += f" · So với {_fmt_date_display(prev_date)} (cùng ngày tháng trước)"

    note_text = "Top siêu thị theo doanh thu. Xem đầy đủ trong file Excel đính kèm."
    if not prev_records:
        note_text = "Chưa có dữ liệu cùng ngày tháng trước để so sánh. " + note_text

    column_header = {
        "type": "box",
        "layout": "horizontal",
        "contents": [
            {"type": "text", "text": "Siêu thị", "size": "xxs", "color": GRAY, "flex": 5},
            {"type": "text", "text": f"Tháng này\n{label_now}", "size": "xxs", "color": NAVY, "weight": "bold", "align": "end", "flex": 5, "wrap": True},
            {"type": "text", "text": f"Tháng trước\n{label_prev}", "size": "xxs", "color": GRAY, "align": "end", "flex": 5, "wrap": True},
        ],
    }

    contents = {
        "type": "bubble",
        "size": "giga",
        "header": {
            "type": "box",
            "layout": "vertical",
            "backgroundColor": NAVY,
            "paddingAll": "16px",
            "contents": [
                {"type": "text", "text": "📊 BÁO CÁO DOANH THU", "color": "#FFFFFF", "weight": "bold", "size": "lg"},
                {"type": "text", "text": f"{_fmt_date_display(latest_date)} · {subtitle}", "color": "#E0E7EF", "size": "xs", "margin": "sm", "wrap": True},
            ],
        },
        "body": {
            "type": "box",
            "layout": "vertical",
            "paddingAll": "16px",
            "contents": [
                {
                    "type": "box",
                    "layout": "vertical",
                    "backgroundColor": "#F5F7FA",
                    "cornerRadius": "8px",
                    "paddingAll": "12px",
                    "contents": [
                        column_header,
                        {"type": "separator", "margin": "md"},
                        _metric_compare_block("TỔNG DOANH THU", total_now, total_prev),
                        {"type": "separator", "margin": "lg"},
                        _metric_compare_block("DOANH THU ONLINE", online_now, online_prev),
                        {"type": "separator", "margin": "lg"},
                        _metric_compare_block("GIÁ TRỊ BILL TRUNG BÌNH", avg_bill_now, avg_bill_prev, is_last=True),
                    ],
                },
                {"type": "separator", "margin": "lg"},
                {
                    "type": "box",
                    "layout": "horizontal",
                    "margin": "lg",
                    "contents": [
                        {"type": "text", "text": "Siêu thị", "size": "xs", "color": GRAY, "flex": 5},
                        {"type": "text", "text": "T này", "size": "xs", "color": GRAY, "flex": 4, "align": "end"},
                        {"type": "text", "text": "T trước", "size": "xs", "color": GRAY, "flex": 4, "align": "end"},
                        {"type": "text", "text": "%", "size": "xs", "color": GRAY, "flex": 3, "align": "end"},
                    ],
                },
                *store_rows,
                {"type": "separator", "margin": "lg"},
                {"type": "text", "text": note_text, "size": "xxs", "color": GRAY, "margin": "md", "wrap": True},
            ],
        },
    }
    return contents
