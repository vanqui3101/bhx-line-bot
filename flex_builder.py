"""
flex_builder.py - Dựng nội dung LINE Flex Message cho báo cáo doanh thu.
"""

from datetime import datetime

MAX_ROWS = 10  # số siêu thị hiển thị trong thẻ (LINE giới hạn kích thước message)


def _fmt_money(n):
    return f"{n:,.0f}".replace(",", ".")


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
        return "—", "#999999"
    arrow = "▲" if pct >= 0 else "▼"
    color = "#2E7D32" if pct >= 0 else "#C62828"
    return f"{arrow}{abs(pct):.1f}%", color


def _metric_row(label, value_text, pct):
    pct_text, pct_color = _pct_text(pct)
    return {
        "type": "box",
        "layout": "horizontal",
        "margin": "md",
        "contents": [
            {"type": "text", "text": label, "size": "xs", "color": "#888888", "flex": 4},
            {"type": "text", "text": value_text, "size": "sm", "weight": "bold", "color": "#1F4E78", "flex": 4, "align": "end"},
            {"type": "text", "text": pct_text, "size": "xs", "weight": "bold", "color": pct_color, "flex": 2, "align": "end"},
        ],
    }


def build_flex_message(latest_date, latest_records, prev_date, prev_records):
    total_now = sum(_total_dt(r) for r in latest_records)
    online_now = sum(_online_dt(r) for r in latest_records)
    avg_bill_now = _avg_bill_value(latest_records)

    if prev_records:
        total_prev = sum(_total_dt(r) for r in prev_records)
        online_prev = sum(_online_dt(r) for r in prev_records)
        avg_bill_prev = _avg_bill_value(prev_records)
        pct_total = _pct_change(total_now, total_prev)
        pct_online = _pct_change(online_now, online_prev)
        pct_avg_bill = _pct_change(avg_bill_now, avg_bill_prev)
    else:
        pct_total = pct_online = pct_avg_bill = None

    pct_text, pct_color = _pct_text(pct_total)

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
            "contents": [
                {
                    "type": "text",
                    "text": rec.get("ten_st") or rec.get("ma_st"),
                    "size": "sm",
                    "flex": 5,
                    "wrap": True,
                    "color": "#333333"
                },
                {
                    "type": "text",
                    "text": _fmt_money(dt_now),
                    "size": "sm",
                    "flex": 3,
                    "align": "end",
                    "color": "#333333"
                },
                {
                    "type": "text",
                    "text": row_pct_text,
                    "size": "xs",
                    "flex": 2,
                    "align": "end",
                    "color": row_pct_color
                }
            ],
            "margin": "sm"
        })

    subtitle = f"{len(latest_records)} siêu thị"
    if prev_date:
        subtitle += f" · So với {_fmt_date_display(prev_date)} (cùng ngày tháng trước)"

    note_text = "Danh sách trên là top siêu thị theo doanh thu. Xem đầy đủ trong file Excel đính kèm."
    if not prev_records:
        note_text = "Chưa có dữ liệu cùng ngày tháng trước để so sánh. " + note_text

    contents = {
        "type": "bubble",
        "size": "giga",
        "header": {
            "type": "box",
            "layout": "vertical",
            "backgroundColor": "#1F4E78",
            "paddingAll": "16px",
            "contents": [
                {
                    "type": "text",
                    "text": "📊 BÁO CÁO DOANH THU",
                    "color": "#FFFFFF",
                    "weight": "bold",
                    "size": "lg"
                },
                {
                    "type": "text",
                    "text": f"{_fmt_date_display(latest_date)} · {subtitle}",
                    "color": "#E0E7EF",
                    "size": "xs",
                    "margin": "sm"
                }
            ]
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
                        {
                            "type": "text",
                            "text": "TỔNG DOANH THU",
                            "size": "xs",
                            "color": "#888888"
                        },
                        {
                            "type": "box",
                            "layout": "horizontal",
                            "margin": "sm",
                            "contents": [
                                {
                                    "type": "text",
                                    "text": f"{_fmt_money(total_now)} đ",
                                    "size": "xl",
                                    "weight": "bold",
                                    "color": "#1F4E78",
                                    "flex": 3
                                },
                                {
                                    "type": "text",
                                    "text": pct_text,
                                    "size": "md",
                                    "weight": "bold",
                                    "color": pct_color,
                                    "align": "end",
                                    "flex": 1,
                                    "gravity": "bottom"
                                }
                            ]
                        },
                        {"type": "separator", "margin": "md"},
                        _metric_row("Doanh thu online", f"{_fmt_money(online_now)} đ", pct_online),
                        _metric_row("Giá trị bill TB", f"{_fmt_money(avg_bill_now)} đ", pct_avg_bill),
                    ]
                },
                {"type": "separator", "margin": "lg"},
                {
                    "type": "box",
                    "layout": "horizontal",
                    "margin": "lg",
                    "contents": [
                        {"type": "text", "text": "Siêu thị", "size": "xs", "color": "#999999", "flex": 5},
                        {"type": "text", "text": "Doanh thu", "size": "xs", "color": "#999999", "flex": 3, "align": "end"},
                        {"type": "text", "text": "So T trước", "size": "xs", "color": "#999999", "flex": 2, "align": "end"}
                    ]
                },
                *store_rows,
                {"type": "separator", "margin": "lg"},
                {
                    "type": "text",
                    "text": note_text,
                    "size": "xxs",
                    "color": "#AAAAAA",
                    "margin": "md",
                    "wrap": True
                }
            ]
        }
    }
    return contents
