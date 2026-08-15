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
