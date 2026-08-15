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
