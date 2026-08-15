total = sum((r["dt_offline"] or 0) + (r["dt_online"] or 0) for r in rows)
            reply = (
                f"✅ Đã lưu dữ liệu ngày {date_display} cho {len(rows)} siêu thị.\n"
                f"Tổng doanh thu: {total:,.0f} đ\n\n"
                f"Gõ \"Báo cáo doanh thu\" để xem báo cáo chi tiết."
            ).replace(",", ".")
            reply_text(messaging_api, event.reply_token, reply)

        except Exception as e:
            traceback.print_exc()
            reply_text(messaging_api, event.reply_token, f"Có lỗi khi xử lý file: {e}")


@handler.add(MessageEvent, message=TextMessageContent)
def handle_text_message(event):
    text = event.message.text or ""
    if not REPORT_COMMAND_PATTERN.search(text):
        return

    with ApiClient(configuration) as api_client:
        messaging_api = MessagingApi(api_client)
        try:
            latest_date, latest_records, prev_date, prev_records = storage.get_latest_and_previous()

            if latest_date is None:
                reply_text(
                    messaging_api,
                    event.reply_token,
                    "Chưa có dữ liệu nào được lưu. Anh gửi file Excel doanh thu trước nhé.",
                )
                return

            bubble = build_flex_message(latest_date, latest_records, prev_date, prev_records)
            flex_message = FlexMessage(
                alt_text=f"Báo cáo doanh thu {latest_date}",
                contents=FlexContainer.from_dict(bubble),
            )
            messaging_api.reply_message(
                ReplyMessageRequest(reply_token=event.reply_token, messages=[flex_message])
            )
        except Exception as e:
            traceback.print_exc()
            reply_text(messaging_api, event.reply_token, f"Có lỗi khi tạo báo cáo: {e}")


def reply_text(messaging_api, reply_token, text):
    text = text[:4900]
    messaging_api.reply_message(
        ReplyMessageRequest(reply_token=reply_token, messages=[TextMessage(text=text)])
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
