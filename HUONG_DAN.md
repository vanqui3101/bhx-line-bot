# Hướng dẫn tạo LINE Bot báo cáo doanh thu (từ A-Z)

Anh gửi file Excel doanh thu vào LINE, bot sẽ tự đọc và trả về:
- Tin nhắn tóm tắt (Doanh thu / Lượt bill / Giá trị bill — Offline / Online / Tổng)
- File Excel báo cáo để tải về

Không cần biết code — chỉ cần làm theo từng bước dưới đây.

---

## BƯỚC 1 — Tạo tài khoản LINE Official Account + Messaging API

1. Vào https://developers.line.biz/console/ và đăng nhập bằng tài khoản LINE của anh.
2. Chọn **Create a new provider** → đặt tên bất kỳ (vd: "BHX Bot") → Create.
3. Trong provider vừa tạo, chọn **Create a new channel** → chọn **Messaging API**.
4. Điền thông tin: tên kênh (vd: "Bot doanh thu"), mô tả, chuyên mục ngành hàng → Create.
5. Vào tab **Messaging API** của channel vừa tạo:
   - Kéo xuống mục **Channel access token** → bấm **Issue** → copy và lưu lại (đây là `LINE_CHANNEL_ACCESS_TOKEN`).
   - Ở tab **Basic settings**, tìm mục **Channel secret** → copy và lưu lại (đây là `LINE_CHANNEL_SECRET`).
6. Vẫn ở tab **Messaging API**:
   - Tắt (turn off) mục **Auto-reply messages** và **Greeting messages** (để tránh xung đột với bot).
   - Bật (turn on) mục **Use webhook**.
   - Quét mã QR trong trang này bằng LINE trên điện thoại để kết bạn với bot — dùng để test sau này.

Tạm thời chưa cần điền **Webhook URL**, sẽ quay lại sau khi deploy xong (Bước 3).

---

## BƯỚC 2 — Chuẩn bị mã nguồn trên GitHub

1. Tạo tài khoản tại https://github.com nếu chưa có.
2. Bấm **New repository** → đặt tên (vd: `bhx-line-bot`) → chọn **Public** hoặc **Private** đều được → Create repository.
3. Bấm **uploading an existing file** (hoặc **Add file → Upload files**).
4. Kéo thả toàn bộ các file trong gói mã nguồn em gửi bên dưới vào (`app.py`, `report_generator.py`, `requirements.txt`, `Procfile`).
5. Bấm **Commit changes** để lưu.

---

## BƯỚC 3 — Deploy server lên Railway (miễn phí, dễ dùng nhất)

1. Vào https://railway.app → đăng nhập bằng tài khoản GitHub.
2. Bấm **New Project** → **Deploy from GitHub repo** → chọn repo `bhx-line-bot` vừa tạo.
3. Railway sẽ tự nhận diện và build project. Đợi vài phút cho build xong.
4. Vào tab **Variables** của project, thêm 3 biến môi trường:
   | Tên biến | Giá trị |
   |---|---|
   | `LINE_CHANNEL_ACCESS_TOKEN` | token đã copy ở Bước 1 |
   | `LINE_CHANNEL_SECRET` | secret đã copy ở Bước 1 |
   | `PUBLIC_BASE_URL` | để trống trước, điền sau khi có domain ở bước 5 |
5. Vào tab **Settings** → mục **Networking** → bấm **Generate Domain**. Railway sẽ cho anh một đường link dạng:
   `https://ten-app-cua-anh.up.railway.app`
6. Copy link đó, quay lại tab **Variables**, cập nhật `PUBLIC_BASE_URL` = link vừa copy (không có dấu `/` ở cuối).
7. Project sẽ tự khởi động lại (redeploy) sau khi đổi biến môi trường.

---

## BƯỚC 4 — Kết nối Webhook URL với LINE

1. Quay lại https://developers.line.biz/console/ → vào channel đã tạo → tab **Messaging API**.
2. Ở mục **Webhook URL**, bấm **Edit**, dán vào:
   `https://ten-app-cua-anh.up.railway.app/callback`
3. Bấm **Update** → bấm nút **Verify** để kiểm tra kết nối. Nếu hiện "Success" là đã xong.

---

## BƯỚC 5 — Test thử

1. Mở LINE trên điện thoại, vào cuộc trò chuyện với bot (đã kết bạn ở Bước 1).
2. Gửi file Excel doanh thu (đúng định dạng cột như file mẫu anh đã dùng) vào khung chat.
3. Chờ vài giây, bot sẽ trả lời tin nhắn tóm tắt kèm link tải file báo cáo Excel.

---

## Lưu ý quan trọng

- File gửi vào bot phải là **.xlsx** và có đúng các cột: *Ngày, Mã siêu thị, Tên siêu thị, Doanh thu offline, Doanh thu Online, Tổng số bill, Tổng số bill online, Tỉnh/TP* (giống file mẫu ban đầu).
- Nếu sau này BHX đổi tên cột hoặc cấu trúc file, cần sửa lại phần `COL` trong file `report_generator.py` cho khớp.
- Gói Railway miễn phí có giới hạn giờ chạy/tháng — nếu dùng nhiều, có thể cần nâng cấp gói trả phí nhỏ (vài đô/tháng).
- Nếu cần hỗ trợ thêm (nhiều siêu thị/dòng trong 1 file, gửi báo cáo tự động theo giờ cố định...), cứ nhắn em, em điều chỉnh code giúp anh.
