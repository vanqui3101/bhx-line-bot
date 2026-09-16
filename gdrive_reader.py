"""
gdrive_reader.py - Đọc file Excel tự động từ 1 thư mục Google Drive cố định.
MỚI (16/09/2026): dùng tài khoản dịch vụ (Service Account) của Google để bot
tự vào Drive lấy file mới, KHÔNG cần ai đăng nhập tay hay gửi file vào bot nữa.

Cấu hình cần có (biến môi trường Railway):
- GOOGLE_SERVICE_ACCOUNT_JSON: toàn bộ nội dung file JSON của Service Account,
  dán nguyên văn thành 1 dòng (xem hướng dẫn lấy file này trong lịch sử chat).
- GOOGLE_DRIVE_FOLDER_ID: ID của thư mục Drive chứa file data cho bot đọc
  (thư mục phải được share quyền "Viewer" cho đúng email Service Account).

Thư mục Drive đó phải chỉ chứa file .xlsx/.xlsm mà bot đã biết đọc (doanh thu,
ngành hàng, tồn kho, hủy tồn/MMKK, lịch hỗ trợ, lịch phân ca...) — file nào
không nhận diện được thì bot bỏ qua, không báo lỗi ồn ào.
"""
import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build

_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
_service = None


def _get_service():
    """Tạo (hoặc tái sử dụng) kết nối tới Google Drive API bằng Service Account."""
    global _service
    if _service is not None:
        return _service
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not raw:
        raise RuntimeError("Chưa cấu hình biến môi trường GOOGLE_SERVICE_ACCOUNT_JSON")
    info = json.loads(raw)
    creds = service_account.Credentials.from_service_account_info(info, scopes=_SCOPES)
    _service = build("drive", "v3", credentials=creds, cache_discovery=False)
    return _service


def list_files_in_folder(folder_id):
    """Trả về danh sách file .xlsx/.xlsm trong thư mục (không lấy thư mục con),
    sắp xếp CŨ -> MỚI theo modifiedTime. Mỗi phần tử: {id, name, modifiedTime}."""
    service = _get_service()
    query = f"'{folder_id}' in parents and trashed = false"
    result = []
    page_token = None
    while True:
        resp = service.files().list(
            q=query,
            fields="nextPageToken, files(id, name, modifiedTime, mimeType)",
            pageToken=page_token,
            pageSize=100,
        ).execute()
        for f in resp.get("files", []):
            name = f.get("name", "")
            if name.lower().endswith((".xlsx", ".xlsm")):
                result.append({
                    "id": f["id"],
                    "name": name,
                    "modifiedTime": f.get("modifiedTime", ""),
                })
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    result.sort(key=lambda x: x["modifiedTime"])
    return result


def download_file(file_id, dest_path):
    """Tải nội dung file (dạng nhị phân thông thường, không phải Google Docs/Sheets
    online) về đường dẫn dest_path trên máy chủ bot."""
    service = _get_service()
    content = service.files().get_media(fileId=file_id).execute()
    with open(dest_path, "wb") as f:
        f.write(content)
