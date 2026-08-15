"""
excel_reader.py - Đọc file Excel doanh thu nhiều dòng (mỗi dòng 1 siêu thị)
và trả về danh sách bản ghi theo đúng schema mà storage.py cần.
"""

import openpyxl
from datetime import datetime

COL = {
    "ngay": "Ngày",
    "ma_st": "Mã siêu thị",
    "ten_st": "Tên siêu thị",
    "dt_offline": "Doanh thu offline",
    "dt_online": "Doanh thu Online",
    "bill_offline": "Tổng số bill",
    "bill_online": "Tổng số bill online",
    "tinh": "Tỉnh/TP",
}


def _find_col_index(header_row, name):
    for i, cell in enumerate(header_row, start=1):
        if str(cell.value).strip() == name:
            return i
    return None


def read_all_rows(input_path):
    """Đọc toàn bộ dòng dữ liệu trong sheet đầu tiên.
    Trả về (ngay_str, list_of_dict) — mỗi dict là 1 siêu thị."""
    wb = openpyxl.load_workbook(input_path, data_only=True)
    ws = wb.worksheets[0]

    header_row = list(ws[1])
    col_idx = {key: _find_col_index(header_row, label) for key, label in COL.items()}
    missing = [k for k, v in col_idx.items() if v is None]
    if missing:
        raise ValueError(
            f"Không tìm thấy các cột: {[COL[m] for m in missing]}. "
            "Kiểm tra lại file có đúng định dạng doanh thu theo siêu thị không."
        )

    rows = []
    for r in range(2, ws.max_row + 1):
        ma_st = ws.cell(row=r, column=col_idx["ma_st"]).value
        if ma_st is None:
            continue
        ngay = ws.cell(row=r, column=col_idx["ngay"]).value
        if isinstance(ngay, datetime):
            ngay_str = ngay.strftime("%Y-%m-%d")
        else:
            ngay_str = str(ngay)[:10]

        rows.append({
            "ngay": ngay_str,
            "ma_st": str(ma_st),
            "ten_st": ws.cell(row=r, column=col_idx["ten_st"]).value,
            "tinh": ws.cell(row=r, column=col_idx["tinh"]).value,
            "dt_offline": float(ws.cell(row=r, column=col_idx["dt_offline"]).value or 0),
            "dt_online": float(ws.cell(row=r, column=col_idx["dt_online"]).value or 0),
            "bill_offline": float(ws.cell(row=r, column=col_idx["bill_offline"]).value or 0),
            "bill_online": float(ws.cell(row=r, column=col_idx["bill_online"]).value or 0),
        })

    if not rows:
        raise ValueError("Không tìm thấy dòng dữ liệu nào trong file.")

    return rows[0]["ngay"], rows
