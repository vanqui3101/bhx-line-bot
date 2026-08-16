"""
excel_reader.py - Đọc 2 loại file Excel mà bot nhận:

1) File "doanh thu theo siêu thị" (nhiều dòng, mỗi dòng 1 siêu thị / 1 ngày)
   -> read_all_rows()  (giữ nguyên logic cũ, không đổi)

2) File "doanh thu chi tiết" theo ngành hàng / sản phẩm (nhiều dòng, mỗi dòng
   1 sản phẩm bán ra trong ngày, cùng 1 siêu thị)
   -> read_category_rows()  (mới thêm, phục vụ báo cáo "MỤC TIÊU KHUYẾN MÃI":
      Nấm / Bánh trung thu / Trà C2)

detect_file_type() dùng để tự động phân biệt 2 loại file khi người dùng gửi
file vào bot, không cần người dùng khai báo loại file.
"""
import openpyxl
from datetime import datetime

# ---------------------------------------------------------------------------
# LOẠI 1: FILE DOANH THU THEO SIÊU THỊ (giữ nguyên như code cũ)
# ---------------------------------------------------------------------------
COL = {
    "ngay": "Ngày",
    "ma_st": "Mã siêu thị",
    "ten_st": "Tên siêu thị",
    # Dùng số liệu CHƯA VAT theo lựa chọn của anh
    "dt_offline": "Doanh thu offline",
    "dt_online": "Doanh thu Online",
    "bill_offline": "Tổng số bill",
    "bill_online": "Tổng số bill online",
    "tinh": "Tỉnh/TP",
}

# Cột đặc trưng của file "chi tiết ngành hàng" (per-product)
CATEGORY_COL = {
    "ngay": "Ngày xuất",
    "ten_st": "Tên siêu thị",
    "ten_sp": "Tên sản phẩm",
    "nganh_hang": "Ngành hàng",
    "nhom_hang": "Nhóm hàng",
    "ten_hang": "Tên Hãng",
    "don_vi": "Đơn vị",
    "sl_thuc_xuat": "SL thực xuất",
    "thanh_tien": "Thành tiền phải thu khách hàng (chưa VAT)",
    "hinh_thuc_xuat": "Tên hình thức xuất bán",
}


def _header_row(ws):
    return [str(c.value).strip() if c.value is not None else "" for c in ws[1]]


def detect_file_type(input_path):
    """Trả về 'revenue', 'category', hoặc None nếu không nhận diện được."""
    wb = openpyxl.load_workbook(input_path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    headers = set(_header_row(ws))
    wb.close()

    if {"Doanh thu offline", "Doanh thu Online", "Mã siêu thị"}.issubset(headers):
        return "revenue"
    if {"Tên sản phẩm", "Ngành hàng", "Nhóm hàng"}.issubset(headers):
        return "category"
    return None


def _find_col_index(header_row, name):
    for i, cell in enumerate(header_row, start=1):
        if str(cell.value).strip() == name:
            return i
    return None


def read_all_rows(input_path):
    """Đọc toàn bộ dòng dữ liệu trong sheet đầu tiên (file doanh thu theo siêu thị).
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


# ---------------------------------------------------------------------------
# LOẠI 2: FILE CHI TIẾT NGÀNH HÀNG (Nấm / Bánh trung thu / Trà C2)
# ---------------------------------------------------------------------------

# Các sản phẩm C2 cần theo dõi, quy đổi ra "chai".
# key = tên hiển thị, value = từ khoá để nhận diện trong "Tên sản phẩm"
C2_TARGETS = [
    ("Trà xanh hương chanh 360ml", "HƯƠNG CHANH"),
    ("Trà hồng vải", "HỒNG VẢI"),
    ("Trà đen dâu anh đào", "DÂU ANH ĐÀO"),
    ("Trà xanh chanh bạc hà", "BẠC HÀ"),
    ("Nước sâm C2 Cool", "SÂM"),
    ("Trà đen tắc", "TẮC"),
]

UNIT_TO_CHAI = {
    "THÙNG": 24,
    "LỐC": 6,
    "CHAI": 1,
}


def _unit_multiplier(don_vi):
    if not don_vi:
        return 1
    key = str(don_vi).strip().upper()
    return UNIT_TO_CHAI.get(key, 1)


# 5 nhóm ngành hàng gộp theo yêu cầu, hiển thị trong bảng "DOANH THU THEO NGÀNH HÀNG"
# của thẻ báo cáo doanh thu. key = tên hiển thị, value = danh sách tên "Ngành hàng"
# gốc trong file cần cộng gộp lại.
NGANH_HANG_GROUPS = [
    ("Bia Nước", ["Bia Các Loại", "Thức uống giải khát các loại"]),
    ("Sữa - Đông Mát", ["Sữa - Thức uống bổ dưỡng các loại", "Thực phẩm đông lạnh - Hàng mát các loại"]),
    ("Thịt Gia Cầm", ["Thịt gia cầm gia súc các loại"]),
    ("Trái Cây - Rau Củ", ["Trái Cây Các Loại", "Rau Củ Các Loại"]),
    ("Thủy Hải Sản", ["Thủy Hải Sản Các Loại"]),
]


def read_category_rows(input_path):
    """Đọc file chi tiết ngành hàng, trả về dict tổng hợp:
    {
        "ngay": "2026-08-15",
        "ten_st": "BHX_STR_CLD - Thửa 1289 An Nghiệp",
        "nam": {"doanh_thu": 798095},
        "banh_trung_thu": {
            "items": [{"ten": "...", "sl": 1, "thanh_tien": 78704}, ...],
            "tong_sl": 5,
            "tong_tien": 307407,
        },
        "c2": {
            "items": [{"ten": "...", "chai": 494, "thanh_tien": 2872727}, ...],
            "tong_chai": 694,
            "tong_tien": 4054764,
        },
    }
    """
    wb = openpyxl.load_workbook(input_path, data_only=True)
    ws = wb.worksheets[0]
    header_row = list(ws[1])
    col_idx = {key: _find_col_index(header_row, label) for key, label in CATEGORY_COL.items()}
    missing = [k for k, v in col_idx.items() if v is None]
    if missing:
        raise ValueError(
            f"Không tìm thấy các cột: {[CATEGORY_COL[m] for m in missing]}. "
            "Kiểm tra lại file có đúng định dạng chi tiết ngành hàng không."
        )

    def cell(r, key):
        return ws.cell(row=r, column=col_idx[key]).value

    ngay_str = None
    ten_st = None

    nam_total = 0.0

    # gộp theo tên sản phẩm để không lặp dòng trùng
    btt_map = {}   # ten_sp -> {sl, thanh_tien}
    c2_map = {label: {"chai": 0.0, "thanh_tien": 0.0} for label, _ in C2_TARGETS}
    nganh_hang_map = {}  # ten_nganh_hang -> thanh_tien (tổng doanh thu toàn bộ ngành hàng)

    for r in range(2, ws.max_row + 1):
        ten_sp = cell(r, "ten_sp")
        if ten_sp is None:
            continue
        ten_sp_str = str(ten_sp).strip()
        ten_sp_upper = ten_sp_str.upper()

        if ngay_str is None:
            ngay_val = cell(r, "ngay")
            if isinstance(ngay_val, datetime):
                ngay_str = ngay_val.strftime("%Y-%m-%d")
            elif ngay_val:
                ngay_str = str(ngay_val)[:10]
        if ten_st is None:
            v = cell(r, "ten_st")
            if v:
                ten_st = str(v).strip()

        sl_thuc_xuat = float(cell(r, "sl_thuc_xuat") or 0)
        thanh_tien = float(cell(r, "thanh_tien") or 0)
        nhom_hang = str(cell(r, "nhom_hang") or "").strip()
        ten_hang = str(cell(r, "ten_hang") or "").strip().upper()
        don_vi = cell(r, "don_vi")
        hinh_thuc = str(cell(r, "hinh_thuc_xuat") or "").strip().upper()

        # ---- NẤM ----
        if nhom_hang == "Nấm Các Loại":
            nam_total += thanh_tien

        # ---- TỔNG DOANH THU THEO NGÀNH HÀNG (tất cả ngành hàng) ----
        nganh_hang = str(cell(r, "nganh_hang") or "").strip()
        if nganh_hang:
            nganh_hang_map[nganh_hang] = nganh_hang_map.get(nganh_hang, 0.0) + thanh_tien

        # ---- BÁNH TRUNG THU ----
        if "TRUNG THU" in ten_sp_upper:
            # bỏ qua hàng tặng / khuyến mãi, chỉ tính hàng bán thực tế
            if "TẶNG" not in hinh_thuc and sl_thuc_xuat > 0:
                entry = btt_map.setdefault(ten_sp_str, {"sl": 0.0, "thanh_tien": 0.0})
                entry["sl"] += sl_thuc_xuat
                entry["thanh_tien"] += thanh_tien

        # ---- TRÀ C2 ----
        if ten_hang == "C2" or "C2" in ten_sp_upper:
            if sl_thuc_xuat > 0:
                for label, keyword in C2_TARGETS:
                    if keyword in ten_sp_upper:
                        mult = _unit_multiplier(don_vi)
                        c2_map[label]["chai"] += sl_thuc_xuat * mult
                        c2_map[label]["thanh_tien"] += thanh_tien
                        break

    if ngay_str is None:
        raise ValueError("Không tìm thấy dữ liệu ngày trong file.")

    btt_items = [
        {"ten": ten, "sl": v["sl"], "thanh_tien": v["thanh_tien"]}
        for ten, v in sorted(btt_map.items())
        if v["sl"] > 0
    ]
    c2_items = [
        {"ten": label, "chai": v["chai"], "thanh_tien": v["thanh_tien"]}
        for label, v in c2_map.items()
        if v["chai"] > 0
    ]

    nganh_hang_items = [
        {"ten": label, "thanh_tien": sum(nganh_hang_map.get(name, 0.0) for name in source_names)}
        for label, source_names in NGANH_HANG_GROUPS
    ]

    return {
        "ngay": ngay_str,
        "ten_st": ten_st or "—",
        "nam": {"doanh_thu": nam_total},
        "banh_trung_thu": {
            "items": btt_items,
            "tong_sl": sum(i["sl"] for i in btt_items),
            "tong_tien": sum(i["thanh_tien"] for i in btt_items),
        },
        "c2": {
            "items": c2_items,
            "tong_chai": sum(i["chai"] for i in c2_items),
            "tong_tien": sum(i["thanh_tien"] for i in c2_items),
        },
        "nganh_hang": {
            "items": nganh_hang_items,
            "tong_tien": sum(i["thanh_tien"] for i in nganh_hang_items),
        },
    }
