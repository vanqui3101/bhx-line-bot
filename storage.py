"""
storage.py - Lưu trữ dữ liệu doanh thu theo ngày bằng SQLite.

LƯU Ý: trên gói Railway miễn phí, ổ đĩa không được lưu vĩnh viễn qua các lần
deploy lại (redeploy) code. Dữ liệu vẫn giữ nguyên khi bot chạy bình thường,
chỉ mất khi anh chủ động deploy code mới. Nếu cần lưu vĩnh viễn, có thể nâng
cấp sau bằng Railway Volume hoặc Google Sheets.

Có 2 nhóm dữ liệu:
1. records / snapshot_times  -> báo cáo DOANH THU (giữ nguyên như code cũ)
2. category_reports          -> báo cáo NGÀNH HÀNG (Nấm / Bánh trung thu / C2),
   mới thêm để phục vụ lệnh "MỤC TIÊU KHUYẾN MÃI"
"""

import sqlite3
import os
import json
import calendar
from datetime import datetime, date

DB_PATH = os.path.join(os.path.dirname(__file__), "data", "bot.db")


def _connect():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS records (
            ngay TEXT NOT NULL,
            ma_st TEXT NOT NULL,
            ten_st TEXT,
            tinh TEXT,
            dt_offline REAL,
            dt_online REAL,
            bill_offline REAL,
            bill_online REAL,
            PRIMARY KEY (ngay, ma_st)
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS snapshot_times (
            ngay TEXT PRIMARY KEY,
            gio TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS category_reports (
            ngay TEXT NOT NULL,
            ten_st TEXT NOT NULL,
            payload_json TEXT NOT NULL,
            gio TEXT,
            PRIMARY KEY (ngay, ten_st)
        )
    """)
    return conn


# ---------------------------------------------------------------------------
# BÁO CÁO DOANH THU (giữ nguyên như code cũ)
# ---------------------------------------------------------------------------

def save_records(rows):
    """Lưu (hoặc cập nhật) danh sách bản ghi doanh thu theo ngày."""
    conn = _connect()
    with conn:
        for r in rows:
            conn.execute("""
                INSERT INTO records (ngay, ma_st, ten_st, tinh, dt_offline, dt_online, bill_offline, bill_online)
                VALUES (:ngay, :ma_st, :ten_st, :tinh, :dt_offline, :dt_online, :bill_offline, :bill_online)
                ON CONFLICT(ngay, ma_st) DO UPDATE SET
                    ten_st=excluded.ten_st,
                    tinh=excluded.tinh,
                    dt_offline=excluded.dt_offline,
                    dt_online=excluded.dt_online,
                    bill_offline=excluded.bill_offline,
                    bill_online=excluded.bill_online
            """, r)
    conn.close()


def get_distinct_dates():
    """Trả về danh sách các ngày có dữ liệu, sắp xếp mới nhất trước."""
    conn = _connect()
    cur = conn.execute("SELECT DISTINCT ngay FROM records ORDER BY ngay DESC")
    dates = [row[0] for row in cur.fetchall()]
    conn.close()
    return dates


def get_records_by_date(ngay):
    conn = _connect()
    cur = conn.execute("SELECT * FROM records WHERE ngay = ?", (ngay,))
    cols = [d[0] for d in cur.description]
    rows = [dict(zip(cols, row)) for row in cur.fetchall()]
    conn.close()
    return rows


def save_snapshot_time(ngay, gio):
    """Lưu giờ mà dữ liệu ngày này được gửi vào bot (để hiển thị khung giờ so sánh)."""
    conn = _connect()
    with conn:
        conn.execute("""
            INSERT INTO snapshot_times (ngay, gio) VALUES (?, ?)
            ON CONFLICT(ngay) DO UPDATE SET gio=excluded.gio
        """, (ngay, gio))
    conn.close()


def get_snapshot_time(ngay):
    if not ngay:
        return None
    conn = _connect()
    cur = conn.execute("SELECT gio FROM snapshot_times WHERE ngay = ?", (ngay,))
    row = cur.fetchone()
    conn.close()
    return row[0] if row else None


def _same_day_last_month(ngay_str):
    """Trả về ngày cùng số ngày của tháng trước (vd 15/08 -> 15/07).
    Nếu tháng trước không có ngày đó (vd 31 -> tháng thiếu), lấy ngày cuối tháng trước."""
    d = datetime.strptime(ngay_str, "%Y-%m-%d").date()
    year, month = d.year, d.month - 1
    if month == 0:
        month = 12
        year -= 1
    last_day_of_prev_month = calendar.monthrange(year, month)[1]
    day = min(d.day, last_day_of_prev_month)
    return date(year, month, day).strftime("%Y-%m-%d")


def get_latest_and_previous():
    """Trả về (ngay_moi_nhat, records_moi_nhat, ngay_so_sanh, records_so_sanh).
    Kỳ so sánh là CÙNG NGÀY của THÁNG TRƯỚC (vd 15/08 so với 15/07).
    ngay_so_sanh và records_so_sanh sẽ là None/[] nếu chưa có dữ liệu ngày đó."""
    dates = get_distinct_dates()
    if not dates:
        return None, [], None, []
    latest = dates[0]
    latest_records = get_records_by_date(latest)

    target_prev = _same_day_last_month(latest)
    if target_prev in dates:
        prev_records = get_records_by_date(target_prev)
        return latest, latest_records, target_prev, prev_records

    return latest, latest_records, None, []


# ---------------------------------------------------------------------------
# BÁO CÁO NGÀNH HÀNG (Nấm / Bánh trung thu / C2) - mới thêm
# ---------------------------------------------------------------------------

def save_category_report(ngay, ten_st, payload, gio=None):
    """Lưu (hoặc cập nhật) 1 báo cáo ngành hàng theo (ngày, siêu thị)."""
    conn = _connect()
    with conn:
        conn.execute("""
            INSERT INTO category_reports (ngay, ten_st, payload_json, gio)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(ngay, ten_st) DO UPDATE SET
                payload_json=excluded.payload_json,
                gio=excluded.gio
        """, (ngay, ten_st, json.dumps(payload, ensure_ascii=False), gio))
    conn.close()


def get_latest_category_report():
    """Trả về (ngay, ten_st, payload_dict, gio) của báo cáo ngành hàng mới nhất,
    hoặc (None, None, None, None) nếu chưa có dữ liệu."""
    conn = _connect()
    cur = conn.execute(
        "SELECT ngay, ten_st, payload_json, gio FROM category_reports "
        "ORDER BY ngay DESC, rowid DESC LIMIT 1"
    )
    row = cur.fetchone()
    conn.close()
    if not row:
        return None, None, None, None
    ngay, ten_st, payload_json, gio = row
    return ngay, ten_st, json.loads(payload_json), gio
