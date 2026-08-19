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
    conn.execute("""
        CREATE TABLE IF NOT EXISTS stock_snapshot (
            ten_st TEXT PRIMARY KEY,
            payload_json TEXT NOT NULL,
            gio TEXT
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


# ---------------------------------------------------------------------------
# TỒN KHO (BC tồn theo model) - mới thêm
# ---------------------------------------------------------------------------

def save_stock_snapshot(ten_st, payload, gio=None):
    """Lưu (hoặc ghi đè) snapshot tồn kho mới nhất của 1 siêu thị."""
    conn = _connect()
    with conn:
        conn.execute("""
            INSERT INTO stock_snapshot (ten_st, payload_json, gio)
            VALUES (?, ?, ?)
            ON CONFLICT(ten_st) DO UPDATE SET
                payload_json=excluded.payload_json,
                gio=excluded.gio
        """, (ten_st, json.dumps(payload, ensure_ascii=False), gio))
    conn.close()


def get_latest_stock_snapshot():
    """Trả về (ten_st, payload_dict, gio) của snapshot tồn kho mới nhất,
    hoặc (None, None, None) nếu chưa có dữ liệu."""
    conn = _connect()
    cur = conn.execute(
        "SELECT ten_st, payload_json, gio FROM stock_snapshot ORDER BY rowid DESC LIMIT 1"
    )
    row = cur.fetchone()
    conn.close()
    if not row:
        return None, None, None
    ten_st, payload_json, gio = row
    return ten_st, json.loads(payload_json), gio


# ---------------------------------------------------------------------------
# BÁO CÁO THƯỞNG (FRESH + FMCG) - mới thêm
# ---------------------------------------------------------------------------

def save_thuong_report(ten_st, payload, gio=None):
    """Lưu (ghi đè) báo cáo thưởng mới nhất của 1 siêu thị."""
    conn = _connect()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS thuong_reports (
            ten_st TEXT PRIMARY KEY,
            payload_json TEXT NOT NULL,
            gio TEXT
        )
    """)
    with conn:
        conn.execute("""
            INSERT INTO thuong_reports (ten_st, payload_json, gio)
            VALUES (?, ?, ?)
            ON CONFLICT(ten_st) DO UPDATE SET
                payload_json=excluded.payload_json,
                gio=excluded.gio
        """, (ten_st, json.dumps(payload, ensure_ascii=False), gio))
    conn.close()


def get_latest_thuong_report():
    """Trả về (ten_st, payload_dict, gio) báo cáo thưởng mới nhất,
    hoặc (None, None, None) nếu chưa có dữ liệu."""
    conn = _connect()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS thuong_reports (
            ten_st TEXT PRIMARY KEY,
            payload_json TEXT NOT NULL,
            gio TEXT
        )
    """)
    cur = conn.execute("SELECT ten_st, payload_json, gio FROM thuong_reports ORDER BY rowid DESC LIMIT 1")
    row = cur.fetchone()
    conn.close()
    if not row:
        return None, None, None
    ten_st, payload_json, gio = row
    return ten_st, json.loads(payload_json), gio


# ---------------------------------------------------------------------------
# NHẮC LỊCH HỖ TRỢ SIÊU THỊ KHÁC - mới thêm
# ---------------------------------------------------------------------------

def _connect_support():
    conn = _connect()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS group_members (
            user_id TEXT PRIMARY KEY,
            display_name TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS support_schedule (
            ngay TEXT PRIMARY KEY,
            ten TEXT NOT NULL,
            ca TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS support_reminder_log (
            ngay TEXT NOT NULL,
            gio_nhac TEXT NOT NULL,
            PRIMARY KEY (ngay, gio_nhac)
        )
    """)
    return conn


def save_group_members(members):
    """Ghi đè toàn bộ danh bạ thành viên nhóm. members: list các (user_id, display_name)."""
    conn = _connect_support()
    with conn:
        conn.execute("DELETE FROM group_members")
        conn.executemany(
            "INSERT INTO group_members (user_id, display_name) VALUES (?, ?)",
            members,
        )
    conn.close()


def get_all_group_members():
    conn = _connect_support()
    cur = conn.execute("SELECT user_id, display_name FROM group_members")
    rows = cur.fetchall()
    conn.close()
    return rows


def save_support_schedule_rows(rows):
    """Lưu (upsert) lịch hỗ trợ. rows: list dict {ngay, ten, ca}."""
    conn = _connect_support()
    with conn:
        for r in rows:
            conn.execute("""
                INSERT INTO support_schedule (ngay, ten, ca) VALUES (:ngay, :ten, :ca)
                ON CONFLICT(ngay) DO UPDATE SET ten=excluded.ten, ca=excluded.ca
            """, r)
    conn.close()


def ensure_default_schedule(default_rows):
    """Nếu bảng lịch hỗ trợ đang trống, nạp sẵn danh sách mặc định (default_rows)."""
    conn = _connect_support()
    cur = conn.execute("SELECT COUNT(*) FROM support_schedule")
    count = cur.fetchone()[0]
    conn.close()
    if count == 0:
        save_support_schedule_rows(default_rows)


def get_schedule_for_date(ngay):
    """Trả về (ten, ca) cho đúng ngày (YYYY-MM-DD), hoặc None nếu không có ai."""
    conn = _connect_support()
    cur = conn.execute("SELECT ten, ca FROM support_schedule WHERE ngay = ?", (ngay,))
    row = cur.fetchone()
    conn.close()
    return row if row else None


def da_nhac_chua(ngay, gio_nhac):
    """Kiểm tra đã gửi nhắc cho (ngay, gio_nhac) này chưa — tránh gửi trùng."""
    conn = _connect_support()
    cur = conn.execute(
        "SELECT 1 FROM support_reminder_log WHERE ngay = ? AND gio_nhac = ?", (ngay, gio_nhac)
    )
    row = cur.fetchone()
    conn.close()
    return row is not None


def danh_dau_da_nhac(ngay, gio_nhac):
    conn = _connect_support()
    with conn:
        conn.execute(
            "INSERT OR IGNORE INTO support_reminder_log (ngay, gio_nhac) VALUES (?, ?)",
            (ngay, gio_nhac),
        )
    conn.close()


# ---------------------------------------------------------------------------
# BÀI PHÂN LINE HÀNG NGÀY (THU NGÂN/FRESH/FMCG) - mới thêm
# ---------------------------------------------------------------------------

def _connect_phanline():
    conn = _connect()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS phan_line (
            ngay TEXT PRIMARY KEY,
            data_json TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS phan_line_reminder_log (
            ngay TEXT NOT NULL,
            slot TEXT NOT NULL,
            PRIMARY KEY (ngay, slot)
        )
    """)
    return conn


def save_phan_line(ngay, data):
    """Lưu (ghi đè) bài phân line của 1 ngày. data là dict."""
    conn = _connect_phanline()
    with conn:
        conn.execute("""
            INSERT INTO phan_line (ngay, data_json) VALUES (?, ?)
            ON CONFLICT(ngay) DO UPDATE SET data_json=excluded.data_json
        """, (ngay, json.dumps(data, ensure_ascii=False)))
    conn.close()


def get_phan_line(ngay):
    """Trả về dict bài phân line của ngày đó, hoặc None nếu chưa có."""
    conn = _connect_phanline()
    cur = conn.execute("SELECT data_json FROM phan_line WHERE ngay = ?", (ngay,))
    row = cur.fetchone()
    conn.close()
    return json.loads(row[0]) if row else None


def da_nhac_phanline(ngay, slot):
    conn = _connect_phanline()
    cur = conn.execute(
        "SELECT 1 FROM phan_line_reminder_log WHERE ngay = ? AND slot = ?", (ngay, slot)
    )
    row = cur.fetchone()
    conn.close()
    return row is not None


def danh_dau_da_nhac_phanline(ngay, slot):
    conn = _connect_phanline()
    with conn:
        conn.execute(
            "INSERT OR IGNORE INTO phan_line_reminder_log (ngay, slot) VALUES (?, ?)",
            (ngay, slot),
        )
    conn.close()
