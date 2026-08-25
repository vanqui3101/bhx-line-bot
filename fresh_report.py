"""
fresh_report.py - Xử lý 4 tính năng dùng chung 1 file Excel "hủy tồn + MMKK"
(riêng biệt, khác hẳn 3 file DT/MTKM/TD cũ):

1. Doanh thu thủy hải sản — báo cáo riêng ngành Thủy hải sản.
2. Công việc 1 — chi tiết Hủy tồn + MMKK từng sản phẩm (đúng 1 ngày).
3. Công việc 2 — tổng Hủy tồn + MMKK theo 4 nhóm lớn + Trứng (nhiều ngày).
4. Phân tích số liệu — phân tích trực tiếp dựa trên đúng ngày được hỏi.

Dữ liệu đọc từ storage.get_fresh_records_by_date() / get_fresh_records_range(),
đã được lưu qua excel_reader.read_fresh_rows() khi anh gửi file cho bot.
"""
import re
import calendar
import unicodedata
from datetime import datetime, date, timedelta

import storage
from excel_reader import FRESH_NHOM_LON
from flex_builder import (
    build_seafood_flex_message,
    build_fresh_detail_flex_message,
    build_fresh_group_flex_message,
)

# Sản phẩm loại trừ khỏi bảng Hủy tồn — số liệu bất thường, không phải hao hụt
# bán lẻ thường ngày (thịt heo nửa mảnh nhập nguyên con, không cùng bản chất
# với hủy tồn rau củ/thịt/thủy hải sản lẻ).
SAN_PHAM_LOAI_TRU_HUY_TON = {"HEO NỬA MẢNH"}

NHOM_LON_THU_TU = ["Rau củ", "Trái cây", "Thịt", "Thủy hải sản", "Trứng"]


def _today_vn():
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("Asia/Ho_Chi_Minh")).date()
    except Exception:
        return datetime.now().date()


def _fmt_dm(d):
    return d.strftime("%d/%m")


def _fmt_dmy(d):
    return d.strftime("%d/%m/%Y")


def _ngay_str(d):
    return d.strftime("%Y-%m-%d")


def _bo_dau(text):
    """Bỏ dấu tiếng Việt (và hạ chữ thường) để nhận diện câu lệnh không phụ
    thuộc cách gõ dấu, vd "hôm qua"/"hom qua", "hủy"/"huỷ" đều so khớp được."""
    text = text or ""
    nfkd = unicodedata.normalize("NFKD", text)
    khong_dau = "".join(c for c in nfkd if not unicodedata.combining(c))
    khong_dau = khong_dau.replace("đ", "d").replace("Đ", "D")
    return khong_dau.lower()


# ---------------------------------------------------------------------------
# Nhận diện ngày/khoảng ngày từ câu lệnh tự do (so khớp trên bản KHÔNG DẤU)
# ---------------------------------------------------------------------------

_SO_NGAY_TRUOC_RE = re.compile(r"(\d+)\s*ngay\s*truoc")
# Ngày đầu có thể chỉ ghi số ngày (không kèm tháng), mượn tháng/năm từ ngày sau
# (vd "từ ngày 6 đến 10/8" nghĩa là 6/8 -> 10/8).
_KHOANG_NGAY_RE = re.compile(
    r"tu\s*ngay\s*(\d{1,2})(?:[/\-](\d{1,2})(?:[/\-](\d{2,4}))?)?"
    r"\s*(?:den|->|-)\s*(?:ngay\s*)?(\d{1,2})[/\-](\d{1,2})(?:[/\-](\d{2,4}))?"
)


def parse_date_request(text, today=None):
    """Phân tích câu lệnh tự do, trả về (mode, ngay_tu, ngay_den):
    - mode == "single": 1 ngày cụ thể (dùng cho Công việc 1 / Phân tích)
    - mode == "range":  nhiều ngày (dùng cho Công việc 2)
    Nếu không nhận diện được ngày cụ thể nào trong câu, mặc định là "hôm qua"."""
    if today is None:
        today = _today_vn()
    t = _bo_dau(text)

    m = _KHOANG_NGAY_RE.search(t)
    if m:
        d1, mo1, y1, d2, mo2, y2 = m.groups()
        # Ngày đầu thiếu tháng/năm -> mượn từ ngày sau (cùng tháng/năm).
        mo1 = mo1 or mo2
        y1 = y1 or y2 or today.year
        y2 = y2 or today.year
        y1, y2 = int(y1), int(y2)
        if y1 < 100:
            y1 += 2000
        if y2 < 100:
            y2 += 2000
        try:
            ngay_tu = date(y1, int(mo1), int(d1))
            ngay_den = date(y2, int(mo2), int(d2))
        except ValueError:
            return "single", today - timedelta(days=1), today - timedelta(days=1)
        if ngay_tu > ngay_den:
            ngay_tu, ngay_den = ngay_den, ngay_tu
        return "range", ngay_tu, ngay_den

    m = _SO_NGAY_TRUOC_RE.search(t)
    if m:
        n = int(m.group(1))
        ngay_tu = today - timedelta(days=n)
        ngay_den = today - timedelta(days=1)
        return "range", ngay_tu, ngay_den

    if "hom nay" in t:
        return "single", today, today

    # Mặc định (kể cả khi câu chỉ có "hủy mmkk" không kèm ngày): hôm qua
    y = today - timedelta(days=1)
    return "single", y, y


def wants_comparison(text):
    """Câu có yêu cầu so sánh với ngày hôm qua/hôm trước không (chỉ dùng cho
    lệnh Phân tích số liệu — mặc định KHÔNG so sánh, trừ khi anh nói rõ)."""
    t = _bo_dau(text)
    return "so sanh" in t and ("hom qua" in t or "hom truoc" in t)


# ---------------------------------------------------------------------------
# Helper tính kg / đơn vị gốc cho 1 dòng fresh_records
# ---------------------------------------------------------------------------

def _kg(row, field):
    return (row.get(field) or 0) * (row.get("dvt") or 0)


def _qty_text(row, field):
    """Trả về (giá_trị_so_sánh, chuỗi hiển thị) theo đúng đơn vị gốc sản phẩm."""
    raw = row.get(field) or 0
    if row.get("don_vi") == "kg":
        val = round(raw * (row.get("dvt") or 0), 2)
        return val, f"{val:g} kg"
    val = raw
    if val == int(val):
        val = int(val)
    return raw, f"{val} {row.get('don_vi')}"


def _is_noise(sortval, don_vi):
    """Bỏ dòng gần như 0 do sai số làm tròn (chỉ áp dụng cho hàng cân kg)."""
    return don_vi == "kg" and sortval <= 0.005


# ---------------------------------------------------------------------------
# CÔNG VIỆC 1 — chi tiết Hủy tồn + MMKK từng sản phẩm (1 ngày)
# ---------------------------------------------------------------------------

def build_cong_viec_1(ten_st, ngay):
    """Trả về flex bubble (dict) hoặc None nếu chưa có dữ liệu ngày này."""
    rows = storage.get_fresh_records_by_date(_ngay_str(ngay))
    if not rows:
        return None

    def collect(field, loai_tru_heo):
        by_nganh = {}
        for r in rows:
            if loai_tru_heo and r["ten_sp"].strip().upper() in SAN_PHAM_LOAI_TRU_HUY_TON:
                continue
            sortval, text = _qty_text(r, field)
            if not sortval or sortval <= 0:
                continue
            if _is_noise(sortval, r.get("don_vi")):
                continue
            by_nganh.setdefault(r["nganh_hang"] or "Khác", []).append((r["ten_sp"], text))
        grouped = []
        total = 0
        for nganh_hang in sorted(by_nganh.keys()):
            items = sorted(by_nganh[nganh_hang], key=lambda x: x[0])
            grouped.append((nganh_hang, items))
            total += len(items)
        return grouped, total

    huy_groups, so_huy = collect("sl_huy", loai_tru_heo=True)
    mmkk_groups, so_mmkk = collect("sl_mmkk", loai_tru_heo=False)

    return build_fresh_detail_flex_message(
        ten_st, _fmt_dmy(ngay), huy_groups, mmkk_groups, so_huy, so_mmkk
    )


# ---------------------------------------------------------------------------
# CÔNG VIỆC 2 — tổng Hủy tồn + MMKK theo nhóm lớn (nhiều ngày)
# ---------------------------------------------------------------------------

def build_cong_viec_2(ten_st, ngay_tu, ngay_den):
    rows = storage.get_fresh_records_range(_ngay_str(ngay_tu), _ngay_str(ngay_den))
    if not rows:
        return None

    huy_by_nhom = {n: 0.0 for n in NHOM_LON_THU_TU}
    mmkk_by_nhom = {n: 0.0 for n in NHOM_LON_THU_TU}
    huy_don_vi = {"Rau củ": "kg", "Trái cây": "kg", "Thịt": "kg", "Thủy hải sản": "kg", "Trứng": "hộp"}

    for r in rows:
        nhom = FRESH_NHOM_LON.get(r["nganh_hang"])
        if not nhom:
            continue
        loai_tru = r["ten_sp"].strip().upper() in SAN_PHAM_LOAI_TRU_HUY_TON
        if nhom == "Trứng":
            if not loai_tru:
                huy_by_nhom[nhom] += r.get("sl_huy") or 0
            mmkk_by_nhom[nhom] += r.get("sl_mmkk") or 0
        else:
            if not loai_tru:
                huy_by_nhom[nhom] += _kg(r, "sl_huy")
            mmkk_by_nhom[nhom] += _kg(r, "sl_mmkk")

    nhom_rows = []
    for nhom in NHOM_LON_THU_TU:
        dv = huy_don_vi[nhom]
        huy_val = round(huy_by_nhom[nhom], 2)
        mmkk_val = round(mmkk_by_nhom[nhom], 2)
        tong_val = round(huy_val + mmkk_val, 2)
        if dv == "kg":
            huy_text, mmkk_text, tong_text = f"{huy_val:g} kg", f"{mmkk_val:g} kg", f"{tong_val:g} kg"
        else:
            huy_text, mmkk_text, tong_text = f"{huy_val:g} hộp", f"{mmkk_val:g} hộp", f"{tong_val:g} hộp"
        nhom_rows.append((nhom, huy_text, mmkk_text, tong_text))

    so_ngay = (ngay_den - ngay_tu).days + 1
    return build_fresh_group_flex_message(ten_st, _fmt_dm(ngay_tu), _fmt_dm(ngay_den), so_ngay, nhom_rows)


# ---------------------------------------------------------------------------
# DOANH THU THỦY HẢI SẢN
# ---------------------------------------------------------------------------

def build_doanh_thu_thuy_hai_san(ten_st):
    """Dùng TOÀN BỘ dữ liệu FRESH đã lưu (từ ngày sớm nhất đến ngày mới nhất)."""
    dates = storage.get_fresh_distinct_dates()
    if not dates:
        return None
    ngay_tu = datetime.strptime(min(dates), "%Y-%m-%d").date()
    ngay_den = datetime.strptime(max(dates), "%Y-%m-%d").date()
    rows = storage.get_fresh_records_range(_ngay_str(ngay_tu), _ngay_str(ngay_den))

    by_sp = {}
    for r in rows:
        if r["nganh_hang"] not in ("Thủy Hải Sản Tập Trung", "Thủy Hải Sản Nhập Khẩu"):
            continue
        d = by_sp.setdefault(r["ten_sp"], {"nhap": 0.0, "xuat": 0.0, "huy": 0.0, "mmkk": 0.0, "tien": 0.0})
        d["nhap"] += _kg(r, "sl_nhap")
        d["xuat"] += _kg(r, "sl_xuat")
        d["huy"] += _kg(r, "sl_huy")
        d["mmkk"] += _kg(r, "sl_mmkk")
        d["tien"] += r.get("thanh_tien") or 0

    if not by_sp:
        return None

    items = []
    doanh_thu_tong = 0.0
    gia_tri_huy = 0.0
    for ten_sp, d in sorted(by_sp.items(), key=lambda kv: -kv[1]["xuat"]):
        pct = (d["xuat"] / d["nhap"] * 100) if d["nhap"] > 0 else None
        items.append((ten_sp, f"{d['xuat']:.2f}", pct))
        doanh_thu_tong += d["tien"]
        gia_binh_quan = (d["tien"] / d["xuat"]) if d["xuat"] > 0 else 0
        gia_tri_huy += gia_binh_quan * (d["huy"] + d["mmkk"])

    so_ngay_da_qua = (ngay_den - ngay_tu).days + 1
    so_ngay_ca_thang = calendar.monthrange(ngay_den.year, ngay_den.month)[1]
    he_so = so_ngay_ca_thang / so_ngay_da_qua if so_ngay_da_qua else 1.0
    du_kien_cuoi_thang = doanh_thu_tong * he_so

    return build_seafood_flex_message(
        ten_st, _fmt_dmy(ngay_tu), _fmt_dmy(ngay_den), so_ngay_da_qua,
        items, doanh_thu_tong, gia_tri_huy, du_kien_cuoi_thang,
    )


# ---------------------------------------------------------------------------
# PHÂN TÍCH SỐ LIỆU — mặc định chỉ phân tích ĐÚNG ngày được hỏi, không tự so
# sánh với ngày hôm qua/trước đó trừ khi được yêu cầu rõ.
# ---------------------------------------------------------------------------

def _tong_theo_nhom(rows):
    huy = {n: 0.0 for n in NHOM_LON_THU_TU}
    mmkk = {n: 0.0 for n in NHOM_LON_THU_TU}
    for r in rows:
        nhom = FRESH_NHOM_LON.get(r["nganh_hang"])
        if not nhom:
            continue
        loai_tru = r["ten_sp"].strip().upper() in SAN_PHAM_LOAI_TRU_HUY_TON
        if nhom == "Trứng":
            if not loai_tru:
                huy[nhom] += r.get("sl_huy") or 0
            mmkk[nhom] += r.get("sl_mmkk") or 0
        else:
            if not loai_tru:
                huy[nhom] += _kg(r, "sl_huy")
            mmkk[nhom] += _kg(r, "sl_mmkk")
    return huy, mmkk


def build_phan_tich(ngay, so_sanh_voi_hom_qua=False):
    """Trả về chuỗi text phân tích, hoặc None nếu chưa có dữ liệu ngày này."""
    rows = storage.get_fresh_records_by_date(_ngay_str(ngay))
    if not rows:
        return None

    huy, mmkk = _tong_theo_nhom(rows)
    tong = {n: round(huy[n] + mmkk[n], 2) for n in NHOM_LON_THU_TU}

    # Nhóm có vấn đề lớn nhất (rau củ/trái cây/thịt/thủy hải sản theo kg;
    # trứng không cộng chung do khác đơn vị).
    nhom_kg = {n: tong[n] for n in NHOM_LON_THU_TU if n != "Trứng"}
    nhom_max = max(nhom_kg, key=nhom_kg.get) if nhom_kg else None

    # Sản phẩm MMKK lớn nhất trong ngày -> chỉ đúng nguyên nhân cụ thể
    sp_max = None
    sp_max_val = 0
    for r in rows:
        _, mmkk_kg = _kg(r, "sl_mmkk"), _kg(r, "sl_mmkk")
        val = _kg(r, "sl_mmkk") if r.get("don_vi") == "kg" else (r.get("sl_mmkk") or 0)
        if val > sp_max_val:
            sp_max_val = val
            sp_max = r

    lines = [f"PHÂN TÍCH HỦY + MMKK — {_fmt_dmy(ngay)}", ""]
    for n in NHOM_LON_THU_TU:
        dv = "hộp" if n == "Trứng" else "kg"
        lines.append(f"{n}: hủy {huy[n]:.2f} {dv}, MMKK {mmkk[n]:.2f} {dv}, tổng {tong[n]:.2f} {dv}")
    lines.append("")

    if nhom_max and nhom_kg[nhom_max] > 0:
        lines.append(f"Nhóm đáng chú ý nhất: {nhom_max} ({nhom_kg[nhom_max]:.2f} kg).")
    if sp_max is not None and sp_max_val > 0:
        dv = sp_max.get("don_vi")
        val_text = f"{sp_max_val:.2f} kg" if dv == "kg" else f"{int(sp_max_val) if sp_max_val==int(sp_max_val) else sp_max_val} {dv}"
        chiem_pct = (sp_max_val / nhom_kg[FRESH_NHOM_LON.get(sp_max['nganh_hang'], '')] * 100) if FRESH_NHOM_LON.get(sp_max['nganh_hang']) in nhom_kg and nhom_kg[FRESH_NHOM_LON.get(sp_max['nganh_hang'])] > 0 else None
        note = f" (chiếm {chiem_pct:.0f}% MMKK+hủy của nhóm {FRESH_NHOM_LON.get(sp_max['nganh_hang'])})" if chiem_pct else ""
        lines.append(f"Nguyên nhân cụ thể: {sp_max['ten_sp']} một mình chiếm {val_text} MMKK/hủy{note}. "
                      f"Đây là vấn đề ở đúng 1 sản phẩm, không phải xu hướng chung — cần kiểm tra lại: cân sai lúc nhập, "
                      f"nhập sai số liệu, hay hao hụt/mất hàng thật.")

    if so_sanh_voi_hom_qua:
        ngay_truoc = ngay - timedelta(days=1)
        rows_truoc = storage.get_fresh_records_by_date(_ngay_str(ngay_truoc))
        if rows_truoc:
            huy_t, mmkk_t = _tong_theo_nhom(rows_truoc)
            tong_t = {n: round(huy_t[n] + mmkk_t[n], 2) for n in NHOM_LON_THU_TU}
            lines.append("")
            lines.append(f"So với {_fmt_dmy(ngay_truoc)}:")
            for n in NHOM_LON_THU_TU:
                if n == "Trứng":
                    continue
                delta = tong[n] - tong_t[n]
                lines.append(f"{n}: {tong_t[n]:.2f} kg -> {tong[n]:.2f} kg ({delta:+.2f} kg)")
        else:
            lines.append("")
            lines.append(f"(Chưa có dữ liệu ngày {_fmt_dmy(ngay_truoc)} để so sánh.)")

    return "\n".join(lines)
