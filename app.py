# ============================================================
# APP.PY - MINI CRM THU CƯỚC KHDN
# Phiên bản single-file để dễ deploy lên Streamlit Cloud
# Không dùng src/, không dùng YAML, không phụ thuộc cấu trúc thư mục
# ============================================================

import io
import re
import sqlite3
import unicodedata
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pandas as pd
import streamlit as st


# ============================================================
# KHỐI 1: CẤU HÌNH HỆ THỐNG - ZERO HARDCODING
# ============================================================
# Lý do:
# - Những thứ dễ thay đổi như tên cột Excel, mẫu tin nhắn, trạng thái, hạn thanh toán
#   không nên viết chết rải rác trong logic.
# - Khi công ty đổi tên cột hoặc đổi mẫu tin, bạn sửa APP_CONFIG là chính.
# - Đây là cách giữ tư duy kiến trúc tách lớp ngay cả khi chỉ dùng 1 file app.py.

APP_CONFIG = {
    "app_title": "📞 Mini-CRM Thu Cước KHDN",
    "db_path": "crm_thu_cuoc_khdn.db",

    "staff": {
        "unit_name": "VNPT Long Thành- Nhơn Trạch",
        "staff_name": "Thuận",
        "staff_phone": "0837892579",
        "website": "https://vnptdongnai.vn/",
        "payment_deadline": "ngày 15",
    },

    "message_templates": {
        "general": """{unit_name} thông báo:
Đã có thông báo cước kỳ cước tháng {billing_period}. Anh/chị truy cập trang {website} để lấy thông báo cước và hóa đơn của công ty.
Anh/chị vui lòng thanh toán cước trước {payment_deadline}. Sau ngày 15 hệ thống sẽ tự động đưa lưới khóa khi ghi nhận còn tồn nợ.
Liên hệ nhân viên kinh doanh: {staff_name}: {staff_phone} để được hỗ trợ gạch nợ sớm nhất sau khi thanh toán.
Nếu đã thanh toán cước, anh/chị vui lòng bỏ qua tin nhắn trên.
Trân trọng""",

        "detail_after_verified": """{unit_name} thông báo:
Dạ em gửi anh/chị thông tin cước kỳ {billing_period} của {group_name}.
Tổng tiền còn cần thanh toán theo hệ thống: {total_debt_display}.
Số lượng mã thanh toán: {ma_tt_count}.
Anh/chị vui lòng thanh toán cước trước {payment_deadline}. Sau ngày 15 hệ thống sẽ tự động đưa lưới khóa khi ghi nhận còn tồn nợ.
Liên hệ nhân viên kinh doanh: {staff_name}: {staff_phone} để được hỗ trợ gạch nợ sớm nhất sau khi thanh toán.
Nếu đã thanh toán cước, anh/chị vui lòng bỏ qua tin nhắn trên.
Trân trọng"""
    },

    "status": {
        "paid": {
            "label": "Đã đóng",
            "emoji": "🟢",
        },
        "contacted_unpaid": {
            "label": "Đã liên hệ nhưng chưa đóng",
            "emoji": "🟡",
        },
        "uncontacted": {
            "label": "Chưa liên hệ / thiếu số",
            "emoji": "🔴",
        },
        "need_check": {
            "label": "Cần kiểm tra",
            "emoji": "⚪",
        },
    },

    "contact_results": [
        "Đã gọi - hẹn thanh toán",
        "Đã gọi - xin số kế toán mới",
        "Đã gọi - khách báo đã đóng",
        "Không nghe máy",
        "Sai số điện thoại",
        "Không có số điện thoại",
        "Cần kiểm tra lại",
    ],

    # Mapping tên cột linh hoạt.
    # Bên trái là tên chuẩn trong hệ thống.
    # Bên phải là các tên có thể xuất hiện trong Excel.
    "column_candidates": {
        "tn08": {
            "required": {
                "ma_tt": [
                    "ma_tt", "mã tt", "ma tt", "mã thanh toán", "ma thanh toan",
                    "mã thanh toán", "mã thanh toán"
                ],
                "debt_amount": [
                    "total nợ thu vét", "total no thu vet", "total nợ thu vét",
                    "tiền nợ", "tien no", "nợ còn lại", "no con lai",
                    "tổng nợ", "tong no", "total"
                ],
            },
            "optional": {
                "customer_name": [
                    "tên khách hàng", "ten khach hang", "tên khách hàng",
                    "tên thanh toán", "ten thanh toan", "ten_tt", "tên tt",
                    "khách hàng", "khach hang"
                ],
                "address": [
                    "địa chỉ", "dia chi", "địa chỉ kh", "địa chỉ kh",
                    "địa chỉ thanh toán", "dia chi thanh toan"
                ],
                "phone": [
                    "số điện thoại", "so dien thoai", "số dt liên hệ",
                    "sdt", "điện thoại", "dien thoai", "phone"
                ],
                "representative_code": [
                    "mã đại diện", "ma dai dien", "ma_dd", "mã đd",
                    "mã nhóm", "ma nhom"
                ],
            },
        },

        "paid": {
            "required": {
                "ma_tt": [
                    "ma_tt", "mã tt", "ma tt", "mã thanh toán", "ma thanh toan",
                    "mã thanh toán", "mã thanh toán"
                ],
            },
            "optional": {
                "paid_amount": [
                    "số tiền", "so tien", "tiền đóng", "tien dong",
                    "số tiền đã đóng", "so tien da dong", "amount"
                ],
                "paid_date": [
                    "ngày đóng", "ngay dong", "ngày thanh toán",
                    "ngay thanh toan", "payment date"
                ],
            },
        },
    },

    # Regex nhận diện MA_TT trong file nhóm kiểu file Giang.
    # Có thể mở rộng nếu công ty có format mã khác.
    "ma_tt_regex": r"[A-Z]{2,5}-\d{1,3}-\d{5,10}",
}


# ============================================================
# KHỐI 2: TIỆN ÍCH CHUẨN HÓA DỮ LIỆU
# ============================================================

def normalize_text(value) -> str:
    """
    Chuẩn hóa text an toàn.
    None/NaN/NaT đều trả về chuỗi rỗng.
    Tránh lỗi pandas NaN bị đổi thành chuỗi "nan" làm app hiểu nhầm là đã liên hệ.
    """
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    if text.lower() in ["nan", "nat", "none", "null"]:
        return ""
    text = re.sub(r"\s+", " ", text)
    return text
def remove_vietnamese_accents(text: str) -> str:
    """Bỏ dấu tiếng Việt để so khớp tên cột linh hoạt."""
    text = normalize_text(text).lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return text


def normalize_column_key(value) -> str:
    """Chuẩn hóa tên cột để mapping không phụ thuộc hoa/thường/dấu/khoảng trắng."""
    text = remove_vietnamese_accents(value)
    text = re.sub(r"[^a-z0-9]+", "_", text).strip("_")
    return text


def normalize_ma_tt(value) -> str:
    """Chuẩn hóa MA_TT thành khóa ổn định."""
    text = normalize_text(value)
    if text.endswith(".0"):
        text = text[:-2]
    return text.upper()


def normalize_phone(value) -> str:
    """Chuẩn hóa số điện thoại/Zalo ở mức an toàn."""
    text = normalize_text(value)
    if text.endswith(".0"):
        text = text[:-2]
    digits = re.sub(r"\D", "", text)
    if not digits:
        return ""
    if len(digits) == 9:
        digits = "0" + digits
    return digits


def parse_money(value) -> float:
    """Chuyển dữ liệu tiền về số float, chịu được định dạng VN."""
    if value is None or value == "":
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    text = normalize_text(value)
    text = text.replace("đ", "").replace("Đ", "").replace("VND", "").replace("vnd", "")
    text = text.replace(" ", "")

    # Nếu có cả dấu chấm và phẩy, giả định kiểu VN: 1.234.567,89
    if "." in text and "," in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        # Nếu chỉ có dấu phẩy, thường là phân cách nghìn hoặc thập phân.
        # Với tiền cước, ưu tiên bỏ phân cách nghìn.
        text = text.replace(",", "")

    try:
        return float(Decimal(text))
    except (InvalidOperation, ValueError):
        return 0.0


def format_money(value) -> str:
    """Format tiền kiểu Việt Nam."""
    try:
        return f"{float(value):,.0f} đồng".replace(",", ".")
    except Exception:
        return "0 đồng"


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def month_default() -> str:
    """Mặc định kỳ cước là tháng trước vì thực tế thường thu kỳ trước."""
    today = datetime.now()
    month = today.month - 1
    year = today.year
    if month == 0:
        month = 12
        year -= 1
    return f"{month:02d}/{year}"


def group_key_from_text(customer_name: str, address: str, representative_code: str = "") -> str:
    """
    Tạo group_key tự động.
    Ưu tiên mã đại diện nếu có.
    Nếu không có, dùng tên khách + địa chỉ đã chuẩn hóa.
    """
    rep = normalize_text(representative_code)
    if rep:
        return "REP|" + normalize_column_key(rep)

    name_key = normalize_column_key(customer_name)
    address_key = normalize_column_key(address)
    if name_key:
        return "NAME|" + name_key + "|" + address_key[:80]

    return "UNKNOWN"


# ============================================================
# KHỐI 3: DATABASE SQLITE
# ============================================================
# Lý do:
# - Không lưu mọi thứ trong DataFrame tạm, vì refresh app sẽ mất trạng thái.
# - SQLite nhẹ, chạy được trên Streamlit Cloud.
# - Thiết kế có batch import để không ghi đè mất lịch sử.

def get_conn():
    conn = sqlite3.connect(APP_CONFIG["db_path"], check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(conn):
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS import_batches (
            batch_id INTEGER PRIMARY KEY AUTOINCREMENT,
            import_type TEXT NOT NULL,
            file_name TEXT,
            row_count INTEGER DEFAULT 0,
            total_amount REAL DEFAULT 0,
            note TEXT,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS customer_groups (
            group_id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_key TEXT UNIQUE NOT NULL,
            group_name TEXT,
            address TEXT,
            primary_phone TEXT,
            source TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS group_members (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_id INTEGER NOT NULL,
            ma_tt TEXT NOT NULL,
            is_active INTEGER DEFAULT 1,
            created_at TEXT NOT NULL,
            UNIQUE(group_id, ma_tt)
        );

        CREATE TABLE IF NOT EXISTS debt_snapshot_lines (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_id INTEGER NOT NULL,
            ma_tt TEXT NOT NULL,
            group_id INTEGER,
            billing_period TEXT,
            customer_name TEXT,
            address TEXT,
            phone TEXT,
            representative_code TEXT,
            debt_amount REAL DEFAULT 0,
            created_at TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_debt_batch_ma_tt
        ON debt_snapshot_lines(batch_id, ma_tt);

        CREATE INDEX IF NOT EXISTS idx_debt_group
        ON debt_snapshot_lines(group_id);

        CREATE TABLE IF NOT EXISTS paid_updates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_id INTEGER NOT NULL,
            ma_tt TEXT NOT NULL,
            paid_amount REAL DEFAULT 0,
            paid_date TEXT,
            created_at TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_paid_ma_tt
        ON paid_updates(ma_tt);

        CREATE TABLE IF NOT EXISTS interactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_id INTEGER NOT NULL,
            ma_tt TEXT,
            result TEXT NOT NULL,
            promised_payment_date TEXT,
            note TEXT,
            created_by TEXT,
            created_at TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_interactions_group
        ON interactions(group_id);

        CREATE TABLE IF NOT EXISTS contacts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_id INTEGER NOT NULL,
            contact_value TEXT NOT NULL,
            contact_person TEXT,
            role TEXT,
            is_primary INTEGER DEFAULT 1,
            verification_status TEXT DEFAULT 'verified',
            note TEXT,
            created_at TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_contacts_group
        ON contacts(group_id);

        CREATE TABLE IF NOT EXISTS sent_messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_id INTEGER NOT NULL,
            template_key TEXT NOT NULL,
            message_text TEXT NOT NULL,
            created_by TEXT,
            created_at TEXT NOT NULL
        );
        """
    )
    conn.commit()


def create_batch(conn, import_type: str, file_name: str, row_count: int, total_amount: float = 0.0, note: str = "") -> int:
    cur = conn.execute(
        """
        INSERT INTO import_batches
        (import_type, file_name, row_count, total_amount, note, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (import_type, file_name, row_count, total_amount, note, now_str()),
    )
    conn.commit()
    return int(cur.lastrowid)


def get_latest_batch_id(conn, import_type: str):
    row = conn.execute(
        """
        SELECT batch_id
        FROM import_batches
        WHERE import_type = ?
        ORDER BY batch_id DESC
        LIMIT 1
        """,
        (import_type,),
    ).fetchone()
    return int(row["batch_id"]) if row else None


def get_or_create_group(conn, group_key: str, group_name: str, address: str = "", phone: str = "", source: str = "auto") -> int:
    existing = conn.execute(
        "SELECT group_id FROM customer_groups WHERE group_key = ?",
        (group_key,),
    ).fetchone()

    if existing:
        group_id = int(existing["group_id"])
        conn.execute(
            """
            UPDATE customer_groups
            SET
                group_name = COALESCE(NULLIF(?, ''), group_name),
                address = COALESCE(NULLIF(?, ''), address),
                primary_phone = COALESCE(NULLIF(?, ''), primary_phone),
                updated_at = ?
            WHERE group_id = ?
            """,
            (group_name, address, phone, now_str(), group_id),
        )
        conn.commit()
        return group_id

    cur = conn.execute(
        """
        INSERT INTO customer_groups
        (group_key, group_name, address, primary_phone, source, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (group_key, group_name, address, phone, source, now_str(), now_str()),
    )
    conn.commit()
    return int(cur.lastrowid)


def add_group_member(conn, group_id: int, ma_tt: str):
    conn.execute(
        """
        INSERT OR IGNORE INTO group_members
        (group_id, ma_tt, is_active, created_at)
        VALUES (?, ?, 1, ?)
        """,
        (group_id, ma_tt, now_str()),
    )
    conn.commit()


def insert_debt_lines(conn, df: pd.DataFrame, batch_id: int, billing_period: str):
    records = []
    for _, row in df.iterrows():
        ma_tt = row["ma_tt"]
        customer_name = row.get("customer_name", "")
        address = row.get("address", "")
        phone = row.get("phone", "")
        representative_code = row.get("representative_code", "")

        group_key = group_key_from_text(customer_name, address, representative_code)
        group_name = customer_name if customer_name else ma_tt
        group_id = get_or_create_group(
            conn=conn,
            group_key=group_key,
            group_name=group_name,
            address=address,
            phone=phone,
            source="auto_from_tn08",
        )
        add_group_member(conn, group_id, ma_tt)

        records.append(
            (
                batch_id,
                ma_tt,
                group_id,
                billing_period,
                customer_name,
                address,
                phone,
                representative_code,
                float(row.get("debt_amount", 0)),
                now_str(),
            )
        )

    conn.executemany(
        """
        INSERT INTO debt_snapshot_lines
        (batch_id, ma_tt, group_id, billing_period, customer_name, address, phone,
         representative_code, debt_amount, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        records,
    )
    conn.commit()


def insert_paid_updates(conn, df: pd.DataFrame, batch_id: int):
    records = []
    for _, row in df.iterrows():
        records.append(
            (
                batch_id,
                row["ma_tt"],
                float(row.get("paid_amount", 0)),
                row.get("paid_date", ""),
                now_str(),
            )
        )

    conn.executemany(
        """
        INSERT INTO paid_updates
        (batch_id, ma_tt, paid_amount, paid_date, created_at)
        VALUES (?, ?, ?, ?, ?)
        """,
        records,
    )
    conn.commit()


def add_interaction(conn, group_id: int, result: str, promised_payment_date: str, note: str, created_by: str):
    conn.execute(
        """
        INSERT INTO interactions
        (group_id, result, promised_payment_date, note, created_by, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (group_id, result, promised_payment_date, note, created_by, now_str()),
    )
    conn.commit()


def add_contact(conn, group_id: int, contact_value: str, contact_person: str, role: str, note: str):
    # Các số cũ vẫn được lưu lịch sử, nhưng số mới được đưa thành primary_phone cho nhóm.
    conn.execute(
        """
        INSERT INTO contacts
        (group_id, contact_value, contact_person, role, is_primary, verification_status, note, created_at)
        VALUES (?, ?, ?, ?, 1, 'verified', ?, ?)
        """,
        (group_id, contact_value, contact_person, role, note, now_str()),
    )
    conn.execute(
        """
        UPDATE customer_groups
        SET primary_phone = ?, updated_at = ?
        WHERE group_id = ?
        """,
        (contact_value, now_str(), group_id),
    )
    conn.commit()


def add_sent_message(conn, group_id: int, template_key: str, message_text: str, created_by: str):
    conn.execute(
        """
        INSERT INTO sent_messages
        (group_id, template_key, message_text, created_by, created_at)
        VALUES (?, ?, ?, ?, ?)
        """,
        (group_id, template_key, message_text, created_by, now_str()),
    )
    conn.commit()


# ============================================================
# KHỐI 4: ETL - ĐỌC VÀ CHUẨN HÓA FILE
# ============================================================

def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    """Đọc Excel/CSV, ưu tiên sheet đầu tiên nếu là Excel nhiều sheet."""
    file_name = uploaded_file.name.lower()
    if file_name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded_file)
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding="latin1")

    return pd.read_excel(uploaded_file)


def find_matching_column(df: pd.DataFrame, candidates: list[str]):
    normalized_columns = {normalize_column_key(col): col for col in df.columns}

    for candidate in candidates:
        key = normalize_column_key(candidate)
        if key in normalized_columns:
            return normalized_columns[key]

    return None


def standardize_file(df: pd.DataFrame, file_type: str) -> tuple[pd.DataFrame, list[str]]:
    """
    Chuẩn hóa file theo APP_CONFIG.
    Trả về DataFrame chuẩn và danh sách cảnh báo.
    """
    warnings = []
    df = df.copy()
    df.columns = [normalize_text(c) for c in df.columns]

    cfg = APP_CONFIG["column_candidates"][file_type]
    mapping = {}

    for canonical_col, candidates in cfg.get("required", {}).items():
        source_col = find_matching_column(df, candidates)
        if source_col is None:
            raise ValueError(
                f"Thiếu cột bắt buộc '{canonical_col}'. "
                f"Hãy kiểm tra tên cột trong file hoặc thêm tên cột vào APP_CONFIG."
            )
        mapping[canonical_col] = source_col

    for canonical_col, candidates in cfg.get("optional", {}).items():
        source_col = find_matching_column(df, candidates)
        if source_col is not None:
            mapping[canonical_col] = source_col

    out = pd.DataFrame()
    for canonical_col, source_col in mapping.items():
        out[canonical_col] = df[source_col]

    if "ma_tt" in out.columns:
        out["ma_tt"] = out["ma_tt"].apply(normalize_ma_tt)

    if file_type == "tn08":
        for col in ["customer_name", "address", "representative_code"]:
            if col not in out.columns:
                out[col] = ""
            out[col] = out[col].apply(normalize_text)

        if "phone" not in out.columns:
            out["phone"] = ""
        out["phone"] = out["phone"].apply(normalize_phone)

        if "debt_amount" not in out.columns:
            out["debt_amount"] = 0
        out["debt_amount"] = out["debt_amount"].apply(parse_money)

        out = out[["ma_tt", "customer_name", "address", "phone", "representative_code", "debt_amount"]]

    if file_type == "paid":
        if "paid_amount" not in out.columns:
            out["paid_amount"] = 0
        out["paid_amount"] = out["paid_amount"].apply(parse_money)

        if "paid_date" not in out.columns:
            out["paid_date"] = ""
        out["paid_date"] = out["paid_date"].apply(normalize_text)

        out = out[["ma_tt", "paid_amount", "paid_date"]]

    out = out[out["ma_tt"].astype(str).str.len() > 0].copy()

    duplicated = int(out["ma_tt"].duplicated().sum())
    if duplicated > 0:
        warnings.append(f"Có {duplicated} dòng trùng MA_TT. Hệ thống vẫn import nhưng bạn nên kiểm tra.")

    if file_type == "tn08":
        no_phone = int((out["phone"] == "").sum())
        if no_phone > 0:
            warnings.append(f"Có {no_phone} dòng thiếu số điện thoại.")

    return out, warnings


def import_group_mapping_from_giang_file(conn, uploaded_file) -> tuple[int, int]:
    """
    Import file nhóm kiểu THU CƯỚC GIANG:
    - Mỗi sheet có thể đại diện cho một nhóm khách hàng.
    - Hệ thống quét toàn bộ ô trong từng sheet để tìm MA_TT.
    - Tên sheet được dùng làm group_name.
    - Không phụ thuộc vị trí dòng/cột cụ thể.
    """
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)
    ma_tt_pattern = re.compile(APP_CONFIG["ma_tt_regex"], flags=re.IGNORECASE)

    group_count = 0
    member_count = 0

    skip_sheet_keywords = ["tn08", "cn tn08", "data", "database", "sheet"]

    for sheet_name, df in all_sheets.items():
        sheet_name_clean = normalize_text(sheet_name)

        if normalize_column_key(sheet_name_clean) in [normalize_column_key(x) for x in skip_sheet_keywords]:
            continue

        found_ma_tts = set()

        for value in df.to_numpy().flatten():
            text = normalize_text(value).upper()
            if not text:
                continue
            matches = ma_tt_pattern.findall(text)
            for m in matches:
                found_ma_tts.add(normalize_ma_tt(m))

        if not found_ma_tts:
            continue

        group_key = "MANUAL|" + normalize_column_key(sheet_name_clean)
        group_id = get_or_create_group(
            conn=conn,
            group_key=group_key,
            group_name=sheet_name_clean,
            address="",
            phone="",
            source="manual_from_giang_file",
        )

        group_count += 1
        for ma_tt in found_ma_tts:
            add_group_member(conn, group_id, ma_tt)
            # Cập nhật lại các dòng công nợ mới nhất nếu đã import TN08 trước đó.
            conn.execute(
                """
                UPDATE debt_snapshot_lines
                SET group_id = ?
                WHERE ma_tt = ?
                """,
                (group_id, ma_tt),
            )
            member_count += 1

    conn.commit()
    return group_count, member_count


# ============================================================
# KHỐI 5: BUSINESS LOGIC - TRẠNG THÁI, GOM NHÓM, TIN NHẮN
# ============================================================

def get_group_dashboard(conn) -> pd.DataFrame:
    """
    Dashboard theo nhóm khách hàng/công ty.
    Logic giống file Giang:
    - Một nhóm có nhiều MA_TT.
    - Tổng tiền = SUM tiền nợ của các MA_TT trong batch TN08 mới nhất.
    """
    latest_tn08_batch = get_latest_batch_id(conn, "tn08")
    latest_paid_batch = get_latest_batch_id(conn, "paid")

    if latest_tn08_batch is None:
        return pd.DataFrame()

    paid_filter_sql = ""
    params = [latest_tn08_batch]

    if latest_paid_batch is not None:
        paid_filter_sql = "WHERE batch_id = ?"
        params.append(latest_paid_batch)

    query = f"""
        WITH latest_debt AS (
            SELECT
                d.group_id,
                d.ma_tt,
                d.billing_period,
                d.customer_name,
                d.address,
                d.phone,
                d.debt_amount
            FROM debt_snapshot_lines d
            WHERE d.batch_id = ?
        ),
        paid AS (
            SELECT
                ma_tt,
                SUM(paid_amount) AS paid_amount
            FROM paid_updates
            {paid_filter_sql}
            GROUP BY ma_tt
        ),
        last_interaction AS (
            SELECT i.*
            FROM interactions i
            INNER JOIN (
                SELECT group_id, MAX(created_at) AS max_created_at
                FROM interactions
                GROUP BY group_id
            ) x
            ON i.group_id = x.group_id AND i.created_at = x.max_created_at
        )
        SELECT
            g.group_id,
            g.group_name,
            g.address AS group_address,
            COALESCE(g.primary_phone, '') AS primary_phone,
            MIN(d.billing_period) AS billing_period,
            COUNT(DISTINCT d.ma_tt) AS ma_tt_count,
            SUM(d.debt_amount) AS total_debt,
            SUM(CASE WHEN p.ma_tt IS NOT NULL THEN 1 ELSE 0 END) AS paid_ma_tt_count,
            li.result AS last_result,
            li.promised_payment_date,
            li.note AS last_note,
            li.created_at AS last_contacted_at
        FROM latest_debt d
        LEFT JOIN customer_groups g ON g.group_id = d.group_id
        LEFT JOIN paid p ON p.ma_tt = d.ma_tt
        LEFT JOIN last_interaction li ON li.group_id = d.group_id
        GROUP BY
            g.group_id,
            g.group_name,
            g.address,
            g.primary_phone,
            li.result,
            li.promised_payment_date,
            li.note,
            li.created_at
        ORDER BY total_debt DESC
    """

    df = pd.read_sql_query(query, conn, params=params)

    if df.empty:
        return df

    def calc_status(row):
        if row["ma_tt_count"] > 0 and row["paid_ma_tt_count"] >= row["ma_tt_count"]:
            return "paid"

        if not normalize_text(row.get("primary_phone", "")):
            return "uncontacted"

        last_contacted_at = normalize_text(row.get("last_contacted_at", ""))
        if last_contacted_at:
            return "contacted_unpaid"

        return "uncontacted"

    df["status_code"] = df.apply(calc_status, axis=1)
    df["status_label"] = df["status_code"].apply(lambda x: APP_CONFIG["status"][x]["label"])
    df["status_emoji"] = df["status_code"].apply(lambda x: APP_CONFIG["status"][x]["emoji"])
    df["total_debt_display"] = df["total_debt"].apply(format_money)
    df["display_name"] = df["status_emoji"] + " " + df["group_name"].fillna("Không rõ tên khách")

    return df


def get_group_detail(conn, group_id: int) -> pd.DataFrame:
    latest_tn08_batch = get_latest_batch_id(conn, "tn08")
    latest_paid_batch = get_latest_batch_id(conn, "paid")

    if latest_tn08_batch is None:
        return pd.DataFrame()

    params = [latest_tn08_batch, group_id]
    paid_filter_sql = ""

    if latest_paid_batch is not None:
        paid_filter_sql = "WHERE batch_id = ?"
        params = [latest_tn08_batch, group_id, latest_paid_batch]

    query = f"""
        WITH paid AS (
            SELECT ma_tt, SUM(paid_amount) AS paid_amount
            FROM paid_updates
            {paid_filter_sql}
            GROUP BY ma_tt
        )
        SELECT
            d.ma_tt,
            d.customer_name,
            d.address,
            d.phone,
            d.representative_code,
            d.billing_period,
            d.debt_amount,
            COALESCE(p.paid_amount, 0) AS paid_amount,
            CASE WHEN p.ma_tt IS NOT NULL THEN 'Đã đóng / có trong DS đã đóng'
                 ELSE 'Còn nợ theo TN08'
            END AS line_status
        FROM debt_snapshot_lines d
        LEFT JOIN paid p ON p.ma_tt = d.ma_tt
        WHERE d.batch_id = ? AND d.group_id = ?
        ORDER BY d.debt_amount DESC
    """

    return pd.read_sql_query(query, conn, params=params)


def get_interactions(conn, group_id: int) -> pd.DataFrame:
    return pd.read_sql_query(
        """
        SELECT result, promised_payment_date, note, created_by, created_at
        FROM interactions
        WHERE group_id = ?
        ORDER BY created_at DESC
        LIMIT 20
        """,
        conn,
        params=(group_id,),
    )


def get_contacts(conn, group_id: int) -> pd.DataFrame:
    return pd.read_sql_query(
        """
        SELECT contact_value, contact_person, role, verification_status, note, created_at
        FROM contacts
        WHERE group_id = ?
        ORDER BY created_at DESC
        LIMIT 20
        """,
        conn,
        params=(group_id,),
    )


def render_message(template_key: str, group_row: dict) -> str:
    staff = APP_CONFIG["staff"]
    template = APP_CONFIG["message_templates"][template_key]

    context = {
        "unit_name": staff["unit_name"],
        "website": staff["website"],
        "payment_deadline": staff["payment_deadline"],
        "staff_name": staff["staff_name"],
        "staff_phone": staff["staff_phone"],
        "billing_period": group_row.get("billing_period") or month_default(),
        "group_name": group_row.get("group_name") or "",
        "ma_tt_count": int(group_row.get("ma_tt_count") or 0),
        "total_debt_display": group_row.get("total_debt_display") or format_money(group_row.get("total_debt", 0)),
    }

    return template.format(**context)


def export_report_excel(group_df: pd.DataFrame, detail_df: pd.DataFrame | None = None) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        group_df.to_excel(writer, index=False, sheet_name="Bao cao nhom")
        if detail_df is not None and not detail_df.empty:
            detail_df.to_excel(writer, index=False, sheet_name="Chi tiet MA_TT")
    output.seek(0)
    return output.getvalue()


# ============================================================
# KHỐI 6: STREAMLIT UI
# ============================================================

st.set_page_config(
    page_title="Mini-CRM Thu Cước KHDN",
    page_icon="📞",
    layout="wide",
)

conn = get_conn()
init_db(conn)

st.title(APP_CONFIG["app_title"])
st.caption("Bản single-file: dễ deploy GitHub/Streamlit Cloud, không cần src/, không cần YAML.")

with st.expander("Tư duy thiết kế của app này", expanded=False):
    st.write(
        """
        App này không quản lý theo kiểu 1 MA_TT = 1 khách hàng.
        Logic chính là: 1 công ty/nhóm khách hàng có thể có nhiều MA_TT, giống cách file Giang đang gom CẢNG GÒ DẦU.

        Dữ liệu được xử lý theo 5 lớp trong cùng một file:
        1. Config: mapping cột, mẫu tin, trạng thái.
        2. Database: SQLite lưu batch import, nhóm khách, MA_TT, lịch sử gọi.
        3. ETL: đọc Excel/CSV và chuẩn hóa cột.
        4. Business Logic: gom nhóm, tính tổng tiền, tính trạng thái.
        5. UI: upload, xử lý khách, copy tin nhắn, báo cáo.
        """
    )

tab_upload, tab_work, tab_report, tab_db = st.tabs(
    ["1. Nạp dữ liệu", "2. Bàn làm việc", "3. Báo cáo", "4. Quản trị dữ liệu"]
)


# ------------------------------------------------------------
# TAB 1: NẠP DỮ LIỆU
# ------------------------------------------------------------
with tab_upload:
    st.subheader("1. Nạp dữ liệu")

    billing_period = st.text_input(
        "Kỳ cước cho file TN08",
        value=month_default(),
        help="Ví dụ: 05/2026 hoặc 05. Nên kiểm tra vì kỳ cước thực tế thường là tháng trước.",
    )

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### Upload file TN08 - danh sách nợ")
        file_tn08 = st.file_uploader(
            "Kéo thả file TN08 tại đây",
            type=["xlsx", "xls", "csv"],
            key="tn08_upload",
        )

        if file_tn08 is not None:
            try:
                df_raw = read_uploaded_file(file_tn08)
                df_clean, warnings = standardize_file(df_raw, "tn08")

                st.success(f"Đọc được {len(df_clean)} dòng TN08.")
                st.metric("Tổng tiền nợ nhận diện", format_money(df_clean["debt_amount"].sum()))

                for w in warnings:
                    st.warning(w)

                with st.expander("Xem trước dữ liệu TN08 sau chuẩn hóa"):
                    st.dataframe(df_clean.head(50), use_container_width=True)

                if st.button("Xác nhận nạp TN08 vào hệ thống"):
                    batch_id = create_batch(
                        conn,
                        import_type="tn08",
                        file_name=file_tn08.name,
                        row_count=len(df_clean),
                        total_amount=float(df_clean["debt_amount"].sum()),
                        note=f"Kỳ cước {billing_period}",
                    )
                    insert_debt_lines(conn, df_clean, batch_id, billing_period)
                    st.success(f"Đã nạp TN08 thành công. Batch ID: {batch_id}")

            except Exception as e:
                st.error(f"Lỗi khi nạp TN08: {e}")

    with col2:
        st.markdown("### Upload file khách đã đóng")
        file_paid = st.file_uploader(
            "Kéo thả file đã đóng tại đây",
            type=["xlsx", "xls", "csv"],
            key="paid_upload",
        )

        if file_paid is not None:
            try:
                df_raw_paid = read_uploaded_file(file_paid)
                df_paid, warnings_paid = standardize_file(df_raw_paid, "paid")

                st.success(f"Đọc được {len(df_paid)} dòng đã đóng.")
                for w in warnings_paid:
                    st.warning(w)

                with st.expander("Xem trước dữ liệu đã đóng sau chuẩn hóa"):
                    st.dataframe(df_paid.head(50), use_container_width=True)

                if st.button("Xác nhận nạp danh sách đã đóng"):
                    batch_id = create_batch(
                        conn,
                        import_type="paid",
                        file_name=file_paid.name,
                        row_count=len(df_paid),
                        total_amount=float(df_paid["paid_amount"].sum()),
                        note="Danh sách khách đã đóng",
                    )
                    insert_paid_updates(conn, df_paid, batch_id)
                    st.success(f"Đã nạp DS đã đóng thành công. Batch ID: {batch_id}")

            except Exception as e:
                st.error(f"Lỗi khi nạp file đã đóng: {e}")

    st.divider()

    st.markdown("### Tùy chọn: Upload file nhóm kiểu THU CƯỚC GIANG")
    st.write(
        """
        Nếu bạn có file giống **THU CƯỚC GIANG**, app sẽ quét từng sheet, lấy tên sheet làm tên nhóm
        và tự tìm các MA_TT bên trong sheet. Cách này giúp app gom nhóm giống file Giang hơn.
        """
    )

    group_file = st.file_uploader(
        "Upload file nhóm kiểu Giang",
        type=["xlsx", "xls"],
        key="group_mapping_upload",
    )

    if group_file is not None:
        try:
            if st.button("Import nhóm từ file Giang"):
                group_count, member_count = import_group_mapping_from_giang_file(conn, group_file)
                st.success(f"Đã import {group_count} nhóm và {member_count} MA_TT vào mapping nhóm.")
        except Exception as e:
            st.error(f"Lỗi khi import nhóm từ file Giang: {e}")


# ------------------------------------------------------------
# TAB 2: BÀN LÀM VIỆC
# ------------------------------------------------------------
with tab_work:
    st.subheader("2. Bàn làm việc theo nhóm khách hàng")

    group_df = get_group_dashboard(conn)

    if group_df.empty:
        st.info("Chưa có dữ liệu TN08. Hãy vào tab 1 để nạp file TN08 trước.")
    else:
        total_groups = len(group_df)
        total_debt = group_df[group_df["status_code"] != "paid"]["total_debt"].sum()
        red_count = int((group_df["status_code"] == "uncontacted").sum())
        yellow_count = int((group_df["status_code"] == "contacted_unpaid").sum())
        green_count = int((group_df["status_code"] == "paid").sum())

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Số nhóm/công ty", total_groups)
        m2.metric("Tổng tiền còn cần xử lý", format_money(total_debt))
        m3.metric("Chưa liên hệ / thiếu số", red_count)
        m4.metric("Đã liên hệ chưa đóng", yellow_count)

        status_labels = ["Tất cả"] + sorted(group_df["status_label"].dropna().unique().tolist())
        selected_status = st.selectbox("Lọc trạng thái", status_labels)

        search_text = st.text_input("Tìm theo tên công ty / số điện thoại")

        view_df = group_df.copy()

        if selected_status != "Tất cả":
            view_df = view_df[view_df["status_label"] == selected_status]

        if search_text:
            s = search_text.lower()
            view_df = view_df[
                view_df["group_name"].fillna("").str.lower().str.contains(s)
                | view_df["primary_phone"].fillna("").str.lower().str.contains(s)
            ]

        show_cols = [
            "status_emoji",
            "status_label",
            "group_name",
            "primary_phone",
            "ma_tt_count",
            "total_debt_display",
            "billing_period",
            "last_result",
            "promised_payment_date",
            "last_contacted_at",
        ]

        st.dataframe(
            view_df[show_cols],
            use_container_width=True,
            height=360,
        )

        if not view_df.empty:
            group_options = {
                f"{row['status_emoji']} {row['group_name']} | {row['total_debt_display']} | {row['ma_tt_count']} mã": int(row["group_id"])
                for _, row in view_df.iterrows()
            }

            selected_label = st.selectbox("Chọn công ty/nhóm để xử lý", list(group_options.keys()))
            selected_group_id = group_options[selected_label]

            selected_row = group_df[group_df["group_id"] == selected_group_id].iloc[0].to_dict()

            st.divider()
            st.markdown("### Chi tiết nhóm khách hàng")

            c_info, c_action = st.columns([1.1, 1])

            with c_info:
                st.write(f"**Tên nhóm/công ty:** {selected_row.get('group_name', '')}")
                st.write(f"**Trạng thái:** {selected_row.get('status_emoji', '')} {selected_row.get('status_label', '')}")
                st.write(f"**Số điện thoại/Zalo ưu tiên:** {selected_row.get('primary_phone', '') or 'Chưa có'}")
                st.write(f"**Kỳ cước:** {selected_row.get('billing_period', '')}")
                st.write(f"**Số lượng MA_TT:** {int(selected_row.get('ma_tt_count') or 0)}")
                st.write(f"**Tổng tiền:** {selected_row.get('total_debt_display', '')}")

                detail_df = get_group_detail(conn, selected_group_id)
                st.markdown("#### Danh sách MA_TT trong nhóm")
                st.dataframe(
                    detail_df[
                        [
                            "ma_tt",
                            "customer_name",
                            "phone",
                            "billing_period",
                            "debt_amount",
                            "paid_amount",
                            "line_status",
                        ]
                    ],
                    use_container_width=True,
                    height=260,
                )

                st.markdown("#### Lịch sử gọi")
                history_df = get_interactions(conn, selected_group_id)
                st.dataframe(history_df, use_container_width=True, height=180)

                st.markdown("#### Lịch sử số liên hệ")
                contacts_df = get_contacts(conn, selected_group_id)
                st.dataframe(contacts_df, use_container_width=True, height=160)

            with c_action:
                st.markdown("#### Cập nhật kết quả cuộc gọi")

                with st.form("interaction_form", clear_on_submit=True):
                    result = st.selectbox("Kết quả cuộc gọi", APP_CONFIG["contact_results"])
                    promised_payment_date = st.text_input("Ngày hẹn thanh toán", placeholder="Ví dụ: 15/06/2026")
                    note = st.text_area("Ghi chú", placeholder="Ví dụ: gặp chị Lan kế toán, hẹn thứ 6 chuyển khoản...")
                    created_by = st.text_input("Người cập nhật", value=APP_CONFIG["staff"]["staff_name"])
                    submitted = st.form_submit_button("Lưu lịch sử gọi")

                    if submitted:
                        add_interaction(
                            conn,
                            group_id=selected_group_id,
                            result=result,
                            promised_payment_date=promised_payment_date,
                            note=note,
                            created_by=created_by,
                        )
                        st.success("Đã lưu lịch sử gọi.")
                        st.rerun()

                st.markdown("#### Cập nhật số Zalo/SĐT mới")

                with st.form("contact_form", clear_on_submit=True):
                    new_contact = st.text_input("Số Zalo/SĐT mới")
                    contact_person = st.text_input("Tên người phụ trách", placeholder="Ví dụ: Chị Lan")
                    role = st.text_input("Vai trò", value="Kế toán")
                    contact_note = st.text_area("Ghi chú số liên hệ")
                    contact_submit = st.form_submit_button("Lưu số liên hệ mới")

                    if contact_submit:
                        phone = normalize_phone(new_contact)
                        if not phone:
                            st.error("Số điện thoại/Zalo chưa hợp lệ.")
                        else:
                            add_contact(
                                conn,
                                group_id=selected_group_id,
                                contact_value=phone,
                                contact_person=contact_person,
                                role=role,
                                note=contact_note,
                            )
                            st.success("Đã lưu số liên hệ mới.")
                            st.rerun()

            st.divider()
            st.markdown("### Mẫu tin nhắn Zalo để copy")

            template_key = st.selectbox(
                "Chọn mẫu tin",
                options=list(APP_CONFIG["message_templates"].keys()),
                format_func=lambda x: "Mẫu chung an toàn" if x == "general" else "Mẫu chi tiết sau khi xác minh",
            )

            message_text = render_message(template_key, selected_row)
            st.code(message_text, language="text")

            if st.button("Ghi nhận đã copy/gửi mẫu tin này"):
                add_sent_message(
                    conn,
                    group_id=selected_group_id,
                    template_key=template_key,
                    message_text=message_text,
                    created_by=APP_CONFIG["staff"]["staff_name"],
                )
                st.success("Đã ghi nhận lịch sử mẫu tin.")


# ------------------------------------------------------------
# TAB 3: BÁO CÁO
# ------------------------------------------------------------
with tab_report:
    st.subheader("3. Báo cáo KPI")

    group_df = get_group_dashboard(conn)

    if group_df.empty:
        st.info("Chưa có dữ liệu để báo cáo.")
    else:
        unpaid_df = group_df[group_df["status_code"] != "paid"].copy()

        total_debt = unpaid_df["total_debt"].sum()
        paid_groups = int((group_df["status_code"] == "paid").sum())
        contacted_groups = int((group_df["status_code"] == "contacted_unpaid").sum())
        uncontacted_groups = int((group_df["status_code"] == "uncontacted").sum())

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Tổng nhóm/công ty", len(group_df))
        k2.metric("Tổng tiền còn cần xử lý", format_money(total_debt))
        k3.metric("Nhóm đã đóng", paid_groups)
        k4.metric("Nhóm chưa liên hệ / thiếu số", uncontacted_groups)

        st.markdown("### Báo cáo theo nhóm")
        report_cols = [
            "status_emoji",
            "status_label",
            "group_name",
            "primary_phone",
            "ma_tt_count",
            "total_debt",
            "total_debt_display",
            "billing_period",
            "last_result",
            "promised_payment_date",
            "last_contacted_at",
        ]
        st.dataframe(group_df[report_cols], use_container_width=True, height=420)

        excel_bytes = export_report_excel(group_df[report_cols])
        st.download_button(
            label="Tải báo cáo Excel",
            data=excel_bytes,
            file_name="bao_cao_thu_cuoc_khdn.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ------------------------------------------------------------
# TAB 4: QUẢN TRỊ DỮ LIỆU
# ------------------------------------------------------------
with tab_db:
    st.subheader("4. Quản trị dữ liệu")

    st.warning(
        "Khu vực này dùng để kiểm tra dữ liệu import và reset database khi cần. "
        "Chỉ dùng reset khi bạn muốn làm lại từ đầu."
    )

    st.markdown("### Lịch sử import")
    batches_df = pd.read_sql_query(
        """
        SELECT batch_id, import_type, file_name, row_count, total_amount, note, created_at
        FROM import_batches
        ORDER BY batch_id DESC
        LIMIT 50
        """,
        conn,
    )
    st.dataframe(batches_df, use_container_width=True)

    st.markdown("### Thống kê bảng dữ liệu")
    table_names = [
        "customer_groups",
        "group_members",
        "debt_snapshot_lines",
        "paid_updates",
        "interactions",
        "contacts",
        "sent_messages",
    ]

    stats = []
    for table in table_names:
        try:
            count = conn.execute(f"SELECT COUNT(*) AS c FROM {table}").fetchone()["c"]
            stats.append({"table": table, "rows": count})
        except Exception:
            stats.append({"table": table, "rows": "error"})

    st.dataframe(pd.DataFrame(stats), use_container_width=True)

    with st.expander("Reset database"):
        st.write("Dùng khi bạn muốn xóa toàn bộ dữ liệu import/lịch sử để làm lại từ đầu.")
        confirm_reset = st.text_input("Gõ RESET để xác nhận")
        if st.button("Xóa toàn bộ dữ liệu") and confirm_reset == "RESET":
            conn.executescript(
                """
                DELETE FROM sent_messages;
                DELETE FROM contacts;
                DELETE FROM interactions;
                DELETE FROM paid_updates;
                DELETE FROM debt_snapshot_lines;
                DELETE FROM group_members;
                DELETE FROM customer_groups;
                DELETE FROM import_batches;
                """
            )
            conn.commit()
            st.success("Đã reset database.")
            st.rerun()
