# ============================================================
# APP.PY - MINI CRM THU CƯỚC KHDN VNPT
# Phiên bản single-file, dễ deploy Streamlit Cloud
#
# QUY TRÌNH NGHIỆP VỤ ĐÚNG:
# 1. Upload DS giao kỳ cước
# 2. Chọn nhân viên, mặc định: Vương Thanh Thuận
# 3. Upload TN08 hóa đơn chưa thu
# 4. App lọc DS giao theo nhân viên
# 5. App dò MA_TT từ DS giao sang TN08 để lấy Total Nợ thu vét
# 6. App hiển thị tiền cần thu, số điện thoại, mẫu tin nhắn Zalo
#
# LƯU Ý KIẾN TRÚC:
# - Không dùng src/
# - Không dùng yaml
# - Không dùng hardcode rải rác
# - Tất cả cấu hình dễ thay đổi nằm trong APP_CONFIG
# ============================================================

import io
import re
import sqlite3
import unicodedata
from datetime import datetime
from decimal import Decimal, InvalidOperation

import pandas as pd
import streamlit as st


# ============================================================
# KHỐI 1: CONFIGURATION LAYER - ZERO HARDCODING
# ============================================================

APP_CONFIG = {
    "app_title": "📞 Mini-CRM Thu Cước KHDN - VNPT",
    "db_path": "crm_vnpt_thu_cuoc.db",

    "default_staff_filter": "Vương Thanh Thuận",

    "message": {
        "unit_name": "VNPT Long Thành- Nhơn Trạch",
        "website": "https://vnptdongnai.vn/",
        "payment_deadline": "ngày 15",
        "staff_name": "Thuận",
        "staff_phone": "0837892579",
        "default_billing_period": "05",
        "template": """{unit_name} thông báo:
Đã có thông báo cước kỳ cước tháng {billing_period}. Anh/chị truy cập trang {website} để lấy thông báo cước và hóa đơn của công ty.
Anh/chị vui lòng thanh toán cước trước {payment_deadline}. Sau ngày 15 hệ thống sẽ tự động đưa lưới khóa khi ghi nhận còn tồn nợ.
Liên hệ nhân viên kinh doanh: {staff_name}: {staff_phone} để được hỗ trợ gạch nợ sớm nhất sau khi thanh toán.
Nếu đã thanh toán cước, anh/chị vui lòng bỏ qua tin nhắn trên.
Trân trọng."""
    },

    "status": {
        "paid": {
            "label": "Đã đóng",
            "emoji": "🟢",
        },
        "contacted_unpaid": {
            "label": "Đã liên hệ chưa đóng",
            "emoji": "🟡",
        },
        "uncontacted": {
            "label": "Chưa liên hệ",
            "emoji": "🔴",
        },
        "no_phone": {
            "label": "Thiếu số điện thoại",
            "emoji": "🔴",
        },
        "need_check": {
            "label": "Cần kiểm tra dữ liệu",
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

    # Cấu hình mapping cột DS giao kỳ cước.
    # Sheet DS giao dùng để xác định tuyến của nhân viên và số điện thoại gọi khách.
    "assignment_columns": {
        "staff_name": [
            "nhân viên", "nhan vien", "nv thu", "người phụ trách", "nguoi phu trach",
            "tên nhân viên", "ten nhan vien"
        ],
        "ma_tt": [
            "mã thanh toán", "ma thanh toan", "mã tt", "ma_tt", "ma tt",
            "mã thanh toán", "mã thanh toán"
        ],
        "customer_name": [
            "tên thanh toán", "ten thanh toan", "tên khách hàng", "ten khach hang",
            "khách hàng", "khach hang", "ten_tt", "tên tt"
        ],
        "phone": [
            "số dt liên hệ", "so dt lien he", "số điện thoại", "so dien thoai",
            "điện thoại", "dien thoai", "sdt", "phone"
        ],
        "address": [
            "địa chỉ thanh toán", "dia chi thanh toan", "địa chỉ", "dia chi",
            "địa chỉ kh", "dia chi kh"
        ],
        # Chỉ để tham khảo, không dùng làm tiền cần thu.
        "generated_amount": [
            "tiền phát sinh", "tien phat sinh", "phát sinh", "phat sinh",
            "số tiền", "so tien"
        ],
    },

    # Cấu hình mapping cột TN08.
    # TN08 dùng để lấy số tiền cần thu thực tế.
    "tn08_columns": {
        "ma_tt": [
            "ma_tt", "mã tt", "ma tt", "mã thanh toán", "ma thanh toan",
            "mã thanh toán", "mã thanh toán"
        ],
        "debt_amount": [
            "total nợ thu vét", "total no thu vet", "total nợ thu vét",
            "nợ thu vét", "no thu vet", "tiền nợ", "tien no",
            "nợ còn lại", "no con lai", "tổng nợ", "tong no"
        ],
        "customer_name_tn08": [
            "tên khách hàng", "ten khach hang", "tên khách hàng",
            "tên thanh toán", "ten thanh toan", "ten_tt", "tên tt",
            "khách hàng", "khach hang"
        ],
        "address_tn08": [
            "địa chỉ kh", "dia chi kh", "địa chỉ", "dia chi",
            "địa chỉ thanh toán", "dia chi thanh toan"
        ],
    },

    # Cấu hình mapping file đã đóng.
    "paid_columns": {
        "ma_tt": [
            "ma_tt", "mã tt", "ma tt", "mã thanh toán", "ma thanh toan",
            "mã thanh toán", "mã thanh toán"
        ],
        "paid_amount": [
            "số tiền", "so tien", "tiền đóng", "tien dong",
            "số tiền đã đóng", "so tien da dong", "amount"
        ],
        "paid_date": [
            "ngày đóng", "ngay dong", "ngày thanh toán", "ngay thanh toan",
            "payment date"
        ],
    },
}


# ============================================================
# KHỐI 2: UTILITY LAYER - CHUẨN HÓA DỮ LIỆU
# ============================================================

def normalize_text(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def remove_accents(value) -> str:
    text = normalize_text(value).lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return text


def normalize_key(value) -> str:
    text = remove_accents(value)
    text = re.sub(r"[^a-z0-9]+", "_", text).strip("_")
    return text


def normalize_ma_tt(value) -> str:
    text = normalize_text(value)
    if text.endswith(".0"):
        text = text[:-2]
    return text.upper()


def normalize_phone(value) -> str:
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
    if value is None or value == "":
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    text = normalize_text(value)
    text = text.replace("đ", "").replace("Đ", "").replace("VND", "").replace("vnd", "")
    text = text.replace(" ", "")

    # VN format: 1.234.567,89
    if "." in text and "," in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        # Thường dữ liệu tiền không cần phần thập phân, bỏ dấu phẩy nghìn.
        text = text.replace(",", "")

    try:
        return float(Decimal(text))
    except (InvalidOperation, ValueError):
        return 0.0


def format_money(value) -> str:
    try:
        return f"{float(value):,.0f} đồng".replace(",", ".")
    except Exception:
        return "0 đồng"


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def safe_lower(value) -> str:
    return remove_accents(value)


def build_staff_key(value) -> str:
    return normalize_key(value)


# ============================================================
# KHỐI 3: DATABASE LAYER - SQLITE
# ============================================================

def get_conn():
    conn = sqlite3.connect(APP_CONFIG["db_path"], check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(conn):
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS import_batches (
            batch_id INTEGER PRIMARY KEY AUTOINCREMENT,
            import_type TEXT NOT NULL,
            file_name TEXT,
            staff_filter TEXT,
            row_count INTEGER DEFAULT 0,
            total_amount REAL DEFAULT 0,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS work_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_id INTEGER NOT NULL,
            staff_name TEXT,
            ma_tt TEXT NOT NULL,
            customer_name TEXT,
            phone TEXT,
            address TEXT,
            generated_amount REAL DEFAULT 0,
            debt_amount REAL DEFAULT 0,
            tn08_customer_name TEXT,
            tn08_address TEXT,
            data_status TEXT,
            created_at TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_work_items_batch_ma_tt
        ON work_items(batch_id, ma_tt);

        CREATE TABLE IF NOT EXISTS paid_updates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_id INTEGER NOT NULL,
            ma_tt TEXT NOT NULL,
            paid_amount REAL DEFAULT 0,
            paid_date TEXT,
            created_at TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_paid_updates_ma_tt
        ON paid_updates(ma_tt);

        CREATE TABLE IF NOT EXISTS interactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ma_tt TEXT NOT NULL,
            result TEXT NOT NULL,
            promised_payment_date TEXT,
            note TEXT,
            created_by TEXT,
            created_at TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_interactions_ma_tt
        ON interactions(ma_tt);

        CREATE TABLE IF NOT EXISTS contacts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ma_tt TEXT NOT NULL,
            contact_value TEXT NOT NULL,
            contact_person TEXT,
            role TEXT,
            note TEXT,
            created_at TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_contacts_ma_tt
        ON contacts(ma_tt);

        CREATE TABLE IF NOT EXISTS sent_messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ma_tt TEXT NOT NULL,
            message_text TEXT NOT NULL,
            created_by TEXT,
            created_at TEXT NOT NULL
        );
    """)
    conn.commit()


def create_batch(conn, import_type: str, file_name: str, staff_filter: str, row_count: int, total_amount: float) -> int:
    cur = conn.execute(
        """
        INSERT INTO import_batches
        (import_type, file_name, staff_filter, row_count, total_amount, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (import_type, file_name, staff_filter, row_count, total_amount, now_str())
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
        (import_type,)
    ).fetchone()
    return int(row["batch_id"]) if row else None


def insert_work_items(conn, batch_id: int, df: pd.DataFrame):
    records = []
    for _, row in df.iterrows():
        records.append((
            batch_id,
            row.get("staff_name", ""),
            row.get("ma_tt", ""),
            row.get("customer_name", ""),
            row.get("phone", ""),
            row.get("address", ""),
            float(row.get("generated_amount", 0) or 0),
            float(row.get("debt_amount", 0) or 0),
            row.get("customer_name_tn08", ""),
            row.get("address_tn08", ""),
            row.get("data_status", ""),
            now_str(),
        ))

    conn.executemany(
        """
        INSERT INTO work_items
        (batch_id, staff_name, ma_tt, customer_name, phone, address, generated_amount,
         debt_amount, tn08_customer_name, tn08_address, data_status, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        records
    )
    conn.commit()


def insert_paid_updates(conn, batch_id: int, df: pd.DataFrame):
    records = []
    for _, row in df.iterrows():
        records.append((
            batch_id,
            row.get("ma_tt", ""),
            float(row.get("paid_amount", 0) or 0),
            row.get("paid_date", ""),
            now_str(),
        ))

    conn.executemany(
        """
        INSERT INTO paid_updates
        (batch_id, ma_tt, paid_amount, paid_date, created_at)
        VALUES (?, ?, ?, ?, ?)
        """,
        records
    )
    conn.commit()


def add_interaction(conn, ma_tt: str, result: str, promised_payment_date: str, note: str, created_by: str):
    conn.execute(
        """
        INSERT INTO interactions
        (ma_tt, result, promised_payment_date, note, created_by, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (ma_tt, result, promised_payment_date, note, created_by, now_str())
    )
    conn.commit()


def add_contact(conn, ma_tt: str, contact_value: str, contact_person: str, role: str, note: str):
    conn.execute(
        """
        INSERT INTO contacts
        (ma_tt, contact_value, contact_person, role, note, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (ma_tt, contact_value, contact_person, role, note, now_str())
    )
    conn.commit()


def add_sent_message(conn, ma_tt: str, message_text: str, created_by: str):
    conn.execute(
        """
        INSERT INTO sent_messages
        (ma_tt, message_text, created_by, created_at)
        VALUES (?, ?, ?, ?)
        """,
        (ma_tt, message_text, created_by, now_str())
    )
    conn.commit()


# ============================================================
# KHỐI 4: ETL LAYER - ĐỌC FILE, MAP CỘT, MERGE NHƯ VLOOKUP
# ============================================================

def read_excel_or_csv(uploaded_file, sheet_name=None):
    file_name = uploaded_file.name.lower()
    if file_name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded_file)
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding="latin1")

    return pd.read_excel(uploaded_file, sheet_name=sheet_name)


def list_excel_sheets(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file)
        return xls.sheet_names
    finally:
        uploaded_file.seek(0)


def find_col(df: pd.DataFrame, candidates: list[str]):
    normalized_cols = {normalize_key(col): col for col in df.columns}
    for c in candidates:
        key = normalize_key(c)
        if key in normalized_cols:
            return normalized_cols[key]
    return None


def standardize_assignment(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()
    df.columns = [normalize_text(c) for c in df.columns]

    out = pd.DataFrame()

    required_fields = ["staff_name", "ma_tt"]
    optional_fields = ["customer_name", "phone", "address", "generated_amount"]

    for field in required_fields:
        col = find_col(df, APP_CONFIG["assignment_columns"][field])
        if col is None:
            raise ValueError(f"File DS giao thiếu cột bắt buộc: {field}. Hãy kiểm tra tên cột.")
        out[field] = df[col]

    for field in optional_fields:
        col = find_col(df, APP_CONFIG["assignment_columns"][field])
        if col is not None:
            out[field] = df[col]
        else:
            out[field] = ""

    out["staff_name"] = out["staff_name"].apply(normalize_text)
    out["staff_key"] = out["staff_name"].apply(build_staff_key)
    out["ma_tt"] = out["ma_tt"].apply(normalize_ma_tt)
    out["customer_name"] = out["customer_name"].apply(normalize_text)
    out["phone"] = out["phone"].apply(normalize_phone)
    out["address"] = out["address"].apply(normalize_text)
    out["generated_amount"] = out["generated_amount"].apply(parse_money)

    out = out[out["ma_tt"].astype(str).str.len() > 0].copy()
    return out


def standardize_tn08(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()
    df.columns = [normalize_text(c) for c in df.columns]

    out = pd.DataFrame()

    # Bắt buộc có MA_TT và debt_amount.
    for field in ["ma_tt", "debt_amount"]:
        col = find_col(df, APP_CONFIG["tn08_columns"][field])
        if col is None:
            raise ValueError(
                f"File TN08 thiếu cột bắt buộc: {field}. "
                "Cần có MA_TT và Total Nợ thu vét."
            )
        out[field] = df[col]

    # Optional.
    for field in ["customer_name_tn08", "address_tn08"]:
        col = find_col(df, APP_CONFIG["tn08_columns"][field])
        if col is not None:
            out[field] = df[col]
        else:
            out[field] = ""

    out["ma_tt"] = out["ma_tt"].apply(normalize_ma_tt)
    out["debt_amount"] = out["debt_amount"].apply(parse_money)
    out["customer_name_tn08"] = out["customer_name_tn08"].apply(normalize_text)
    out["address_tn08"] = out["address_tn08"].apply(normalize_text)

    out = out[out["ma_tt"].astype(str).str.len() > 0].copy()

    # Nếu TN08 có trùng MA_TT, cộng tiền theo MA_TT để tránh nhân dòng khi merge.
    out = (
        out.groupby("ma_tt", as_index=False)
        .agg({
            "debt_amount": "sum",
            "customer_name_tn08": "first",
            "address_tn08": "first",
        })
    )

    return out


def standardize_paid(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()
    df.columns = [normalize_text(c) for c in df.columns]

    out = pd.DataFrame()

    col_ma = find_col(df, APP_CONFIG["paid_columns"]["ma_tt"])
    if col_ma is None:
        raise ValueError("File đã đóng thiếu cột MA_TT / Mã thanh toán.")
    out["ma_tt"] = df[col_ma]

    col_paid_amount = find_col(df, APP_CONFIG["paid_columns"]["paid_amount"])
    out["paid_amount"] = df[col_paid_amount] if col_paid_amount else 0

    col_paid_date = find_col(df, APP_CONFIG["paid_columns"]["paid_date"])
    out["paid_date"] = df[col_paid_date] if col_paid_date else ""

    out["ma_tt"] = out["ma_tt"].apply(normalize_ma_tt)
    out["paid_amount"] = out["paid_amount"].apply(parse_money)
    out["paid_date"] = out["paid_date"].apply(normalize_text)

    out = out[out["ma_tt"].astype(str).str.len() > 0].copy()
    return out


def create_working_dataset(df_assignment: pd.DataFrame, df_tn08: pd.DataFrame, selected_staff: str) -> pd.DataFrame:
    """
    Đây là phần thay thế công thức Excel:
    =+VLOOKUP(F46,'cn TN08'!$B:$I,8,0)

    Logic:
    - F46 trong sheet Giao = ma_tt
    - cn TN08 cột B:I = df_tn08
    - cột 8 trong B:I = Total Nợ thu vét = debt_amount
    """
    selected_key = build_staff_key(selected_staff)

    df_staff = df_assignment[df_assignment["staff_key"] == selected_key].copy()

    if df_staff.empty:
        return pd.DataFrame()

    merged = df_staff.merge(
        df_tn08[["ma_tt", "debt_amount", "customer_name_tn08", "address_tn08"]],
        on="ma_tt",
        how="left"
    )

    merged["debt_amount"] = merged["debt_amount"].fillna(0)
    merged["customer_name_tn08"] = merged["customer_name_tn08"].fillna("")
    merged["address_tn08"] = merged["address_tn08"].fillna("")

    def data_status(row):
        if row["debt_amount"] <= 0:
            return "Không tìm thấy nợ trong TN08 / đã hết nợ / cần kiểm tra"
        return "Có nợ theo TN08"

    merged["data_status"] = merged.apply(data_status, axis=1)
    merged["debt_amount_display"] = merged["debt_amount"].apply(format_money)
    merged["generated_amount_display"] = merged["generated_amount"].apply(format_money)

    return merged


# ============================================================
# KHỐI 5: BUSINESS LOGIC LAYER
# ============================================================

def get_current_work_items(conn) -> pd.DataFrame:
    latest_batch = get_latest_batch_id(conn, "working_dataset")
    if latest_batch is None:
        return pd.DataFrame()

    latest_paid_batch = get_latest_batch_id(conn, "paid")

    if latest_paid_batch is None:
        paid_cte = """
            SELECT '' AS ma_tt, 0 AS paid_amount
            WHERE 1 = 0
        """
        params = [latest_batch]
    else:
        paid_cte = """
            SELECT ma_tt, SUM(paid_amount) AS paid_amount
            FROM paid_updates
            WHERE batch_id = ?
            GROUP BY ma_tt
        """
        params = [latest_paid_batch, latest_batch]

    query = f"""
        WITH paid AS (
            {paid_cte}
        ),
        last_interaction AS (
            SELECT i.*
            FROM interactions i
            INNER JOIN (
                SELECT ma_tt, MAX(created_at) AS max_created_at
                FROM interactions
                GROUP BY ma_tt
            ) x ON i.ma_tt = x.ma_tt AND i.created_at = x.max_created_at
        ),
        latest_contact AS (
            SELECT c.*
            FROM contacts c
            INNER JOIN (
                SELECT ma_tt, MAX(created_at) AS max_created_at
                FROM contacts
                GROUP BY ma_tt
            ) x ON c.ma_tt = x.ma_tt AND c.created_at = x.max_created_at
        )
        SELECT
            w.id,
            w.batch_id,
            w.staff_name,
            w.ma_tt,
            w.customer_name,
            COALESCE(NULLIF(c.contact_value, ''), w.phone) AS phone,
            w.address,
            w.generated_amount,
            w.debt_amount,
            w.tn08_customer_name,
            w.tn08_address,
            w.data_status,
            COALESCE(p.paid_amount, 0) AS paid_amount,
            li.result AS last_result,
            li.promised_payment_date,
            li.note AS last_note,
            li.created_by AS last_created_by,
            li.created_at AS last_contacted_at
        FROM work_items w
        LEFT JOIN paid p ON p.ma_tt = w.ma_tt
        LEFT JOIN last_interaction li ON li.ma_tt = w.ma_tt
        LEFT JOIN latest_contact c ON c.ma_tt = w.ma_tt
        WHERE w.batch_id = ?
        ORDER BY w.debt_amount DESC
    """

    df = pd.read_sql_query(query, conn, params=params)

    if df.empty:
        return df

    def calc_status(row):
        if float(row.get("paid_amount", 0) or 0) > 0:
            return "paid"

        if float(row.get("debt_amount", 0) or 0) <= 0:
            return "need_check"

        if not normalize_text(row.get("phone", "")):
            return "no_phone"

        if normalize_text(row.get("last_contacted_at", "")):
            return "contacted_unpaid"

        return "uncontacted"

    df["status_code"] = df.apply(calc_status, axis=1)
    df["status_label"] = df["status_code"].apply(lambda x: APP_CONFIG["status"][x]["label"])
    df["status_emoji"] = df["status_code"].apply(lambda x: APP_CONFIG["status"][x]["emoji"])
    df["debt_amount_display"] = df["debt_amount"].apply(format_money)
    df["generated_amount_display"] = df["generated_amount"].apply(format_money)
    df["paid_amount_display"] = df["paid_amount"].apply(format_money)
    df["customer_display"] = df["customer_name"].where(
        df["customer_name"].astype(str).str.len() > 0,
        df["tn08_customer_name"]
    )

    return df


def get_interactions(conn, ma_tt: str) -> pd.DataFrame:
    return pd.read_sql_query(
        """
        SELECT result, promised_payment_date, note, created_by, created_at
        FROM interactions
        WHERE ma_tt = ?
        ORDER BY created_at DESC
        LIMIT 20
        """,
        conn,
        params=(ma_tt,)
    )


def get_contacts(conn, ma_tt: str) -> pd.DataFrame:
    return pd.read_sql_query(
        """
        SELECT contact_value, contact_person, role, note, created_at
        FROM contacts
        WHERE ma_tt = ?
        ORDER BY created_at DESC
        LIMIT 20
        """,
        conn,
        params=(ma_tt,)
    )


def render_message(billing_period: str) -> str:
    cfg = APP_CONFIG["message"]
    return cfg["template"].format(
        unit_name=cfg["unit_name"],
        billing_period=billing_period,
        website=cfg["website"],
        payment_deadline=cfg["payment_deadline"],
        staff_name=cfg["staff_name"],
        staff_phone=cfg["staff_phone"],
    )


def summary_metrics(df: pd.DataFrame):
    if df.empty:
        return {
            "total_rows": 0,
            "total_debt": 0.0,
            "paid_count": 0,
            "contacted_count": 0,
            "uncontacted_count": 0,
            "need_check_count": 0,
        }

    return {
        "total_rows": len(df),
        "total_debt": float(df[df["status_code"] != "paid"]["debt_amount"].sum()),
        "paid_count": int((df["status_code"] == "paid").sum()),
        "contacted_count": int((df["status_code"] == "contacted_unpaid").sum()),
        "uncontacted_count": int((df["status_code"].isin(["uncontacted", "no_phone"])).sum()),
        "need_check_count": int((df["status_code"] == "need_check").sum()),
    }


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    export_cols = [
        "status_emoji", "status_label", "staff_name", "ma_tt",
        "customer_display", "phone", "address",
        "debt_amount", "debt_amount_display",
        "generated_amount", "generated_amount_display",
        "paid_amount", "paid_amount_display",
        "data_status", "last_result", "promised_payment_date",
        "last_note", "last_contacted_at"
    ]
    export_df = df[[c for c in export_cols if c in df.columns]].copy()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Bao cao thu cuoc")
    output.seek(0)
    return output.getvalue()


# ============================================================
# KHỐI 6: STREAMLIT UI LAYER
# ============================================================

st.set_page_config(
    page_title="Mini-CRM Thu Cước KHDN",
    page_icon="📞",
    layout="wide",
)

conn = get_conn()
init_db(conn)

st.title(APP_CONFIG["app_title"])
st.caption(
    "Quy trình đúng: DS giao kỳ cước → lọc nhân viên → dò TN08 bằng MA_TT → lấy Total Nợ thu vét → gọi/Zalo khách."
)

with st.expander("App này thay thế thao tác Excel nào?", expanded=False):
    st.write(
        """
        App này thay thế thao tác thủ công:

        1. Mở file DS giao kỳ cước.
        2. Lọc nhân viên **Vương Thanh Thuận**.
        3. Copy các dòng thuộc tuyến của bạn.
        4. Copy TN08 vào sheet `cn TN08`.
        5. Dùng công thức `=VLOOKUP(F46,'cn TN08'!$B:$I,8,0)`.
        6. Lấy tiền ở `Total Nợ thu vét` để gọi/Zalo khách.

        App không dùng cột **Tiền phát sinh** làm tiền cần thu. 
        Cột đó chỉ để tham khảo/đối chiếu.
        """
    )


tab_import, tab_work, tab_report, tab_admin = st.tabs(
    ["1. Nạp dữ liệu", "2. Bàn làm việc", "3. Báo cáo", "4. Quản trị"]
)


# ------------------------------------------------------------
# TAB 1: NẠP DỮ LIỆU
# ------------------------------------------------------------
with tab_import:
    st.subheader("1. Nạp dữ liệu theo đúng quy trình thu cước")

    st.markdown("### Bước A - Upload DS giao kỳ cước")
    assignment_file = st.file_uploader(
        "Upload file DS giao kỳ cước",
        type=["xlsx", "xls", "csv"],
        key="assignment_file"
    )

    df_assignment_std = None
    selected_staff = None

    if assignment_file is not None:
        try:
            # Với file Excel nhiều sheet, cho chọn sheet.
            if assignment_file.name.lower().endswith((".xlsx", ".xls")):
                sheets = list_excel_sheets(assignment_file)
                assignment_sheet = st.selectbox("Chọn sheet chứa DS giao kỳ cước", sheets, key="assignment_sheet")
                df_assignment_raw = read_excel_or_csv(assignment_file, sheet_name=assignment_sheet)
            else:
                df_assignment_raw = read_excel_or_csv(assignment_file)

            df_assignment_std = standardize_assignment(df_assignment_raw)

            staff_list = sorted(df_assignment_std["staff_name"].dropna().unique().tolist())

            default_staff = APP_CONFIG["default_staff_filter"]
            default_index = 0
            for i, name in enumerate(staff_list):
                if build_staff_key(name) == build_staff_key(default_staff):
                    default_index = i
                    break

            selected_staff = st.selectbox(
                "Chọn nhân viên cần lọc tuyến",
                staff_list,
                index=default_index if staff_list else 0,
            )

            df_assignment_staff = df_assignment_std[
                df_assignment_std["staff_key"] == build_staff_key(selected_staff)
            ].copy()

            st.success(f"Đã đọc DS giao. Nhân viên đang chọn: {selected_staff}. Số mã thuộc tuyến: {len(df_assignment_staff)}")

            with st.expander("Xem trước DS giao đã lọc theo nhân viên"):
                preview_cols = ["staff_name", "ma_tt", "customer_name", "phone", "address", "generated_amount"]
                st.dataframe(df_assignment_staff[preview_cols].head(100), use_container_width=True)

        except Exception as e:
            st.error(f"Lỗi đọc DS giao kỳ cước: {e}")

    st.divider()

    st.markdown("### Bước B - Upload TN08 hóa đơn chưa thu")
    tn08_file = st.file_uploader(
        "Upload file TN08 - hóa đơn chưa thu",
        type=["xlsx", "xls", "csv"],
        key="tn08_file"
    )

    billing_period = st.text_input(
        "Kỳ cước hiển thị trong mẫu tin nhắn",
        value=APP_CONFIG["message"]["default_billing_period"],
        help="Ví dụ: 05 hoặc 05/2026."
    )

    if tn08_file is not None and df_assignment_std is not None and selected_staff:
        try:
            if tn08_file.name.lower().endswith((".xlsx", ".xls")):
                tn08_sheets = list_excel_sheets(tn08_file)
                # Ưu tiên sheet có tên gần với cn TN08 nếu có.
                default_tn08_index = 0
                for i, sheet in enumerate(tn08_sheets):
                    if "tn08" in normalize_key(sheet):
                        default_tn08_index = i
                        break

                tn08_sheet = st.selectbox("Chọn sheet TN08", tn08_sheets, index=default_tn08_index, key="tn08_sheet")
                df_tn08_raw = read_excel_or_csv(tn08_file, sheet_name=tn08_sheet)
            else:
                df_tn08_raw = read_excel_or_csv(tn08_file)

            df_tn08_std = standardize_tn08(df_tn08_raw)

            df_working = create_working_dataset(
                df_assignment=df_assignment_std,
                df_tn08=df_tn08_std,
                selected_staff=selected_staff
            )

            if df_working.empty:
                st.warning("Không tìm thấy dòng nào thuộc nhân viên đã chọn trong DS giao.")
            else:
                st.success("Đã tạo danh sách tuyến bằng logic VLOOKUP/merge từ DS giao sang TN08.")
                st.metric("Tổng tiền cần thu theo TN08", format_money(df_working["debt_amount"].sum()))

                missing_count = int((df_working["debt_amount"] <= 0).sum())
                if missing_count > 0:
                    st.warning(f"Có {missing_count} mã không tìm thấy nợ trong TN08 hoặc nợ = 0. Cần kiểm tra.")

                with st.expander("Xem trước dữ liệu sau khi dò TN08"):
                    show_cols = [
                        "staff_name", "ma_tt", "customer_name", "phone",
                        "debt_amount_display", "generated_amount_display",
                        "data_status"
                    ]
                    st.dataframe(df_working[show_cols].head(150), use_container_width=True)

                if st.button("Xác nhận lưu tuyến này vào CRM"):
                    batch_id = create_batch(
                        conn=conn,
                        import_type="working_dataset",
                        file_name=f"{assignment_file.name} + {tn08_file.name}",
                        staff_filter=selected_staff,
                        row_count=len(df_working),
                        total_amount=float(df_working["debt_amount"].sum())
                    )
                    insert_work_items(conn, batch_id, df_working)
                    st.success(f"Đã lưu tuyến {selected_staff} vào CRM. Batch ID: {batch_id}")

        except Exception as e:
            st.error(f"Lỗi xử lý TN08 hoặc merge dữ liệu: {e}")

    elif tn08_file is not None and df_assignment_std is None:
        st.info("Bạn cần upload và đọc DS giao kỳ cước trước, sau đó mới upload TN08 để dò tiền.")


# ------------------------------------------------------------
# TAB 2: BÀN LÀM VIỆC
# ------------------------------------------------------------
with tab_work:
    st.subheader("2. Bàn làm việc thu cước")

    df_current = get_current_work_items(conn)

    if df_current.empty:
        st.info("Chưa có dữ liệu tuyến. Hãy vào tab 1 để upload DS giao + TN08 và lưu vào CRM.")
    else:
        metrics = summary_metrics(df_current)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Tổng mã tuyến", metrics["total_rows"])
        c2.metric("Tổng tiền cần thu", format_money(metrics["total_debt"]))
        c3.metric("Đã đóng", metrics["paid_count"])
        c4.metric("Đã liên hệ chưa đóng", metrics["contacted_count"])
        c5.metric("Chưa liên hệ/thiếu số", metrics["uncontacted_count"])

        status_options = ["Tất cả"] + sorted(df_current["status_label"].dropna().unique().tolist())
        selected_status = st.selectbox("Lọc trạng thái", status_options)

        search_text = st.text_input("Tìm theo MA_TT / tên khách / số điện thoại")

        view_df = df_current.copy()

        if selected_status != "Tất cả":
            view_df = view_df[view_df["status_label"] == selected_status]

        if search_text:
            s = safe_lower(search_text)
            view_df = view_df[
                view_df["ma_tt"].apply(safe_lower).str.contains(s, na=False)
                | view_df["customer_display"].apply(safe_lower).str.contains(s, na=False)
                | view_df["phone"].apply(safe_lower).str.contains(s, na=False)
            ]

        show_cols = [
            "status_emoji", "status_label", "ma_tt", "customer_display",
            "phone", "address", "debt_amount_display",
            "data_status", "last_result", "promised_payment_date",
            "last_contacted_at"
        ]

        st.dataframe(
            view_df[show_cols],
            use_container_width=True,
            height=380,
        )

        if view_df.empty:
            st.warning("Không có dữ liệu phù hợp bộ lọc.")
        else:
            option_map = {
                f"{row['status_emoji']} {row['ma_tt']} | {row['customer_display']} | {row['debt_amount_display']}": row["ma_tt"]
                for _, row in view_df.iterrows()
            }

            selected_option = st.selectbox("Chọn mã để xử lý", list(option_map.keys()))
            selected_ma_tt = option_map[selected_option]

            row = df_current[df_current["ma_tt"] == selected_ma_tt].iloc[0].to_dict()

            st.divider()
            st.markdown("### Chi tiết khách hàng")

            left, right = st.columns([1.1, 1])

            with left:
                st.write(f"**Mã thanh toán:** {row.get('ma_tt', '')}")
                st.write(f"**Tên khách:** {row.get('customer_display', '')}")
                st.write(f"**Số điện thoại/Zalo ưu tiên:** {row.get('phone', '') or 'Chưa có'}")
                st.write(f"**Địa chỉ:** {row.get('address', '')}")
                st.write(f"**Tiền cần thu từ TN08:** {row.get('debt_amount_display', '')}")
                st.write(f"**Tiền phát sinh trong DS giao chỉ để tham khảo:** {row.get('generated_amount_display', '')}")
                st.write(f"**Trạng thái dữ liệu:** {row.get('data_status', '')}")
                st.write(f"**Trạng thái xử lý:** {row.get('status_emoji', '')} {row.get('status_label', '')}")

                st.markdown("#### Lịch sử gọi")
                history = get_interactions(conn, selected_ma_tt)
                st.dataframe(history, use_container_width=True, height=180)

                st.markdown("#### Lịch sử số liên hệ")
                contacts = get_contacts(conn, selected_ma_tt)
                st.dataframe(contacts, use_container_width=True, height=160)

            with right:
                st.markdown("#### Cập nhật kết quả cuộc gọi")

                with st.form("interaction_form", clear_on_submit=True):
                    result = st.selectbox("Kết quả cuộc gọi", APP_CONFIG["contact_results"])
                    promised_payment_date = st.text_input("Ngày hẹn thanh toán", placeholder="Ví dụ: 15/06/2026")
                    note = st.text_area("Ghi chú", placeholder="Ví dụ: gặp chị Lan kế toán, hẹn thứ 6 chuyển khoản...")
                    created_by = st.text_input("Người cập nhật", value=APP_CONFIG["message"]["staff_name"])
                    submitted = st.form_submit_button("Lưu lịch sử gọi")

                    if submitted:
                        add_interaction(
                            conn=conn,
                            ma_tt=selected_ma_tt,
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
                                conn=conn,
                                ma_tt=selected_ma_tt,
                                contact_value=phone,
                                contact_person=contact_person,
                                role=role,
                                note=contact_note,
                            )
                            st.success("Đã lưu số liên hệ mới.")
                            st.rerun()

            st.divider()
            st.markdown("### Mẫu tin nhắn Zalo")

            billing_period_work = st.text_input(
                "Kỳ cước trong mẫu tin",
                value=APP_CONFIG["message"]["default_billing_period"],
                key="billing_period_work"
            )

            message_text = render_message(billing_period_work)
            st.code(message_text, language="text")

            if st.button("Ghi nhận đã copy/gửi mẫu tin"):
                add_sent_message(
                    conn=conn,
                    ma_tt=selected_ma_tt,
                    message_text=message_text,
                    created_by=APP_CONFIG["message"]["staff_name"],
                )
                st.success("Đã ghi nhận lịch sử gửi/copy tin.")


# ------------------------------------------------------------
# TAB 3: BÁO CÁO
# ------------------------------------------------------------
with tab_report:
    st.subheader("3. Báo cáo KPI")

    df_current = get_current_work_items(conn)

    if df_current.empty:
        st.info("Chưa có dữ liệu để báo cáo.")
    else:
        metrics = summary_metrics(df_current)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Tổng mã tuyến", metrics["total_rows"])
        c2.metric("Tổng tiền còn cần thu", format_money(metrics["total_debt"]))
        c3.metric("Đã liên hệ chưa đóng", metrics["contacted_count"])
        c4.metric("Chưa liên hệ/thiếu số", metrics["uncontacted_count"])

        report_cols = [
            "status_emoji", "status_label", "staff_name", "ma_tt",
            "customer_display", "phone", "address",
            "debt_amount_display", "generated_amount_display",
            "data_status", "last_result", "promised_payment_date",
            "last_note", "last_contacted_at"
        ]

        st.dataframe(
            df_current[[c for c in report_cols if c in df_current.columns]],
            use_container_width=True,
            height=450,
        )

        st.download_button(
            "Tải báo cáo Excel",
            data=to_excel_bytes(df_current),
            file_name="bao_cao_thu_cuoc_vuong_thanh_thuan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ------------------------------------------------------------
# TAB 4: QUẢN TRỊ
# ------------------------------------------------------------
with tab_admin:
    st.subheader("4. Quản trị dữ liệu")

    st.warning(
        "Chỉ dùng khu vực này để kiểm tra hoặc reset dữ liệu khi cần làm lại từ đầu."
    )

    batches = pd.read_sql_query(
        """
        SELECT batch_id, import_type, file_name, staff_filter, row_count, total_amount, created_at
        FROM import_batches
        ORDER BY batch_id DESC
        LIMIT 50
        """,
        conn
    )
    st.markdown("### Lịch sử import")
    st.dataframe(batches, use_container_width=True)

    stats = []
    for table in ["work_items", "paid_updates", "interactions", "contacts", "sent_messages"]:
        try:
            count = conn.execute(f"SELECT COUNT(*) AS c FROM {table}").fetchone()["c"]
        except Exception:
            count = "error"
        stats.append({"table": table, "rows": count})

    st.markdown("### Thống kê bảng")
    st.dataframe(pd.DataFrame(stats), use_container_width=True)

    with st.expander("Reset database"):
        st.write("Gõ RESET để xác nhận xóa toàn bộ dữ liệu CRM.")
        confirm = st.text_input("Xác nhận")
        if st.button("Xóa toàn bộ dữ liệu") and confirm == "RESET":
            conn.executescript("""
                DELETE FROM sent_messages;
                DELETE FROM contacts;
                DELETE FROM interactions;
                DELETE FROM paid_updates;
                DELETE FROM work_items;
                DELETE FROM import_batches;
            """)
            conn.commit()
            st.success("Đã reset dữ liệu.")
            st.rerun()
