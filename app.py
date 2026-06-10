# ============================================================
# APP.PY - MINI CRM THU CƯỚC KHDN VNPT
# Bản đúng logic:
# DS giao kỳ cước -> lọc nhân viên -> lấy MA_TT -> dò TN08 -> lấy Total Nợ thu vét
# Không dùng file Giang làm nguồn chính.
# Không dùng tiền phát sinh làm tiền cần thu.
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
# 1. CONFIGURATION
# ============================================================

APP_CONFIG = {
    "app_title": "📞 Mini-CRM Thu Cước KHDN - VNPT",
    "db_path": "crm_vnpt_thu_cuoc.db",
    "default_staff": "Vương Thanh Thuận",

    "message": {
        "unit_name": "VNPT Long Thành- Nhơn Trạch",
        "website": "https://vnptdongnai.vn/",
        "payment_deadline": "ngày 15",
        "staff_name": "Thuận",
        "staff_phone": "0837892579",
        "default_period": "05",
        "template": """{unit_name} thông báo:
Đã có thông báo cước kỳ cước tháng {period}. Anh/chị truy cập trang {website} để lấy thông báo cước và hóa đơn của công ty.
Anh/chị vui lòng thanh toán cước trước {payment_deadline}. Sau ngày 15 hệ thống sẽ tự động đưa lưới khóa khi ghi nhận còn tồn nợ.
Liên hệ nhân viên kinh doanh: {staff_name}: {staff_phone} để được hỗ trợ gạch nợ sớm nhất sau khi thanh toán.
Nếu đã thanh toán cước, anh/chị vui lòng bỏ qua tin nhắn trên.
Trân trọng."""
    },

    "assignment_columns": {
        "staff_name": [
            "nhân viên thu cước", "nhan vien thu cuoc",
            "nhân viên", "nhan vien", "nv thu", "người phụ trách", "nguoi phu trach"
        ],
        "ma_tt": [
            "mã thanh toán", "ma thanh toan", "mã tt", "ma_tt", "ma tt"
        ],
        "customer_name": [
            "tên thanh toán", "ten thanh toan", "tên khách hàng", "ten khach hang",
            "khách hàng", "khach hang"
        ],
        "phone": [
            "số dt liên hệ", "so dt lien he", "số điện thoại", "so dien thoai",
            "điện thoại", "dien thoai", "sdt", "phone"
        ],
        "address": [
            "địa chỉ thanh toán", "dia chi thanh toan", "địa chỉ", "dia chi"
        ],
        "generated_amount": [
            "tiền phát sinh", "tien phat sinh", "phát sinh", "phat sinh"
        ],
    },

    "tn08_columns": {
        "ma_tt": [
            "ma_tt", "mã tt", "ma tt", "mã thanh toán", "ma thanh toan"
        ],
        "debt_amount": [
            "total nợ thu vét", "total no thu vet", "total nợ thu vét",
            "nợ thu vét", "no thu vet"
        ],
        "customer_name_tn08": [
            "tên khách hàng", "ten khach hang", "tên thanh toán", "ten thanh toan"
        ],
        "address_tn08": [
            "địa chỉ kh", "dia chi kh", "địa chỉ", "dia chi"
        ],
    },

    "paid_columns": {
        "ma_tt": [
            "ma_tt", "mã tt", "ma tt", "mã thanh toán", "ma thanh toan"
        ],
        "paid_amount": [
            "số tiền", "so tien", "tiền đóng", "tien dong", "số tiền đã đóng"
        ],
        "paid_date": [
            "ngày đóng", "ngay dong", "ngày thanh toán", "ngay thanh toan"
        ],
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

    "status": {
        "paid": ("🟢", "Đã đóng"),
        "contacted": ("🟡", "Đã liên hệ chưa đóng"),
        "uncontacted": ("🔴", "Chưa liên hệ"),
        "no_phone": ("🔴", "Thiếu số điện thoại"),
        "need_check": ("⚪", "Cần kiểm tra dữ liệu"),
    }
}


# ============================================================
# 2. UTILS
# ============================================================

def normalize_text(value) -> str:
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
    return re.sub(r"\s+", " ", text)


def remove_accents(value) -> str:
    text = normalize_text(value).lower()
    text = unicodedata.normalize("NFD", text)
    return "".join(ch for ch in text if unicodedata.category(ch) != "Mn")


def normalize_key(value) -> str:
    text = remove_accents(value)
    return re.sub(r"[^a-z0-9]+", "_", text).strip("_")


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
    try:
        if pd.isna(value):
            return 0.0
    except Exception:
        pass

    if isinstance(value, (int, float)):
        return float(value)

    text = normalize_text(value)
    text = text.replace("đ", "").replace("Đ", "").replace("VND", "").replace("vnd", "")
    text = text.replace(" ", "")

    if "." in text and "," in text:
        text = text.replace(".", "").replace(",", ".")
    else:
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


def staff_key(value) -> str:
    return normalize_key(value)


def find_col(df: pd.DataFrame, candidates: list[str]):
    normalized = {normalize_key(col): col for col in df.columns}
    for c in candidates:
        k = normalize_key(c)
        if k in normalized:
            return normalized[k]
    return None


# ============================================================
# 3. DATABASE
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

        CREATE TABLE IF NOT EXISTS interactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ma_tt TEXT NOT NULL,
            result TEXT NOT NULL,
            promised_payment_date TEXT,
            note TEXT,
            created_by TEXT,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS contacts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ma_tt TEXT NOT NULL,
            contact_value TEXT NOT NULL,
            contact_person TEXT,
            role TEXT,
            note TEXT,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS sent_messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ma_tt TEXT NOT NULL,
            message_text TEXT NOT NULL,
            created_by TEXT,
            created_at TEXT NOT NULL
        );
    """)
    conn.commit()


def create_batch(conn, import_type, file_name, staff_filter, row_count, total_amount):
    cur = conn.execute("""
        INSERT INTO import_batches
        (import_type, file_name, staff_filter, row_count, total_amount, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (import_type, file_name, staff_filter, row_count, total_amount, now_str()))
    conn.commit()
    return int(cur.lastrowid)


def get_latest_batch_id(conn, import_type):
    row = conn.execute("""
        SELECT batch_id FROM import_batches
        WHERE import_type = ?
        ORDER BY batch_id DESC
        LIMIT 1
    """, (import_type,)).fetchone()
    return int(row["batch_id"]) if row else None


def insert_work_items(conn, batch_id, df):
    rows = []
    for _, r in df.iterrows():
        rows.append((
            batch_id,
            r.get("staff_name", ""),
            r.get("ma_tt", ""),
            r.get("customer_name", ""),
            r.get("phone", ""),
            r.get("address", ""),
            float(r.get("generated_amount", 0) or 0),
            float(r.get("debt_amount", 0) or 0),
            r.get("customer_name_tn08", ""),
            r.get("address_tn08", ""),
            r.get("data_status", ""),
            now_str()
        ))

    conn.executemany("""
        INSERT INTO work_items
        (batch_id, staff_name, ma_tt, customer_name, phone, address,
         generated_amount, debt_amount, tn08_customer_name, tn08_address,
         data_status, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, rows)
    conn.commit()


def insert_paid_updates(conn, batch_id, df):
    rows = []
    for _, r in df.iterrows():
        rows.append((
            batch_id,
            r.get("ma_tt", ""),
            float(r.get("paid_amount", 0) or 0),
            r.get("paid_date", ""),
            now_str()
        ))
    conn.executemany("""
        INSERT INTO paid_updates
        (batch_id, ma_tt, paid_amount, paid_date, created_at)
        VALUES (?, ?, ?, ?, ?)
    """, rows)
    conn.commit()


def add_interaction(conn, ma_tt, result, promised_payment_date, note, created_by):
    conn.execute("""
        INSERT INTO interactions
        (ma_tt, result, promised_payment_date, note, created_by, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (ma_tt, result, promised_payment_date, note, created_by, now_str()))
    conn.commit()


def add_contact(conn, ma_tt, contact_value, contact_person, role, note):
    conn.execute("""
        INSERT INTO contacts
        (ma_tt, contact_value, contact_person, role, note, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (ma_tt, contact_value, contact_person, role, note, now_str()))
    conn.commit()


def add_sent_message(conn, ma_tt, message_text, created_by):
    conn.execute("""
        INSERT INTO sent_messages
        (ma_tt, message_text, created_by, created_at)
        VALUES (?, ?, ?, ?)
    """, (ma_tt, message_text, created_by, now_str()))
    conn.commit()


# ============================================================
# 4. ETL
# ============================================================

def list_sheets(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names
    uploaded_file.seek(0)
    return sheets


def read_file(uploaded_file, sheet_name=None):
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded_file)
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding="latin1")
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)


def standardize_assignment(df_raw, fallback_staff=None):
    df = df_raw.copy()
    df.columns = [normalize_text(c) for c in df.columns]

    out = pd.DataFrame()

    ma_col = find_col(df, APP_CONFIG["assignment_columns"]["ma_tt"])
    if not ma_col:
        raise ValueError("DS giao thiếu cột Mã thanh toán.")

    staff_col = find_col(df, APP_CONFIG["assignment_columns"]["staff_name"])

    if staff_col:
        out["staff_name"] = df[staff_col]
    else:
        # Nếu bạn chọn sheet đã lọc sẵn như sheet Thuan, có thể không cần cột nhân viên.
        out["staff_name"] = fallback_staff or APP_CONFIG["default_staff"]

    out["ma_tt"] = df[ma_col]

    for field in ["customer_name", "phone", "address", "generated_amount"]:
        col = find_col(df, APP_CONFIG["assignment_columns"][field])
        out[field] = df[col] if col else ""

    out["staff_name"] = out["staff_name"].apply(normalize_text)
    out["staff_key"] = out["staff_name"].apply(staff_key)
    out["ma_tt"] = out["ma_tt"].apply(normalize_ma_tt)
    out["customer_name"] = out["customer_name"].apply(normalize_text)
    out["phone"] = out["phone"].apply(normalize_phone)
    out["address"] = out["address"].apply(normalize_text)
    out["generated_amount"] = out["generated_amount"].apply(parse_money)

    out = out[out["ma_tt"].astype(str).str.len() > 0].copy()
    return out


def standardize_tn08(df_raw):
    df = df_raw.copy()
    df.columns = [normalize_text(c) for c in df.columns]

    ma_col = find_col(df, APP_CONFIG["tn08_columns"]["ma_tt"])
    debt_col = find_col(df, APP_CONFIG["tn08_columns"]["debt_amount"])

    if not ma_col:
        raise ValueError("TN08 thiếu cột MA_TT.")
    if not debt_col:
        raise ValueError("TN08 thiếu cột Total Nợ thu vét. App không dùng Tiền phát sinh làm tiền cần thu.")

    out = pd.DataFrame()
    out["ma_tt"] = df[ma_col]
    out["debt_amount"] = df[debt_col]

    for field in ["customer_name_tn08", "address_tn08"]:
        col = find_col(df, APP_CONFIG["tn08_columns"][field])
        out[field] = df[col] if col else ""

    out["ma_tt"] = out["ma_tt"].apply(normalize_ma_tt)
    out["debt_amount"] = out["debt_amount"].apply(parse_money)
    out["customer_name_tn08"] = out["customer_name_tn08"].apply(normalize_text)
    out["address_tn08"] = out["address_tn08"].apply(normalize_text)

    out = out[out["ma_tt"].astype(str).str.len() > 0].copy()

    # Nếu TN08 trùng MA_TT, cộng nợ để tránh nhân dòng.
    out = out.groupby("ma_tt", as_index=False).agg({
        "debt_amount": "sum",
        "customer_name_tn08": "first",
        "address_tn08": "first",
    })
    return out


def standardize_paid(df_raw):
    df = df_raw.copy()
    df.columns = [normalize_text(c) for c in df.columns]

    ma_col = find_col(df, APP_CONFIG["paid_columns"]["ma_tt"])
    if not ma_col:
        raise ValueError("File đã đóng thiếu cột MA_TT / Mã thanh toán.")

    out = pd.DataFrame()
    out["ma_tt"] = df[ma_col]

    amount_col = find_col(df, APP_CONFIG["paid_columns"]["paid_amount"])
    out["paid_amount"] = df[amount_col] if amount_col else 0

    date_col = find_col(df, APP_CONFIG["paid_columns"]["paid_date"])
    out["paid_date"] = df[date_col] if date_col else ""

    out["ma_tt"] = out["ma_tt"].apply(normalize_ma_tt)
    out["paid_amount"] = out["paid_amount"].apply(parse_money)
    out["paid_date"] = out["paid_date"].apply(normalize_text)
    out = out[out["ma_tt"].astype(str).str.len() > 0].copy()
    return out


def build_working_dataset(df_assignment, df_tn08, selected_staff):
    selected_key = staff_key(selected_staff)

    df_staff = df_assignment[df_assignment["staff_key"] == selected_key].copy()
    if df_staff.empty:
        return pd.DataFrame()

    # Tối ưu: chỉ giữ các MA_TT thuộc tuyến trước khi merge.
    route_ma = set(df_staff["ma_tt"].dropna().astype(str))
    df_tn08_small = df_tn08[df_tn08["ma_tt"].isin(route_ma)].copy()

    merged = df_staff.merge(
        df_tn08_small[["ma_tt", "debt_amount", "customer_name_tn08", "address_tn08"]],
        on="ma_tt",
        how="left"
    )

    merged["debt_amount"] = merged["debt_amount"].fillna(0)
    merged["customer_name_tn08"] = merged["customer_name_tn08"].fillna("")
    merged["address_tn08"] = merged["address_tn08"].fillna("")

    merged["data_status"] = merged["debt_amount"].apply(
        lambda x: "Có nợ theo TN08" if x > 0 else "Không có trong TN08 / nợ = 0 / cần kiểm tra"
    )

    merged["debt_amount_display"] = merged["debt_amount"].apply(format_money)
    merged["generated_amount_display"] = merged["generated_amount"].apply(format_money)
    return merged


# ============================================================
# 5. BUSINESS LOGIC
# ============================================================

def get_current_items(conn):
    latest_work_batch = get_latest_batch_id(conn, "working_dataset")
    latest_paid_batch = get_latest_batch_id(conn, "paid")

    if latest_work_batch is None:
        return pd.DataFrame()

    if latest_paid_batch is None:
        paid_cte = "SELECT '' AS ma_tt, 0 AS paid_amount WHERE 1=0"
        params = [latest_work_batch]
    else:
        paid_cte = """
            SELECT ma_tt, SUM(paid_amount) AS paid_amount
            FROM paid_updates
            WHERE batch_id = ?
            GROUP BY ma_tt
        """
        params = [latest_paid_batch, latest_work_batch]

    query = f"""
        WITH paid AS ({paid_cte}),
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
            w.*,
            COALESCE(p.paid_amount, 0) AS paid_amount,
            li.result AS last_result,
            li.promised_payment_date,
            li.note AS last_note,
            li.created_by AS last_created_by,
            li.created_at AS last_contacted_at,
            COALESCE(NULLIF(lc.contact_value, ''), w.phone) AS current_phone
        FROM work_items w
        LEFT JOIN paid p ON p.ma_tt = w.ma_tt
        LEFT JOIN last_interaction li ON li.ma_tt = w.ma_tt
        LEFT JOIN latest_contact lc ON lc.ma_tt = w.ma_tt
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

        if not normalize_text(row.get("current_phone", "")):
            return "no_phone"

        if normalize_text(row.get("last_contacted_at", "")):
            return "contacted"

        return "uncontacted"

    df["status_code"] = df.apply(calc_status, axis=1)
    df["status_emoji"] = df["status_code"].apply(lambda x: APP_CONFIG["status"][x][0])
    df["status_label"] = df["status_code"].apply(lambda x: APP_CONFIG["status"][x][1])
    df["debt_amount_display"] = df["debt_amount"].apply(format_money)
    df["generated_amount_display"] = df["generated_amount"].apply(format_money)
    df["paid_amount_display"] = df["paid_amount"].apply(format_money)
    df["customer_display"] = df["customer_name"].where(
        df["customer_name"].astype(str).str.len() > 0,
        df["tn08_customer_name"]
    )

    return df


def get_interactions(conn, ma_tt):
    return pd.read_sql_query("""
        SELECT result, promised_payment_date, note, created_by, created_at
        FROM interactions
        WHERE ma_tt = ?
        ORDER BY created_at DESC
    """, conn, params=(ma_tt,))


def get_contacts(conn, ma_tt):
    return pd.read_sql_query("""
        SELECT contact_value, contact_person, role, note, created_at
        FROM contacts
        WHERE ma_tt = ?
        ORDER BY created_at DESC
    """, conn, params=(ma_tt,))


def render_message(period):
    m = APP_CONFIG["message"]
    return m["template"].format(
        unit_name=m["unit_name"],
        period=period,
        website=m["website"],
        payment_deadline=m["payment_deadline"],
        staff_name=m["staff_name"],
        staff_phone=m["staff_phone"],
    )


def export_excel(df):
    cols = [
        "status_emoji", "status_label", "staff_name", "ma_tt", "customer_display",
        "current_phone", "address", "debt_amount", "debt_amount_display",
        "generated_amount", "generated_amount_display",
        "paid_amount", "data_status", "last_result",
        "promised_payment_date", "last_note", "last_contacted_at"
    ]
    out = df[[c for c in cols if c in df.columns]].copy()
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="Bao cao")
    bio.seek(0)
    return bio.getvalue()


# ============================================================
# 6. UI
# ============================================================

st.set_page_config(page_title="Mini-CRM Thu Cước KHDN", page_icon="📞", layout="wide")

conn = get_conn()
init_db(conn)

st.title(APP_CONFIG["app_title"])
st.caption("Đúng logic: DS giao kỳ cước → lọc Vương Thanh Thuận → dò TN08 bằng MA_TT → lấy Total Nợ thu vét.")

tabs = st.tabs(["1. Nạp dữ liệu", "2. Bàn làm việc", "3. Báo cáo", "4. Quản trị dữ liệu"])

with tabs[0]:
    st.subheader("1. Nạp dữ liệu đúng quy trình")

    st.info(
        "Bước đúng: upload DS giao kỳ cước trước, chọn nhân viên, sau đó upload TN08 để app dò MA_TT sang Total Nợ thu vét."
    )

    assign_file = st.file_uploader("A. Upload DS giao kỳ cước", type=["xlsx", "xls", "csv"], key="assign")
    df_assignment = None
    selected_staff = None

    if assign_file:
        try:
            if assign_file.name.lower().endswith((".xlsx", ".xls")):
                sheets = list_sheets(assign_file)
                default_idx = 0
                for i, s in enumerate(sheets):
                    if normalize_key(s) in ["ds", "thuan", "giao"]:
                        default_idx = i
                        break
                sheet = st.selectbox("Chọn sheet DS giao", sheets, index=default_idx)
                raw_assign = read_file(assign_file, sheet)
            else:
                raw_assign = read_file(assign_file)

            fallback_staff = APP_CONFIG["default_staff"]
            df_assignment = standardize_assignment(raw_assign, fallback_staff=fallback_staff)

            staff_list = sorted(df_assignment["staff_name"].dropna().unique().tolist())
            default_idx = 0
            for i, s in enumerate(staff_list):
                if staff_key(s) == staff_key(APP_CONFIG["default_staff"]):
                    default_idx = i
                    break

            selected_staff = st.selectbox("Chọn nhân viên cần lọc", staff_list, index=default_idx)
            route_df = df_assignment[df_assignment["staff_key"] == staff_key(selected_staff)].copy()

            st.success(f"Đã lọc nhân viên {selected_staff}: {len(route_df)} mã thanh toán.")
            st.dataframe(
                route_df[["staff_name", "ma_tt", "customer_name", "phone", "address", "generated_amount"]].head(100),
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Lỗi đọc DS giao kỳ cước: {e}")

    st.divider()

    tn08_file = st.file_uploader("B. Upload TN08 hóa đơn chưa thu", type=["xlsx", "xls", "csv"], key="tn08")
    period = st.text_input("Kỳ cước cho mẫu tin nhắn", value=APP_CONFIG["message"]["default_period"])

    if tn08_file and df_assignment is not None and selected_staff:
        try:
            if tn08_file.name.lower().endswith((".xlsx", ".xls")):
                sheets = list_sheets(tn08_file)
                default_idx = 0
                for i, s in enumerate(sheets):
                    if "tn08" in normalize_key(s):
                        default_idx = i
                        break
                tn08_sheet = st.selectbox("Chọn sheet TN08", sheets, index=default_idx)
                raw_tn08 = read_file(tn08_file, tn08_sheet)
            else:
                raw_tn08 = read_file(tn08_file)

            df_tn08 = standardize_tn08(raw_tn08)
            working = build_working_dataset(df_assignment, df_tn08, selected_staff)

            if working.empty:
                st.warning("Không có mã nào thuộc nhân viên đã chọn.")
            else:
                st.success("Đã dò TN08 xong. Tiền cần thu được lấy từ cột Total Nợ thu vét.")
                st.metric("Tổng tiền cần thu theo TN08", format_money(working["debt_amount"].sum()))
                st.dataframe(
                    working[["staff_name", "ma_tt", "customer_name", "phone", "debt_amount_display", "generated_amount_display", "data_status"]].head(200),
                    use_container_width=True
                )

                if st.button("Lưu tuyến này vào CRM"):
                    batch_id = create_batch(
                        conn,
                        "working_dataset",
                        f"{assign_file.name} + {tn08_file.name}",
                        selected_staff,
                        len(working),
                        float(working["debt_amount"].sum())
                    )
                    insert_work_items(conn, batch_id, working)
                    st.success(f"Đã lưu vào CRM. Batch ID: {batch_id}")

        except Exception as e:
            st.error(f"Lỗi xử lý TN08: {e}")

    st.divider()

    paid_file = st.file_uploader("C. Upload file khách đã đóng nếu có", type=["xlsx", "xls", "csv"], key="paid")
    if paid_file:
        try:
            raw_paid = read_file(paid_file)
            paid = standardize_paid(raw_paid)
            st.dataframe(paid.head(100), use_container_width=True)
            if st.button("Lưu danh sách đã đóng"):
                batch_id = create_batch(conn, "paid", paid_file.name, "", len(paid), float(paid["paid_amount"].sum()))
                insert_paid_updates(conn, batch_id, paid)
                st.success(f"Đã lưu danh sách đã đóng. Batch ID: {batch_id}")
        except Exception as e:
            st.error(f"Lỗi đọc file đã đóng: {e}")

with tabs[1]:
    st.subheader("2. Bàn làm việc")

    df = get_current_items(conn)

    if df.empty:
        st.info("Chưa có dữ liệu. Hãy nạp DS giao + TN08 ở tab 1.")
    else:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Tổng mã", len(df))
        c2.metric("Tổng tiền còn thu", format_money(df[df["status_code"] != "paid"]["debt_amount"].sum()))
        c3.metric("Đã liên hệ", int((df["status_code"] == "contacted").sum()))
        c4.metric("Chưa liên hệ/thiếu số", int(df["status_code"].isin(["uncontacted", "no_phone"]).sum()))

        statuses = ["Tất cả"] + sorted(df["status_label"].unique().tolist())
        status_filter = st.selectbox("Lọc trạng thái", statuses)

        q = st.text_input("Tìm MA_TT / tên khách / số điện thoại")

        view = df.copy()
        if status_filter != "Tất cả":
            view = view[view["status_label"] == status_filter]

        if q:
            qk = remove_accents(q)
            view = view[
                view["ma_tt"].apply(remove_accents).str.contains(qk, na=False)
                | view["customer_display"].apply(remove_accents).str.contains(qk, na=False)
                | view["current_phone"].apply(remove_accents).str.contains(qk, na=False)
            ]

        show_cols = [
            "status_emoji", "status_label", "ma_tt", "customer_display",
            "current_phone", "address", "debt_amount_display", "data_status",
            "last_result", "promised_payment_date", "last_contacted_at"
        ]
        st.dataframe(view[show_cols], use_container_width=True, height=360)

        if not view.empty:
            options = {
                f"{r['status_emoji']} {r['ma_tt']} | {r['customer_display']} | {r['debt_amount_display']}": r["ma_tt"]
                for _, r in view.iterrows()
            }

            choice = st.selectbox("Chọn mã để xử lý", list(options.keys()))
            ma_tt = options[choice]
            row = df[df["ma_tt"] == ma_tt].iloc[0].to_dict()

            st.divider()
            left, right = st.columns([1.1, 1])

            with left:
                st.markdown("### Chi tiết khách hàng")
                st.write(f"**Mã thanh toán:** {row.get('ma_tt', '')}")
                st.write(f"**Tên khách:** {row.get('customer_display', '')}")
                st.write(f"**Số điện thoại/Zalo ưu tiên:** {row.get('current_phone', '') or 'Chưa có'}")
                st.write(f"**Địa chỉ:** {row.get('address', '')}")
                st.write(f"**Tiền cần thu từ TN08:** {row.get('debt_amount_display', '')}")
                st.write(f"**Tiền phát sinh trong DS giao chỉ để tham khảo:** {row.get('generated_amount_display', '')}")
                st.write(f"**Trạng thái xử lý:** {row.get('status_emoji', '')} {row.get('status_label', '')}")

                st.markdown("#### Lịch sử gọi")
                st.dataframe(get_interactions(conn, ma_tt), use_container_width=True, height=180)

                st.markdown("#### Lịch sử số liên hệ")
                st.dataframe(get_contacts(conn, ma_tt), use_container_width=True, height=140)

            with right:
                st.markdown("### Cập nhật kết quả cuộc gọi")
                with st.form("call_form", clear_on_submit=True):
                    result = st.selectbox("Kết quả cuộc gọi", APP_CONFIG["contact_results"])
                    promised = st.text_input("Ngày hẹn thanh toán", placeholder="VD: 15/06")
                    note = st.text_area("Ghi chú")
                    created_by = st.text_input("Người cập nhật", value=APP_CONFIG["message"]["staff_name"])
                    submit = st.form_submit_button("Lưu lịch sử gọi")

                    if submit:
                        add_interaction(conn, ma_tt, result, promised, note, created_by)
                        st.success("Đã lưu lịch sử gọi.")
                        st.rerun()

                st.markdown("### Cập nhật số Zalo/SĐT mới")
                with st.form("contact_form", clear_on_submit=True):
                    new_phone = st.text_input("Số Zalo/SĐT mới")
                    person = st.text_input("Tên người phụ trách")
                    role = st.text_input("Vai trò", value="Kế toán")
                    contact_note = st.text_area("Ghi chú số liên hệ")
                    submit_contact = st.form_submit_button("Lưu số liên hệ")

                    if submit_contact:
                        phone = normalize_phone(new_phone)
                        if not phone:
                            st.error("Số điện thoại chưa hợp lệ.")
                        else:
                            add_contact(conn, ma_tt, phone, person, role, contact_note)
                            st.success("Đã lưu số liên hệ mới.")
                            st.rerun()

            st.divider()
            st.markdown("### Mẫu tin nhắn Zalo")
            period_msg = st.text_input("Kỳ cước trong tin nhắn", value=APP_CONFIG["message"]["default_period"], key="period_msg")
            msg = render_message(period_msg)
            st.code(msg, language="text")

            if st.button("Ghi nhận đã copy/gửi tin"):
                add_sent_message(conn, ma_tt, msg, APP_CONFIG["message"]["staff_name"])
                st.success("Đã ghi nhận.")

with tabs[2]:
    st.subheader("3. Báo cáo")

    df = get_current_items(conn)

    if df.empty:
        st.info("Chưa có dữ liệu.")
    else:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Tổng mã", len(df))
        c2.metric("Tổng tiền còn thu", format_money(df[df["status_code"] != "paid"]["debt_amount"].sum()))
        c3.metric("Đã liên hệ", int((df["status_code"] == "contacted").sum()))
        c4.metric("Chưa liên hệ/thiếu số", int(df["status_code"].isin(["uncontacted", "no_phone"]).sum()))

        st.dataframe(df, use_container_width=True, height=450)

        st.download_button(
            "Tải báo cáo Excel",
            data=export_excel(df),
            file_name="bao_cao_thu_cuoc_vuong_thanh_thuan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

with tabs[3]:
    st.subheader("4. Quản trị dữ liệu")

    st.markdown("### Lịch sử import")
    batches = pd.read_sql_query("""
        SELECT batch_id, import_type, file_name, staff_filter, row_count, total_amount, created_at
        FROM import_batches
        ORDER BY batch_id DESC
        LIMIT 50
    """, conn)
    st.dataframe(batches, use_container_width=True)

    st.markdown("### Reset database")
    st.warning("Dùng khi muốn xóa dữ liệu test để import lại từ đầu.")
    confirm = st.text_input("Gõ RESET để xác nhận")
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
