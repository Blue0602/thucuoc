from __future__ import annotations
from pathlib import Path
import sqlite3
from typing import Dict, Any, Iterable, Optional
import pandas as pd
from .utils import current_timestamp


def get_connection(db_path: str) -> sqlite3.Connection:
    Path(db_path).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def init_db(conn: sqlite3.Connection) -> None:
    conn.executescript('''
    CREATE TABLE IF NOT EXISTS import_batches (
        batch_id INTEGER PRIMARY KEY AUTOINCREMENT,
        import_type TEXT NOT NULL,
        file_name TEXT,
        sheet_name TEXT,
        row_count INTEGER DEFAULT 0,
        total_amount REAL DEFAULT 0,
        status TEXT DEFAULT 'success',
        note TEXT,
        imported_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS customers (
        ma_tt TEXT PRIMARY KEY,
        representative_code TEXT,
        customer_name TEXT,
        address TEXT,
        current_phone TEXT,
        assigned_staff TEXT,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS debt_snapshots (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        batch_id INTEGER NOT NULL,
        ma_tt TEXT NOT NULL,
        billing_period TEXT,
        debt_amount REAL DEFAULT 0,
        source_customer_name TEXT,
        source_address TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY (batch_id) REFERENCES import_batches(batch_id)
    );
    CREATE INDEX IF NOT EXISTS idx_debt_batch_ma_tt ON debt_snapshots(batch_id, ma_tt);
    CREATE TABLE IF NOT EXISTS paid_updates (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        batch_id INTEGER NOT NULL,
        ma_tt TEXT NOT NULL,
        paid_amount REAL DEFAULT 0,
        paid_date TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY (batch_id) REFERENCES import_batches(batch_id)
    );
    CREATE INDEX IF NOT EXISTS idx_paid_ma_tt ON paid_updates(ma_tt);
    CREATE TABLE IF NOT EXISTS interactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ma_tt TEXT NOT NULL,
        result TEXT NOT NULL,
        note TEXT,
        promised_payment_date TEXT,
        created_by TEXT,
        created_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_interactions_ma_tt ON interactions(ma_tt);
    CREATE TABLE IF NOT EXISTS customer_contacts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ma_tt TEXT NOT NULL,
        contact_type TEXT DEFAULT 'zalo',
        contact_value TEXT NOT NULL,
        contact_person_name TEXT,
        role TEXT,
        is_primary INTEGER DEFAULT 1,
        verification_status TEXT DEFAULT 'verified',
        source TEXT DEFAULT 'phone_call',
        note TEXT,
        created_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_contacts_ma_tt ON customer_contacts(ma_tt);
    CREATE TABLE IF NOT EXISTS sent_messages (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ma_tt TEXT NOT NULL,
        template_name TEXT NOT NULL,
        message_text TEXT NOT NULL,
        copied_by TEXT,
        created_at TEXT NOT NULL
    );
    ''')
    conn.commit()


def create_batch(conn, import_type, file_name, sheet_name, row_count, total_amount=0, status="success", note=None) -> int:
    cur = conn.execute('''INSERT INTO import_batches
        (import_type, file_name, sheet_name, row_count, total_amount, status, note, imported_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
        (import_type, file_name, sheet_name, row_count, total_amount, status, note, current_timestamp()))
    conn.commit()
    return int(cur.lastrowid)


def upsert_customers(conn: sqlite3.Connection, records: Iterable[Dict[str, Any]]) -> int:
    count, now = 0, current_timestamp()
    for r in records:
        ma_tt = r.get("ma_tt")
        if not ma_tt:
            continue
        conn.execute('''INSERT INTO customers
        (ma_tt, representative_code, customer_name, address, current_phone, assigned_staff, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(ma_tt) DO UPDATE SET
            representative_code = COALESCE(NULLIF(excluded.representative_code,''), customers.representative_code),
            customer_name = COALESCE(NULLIF(excluded.customer_name,''), customers.customer_name),
            address = COALESCE(NULLIF(excluded.address,''), customers.address),
            current_phone = COALESCE(NULLIF(excluded.current_phone,''), customers.current_phone),
            assigned_staff = COALESCE(NULLIF(excluded.assigned_staff,''), customers.assigned_staff),
            updated_at = excluded.updated_at''',
            (ma_tt, r.get("representative_code"), r.get("customer_name"), r.get("address"), r.get("phone"), r.get("staff_name"), now, now))
        count += 1
    conn.commit()
    return count


def insert_debt_snapshots(conn, batch_id, rows, billing_period) -> int:
    now = current_timestamp()
    payload = [(batch_id, r.get("ma_tt"), billing_period, float(r.get("debt_amount") or 0), r.get("customer_name"), r.get("address"), now) for r in rows if r.get("ma_tt")]
    conn.executemany('''INSERT INTO debt_snapshots
        (batch_id, ma_tt, billing_period, debt_amount, source_customer_name, source_address, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)''', payload)
    conn.commit()
    return len(payload)


def insert_paid_updates(conn, batch_id, rows) -> int:
    now = current_timestamp()
    payload = [(batch_id, r.get("ma_tt"), float(r.get("paid_amount") or 0), r.get("paid_date"), now) for r in rows if r.get("ma_tt")]
    conn.executemany('''INSERT INTO paid_updates
        (batch_id, ma_tt, paid_amount, paid_date, created_at)
        VALUES (?, ?, ?, ?, ?)''', payload)
    conn.commit()
    return len(payload)


def add_interaction(conn, ma_tt, result, note="", promised_payment_date="", created_by="") -> None:
    conn.execute('''INSERT INTO interactions
        (ma_tt, result, note, promised_payment_date, created_by, created_at)
        VALUES (?, ?, ?, ?, ?, ?)''',
        (ma_tt, result, note, promised_payment_date, created_by, current_timestamp()))
    conn.commit()


def add_contact(conn, ma_tt, contact_value, contact_type="zalo", contact_person_name="", role="", note="") -> None:
    conn.execute('''INSERT INTO customer_contacts
        (ma_tt, contact_type, contact_value, contact_person_name, role, is_primary, verification_status, source, note, created_at)
        VALUES (?, ?, ?, ?, ?, 1, 'verified', 'phone_call', ?, ?)''',
        (ma_tt, contact_type, contact_value, contact_person_name, role, note, current_timestamp()))
    conn.execute("UPDATE customers SET current_phone = ?, updated_at = ? WHERE ma_tt = ?", (contact_value, current_timestamp(), ma_tt))
    conn.commit()


def add_sent_message(conn, ma_tt, template_name, message_text, copied_by="") -> None:
    conn.execute('''INSERT INTO sent_messages (ma_tt, template_name, message_text, copied_by, created_at)
        VALUES (?, ?, ?, ?, ?)''', (ma_tt, template_name, message_text, copied_by, current_timestamp()))
    conn.commit()


def read_df(conn, query: str, params: tuple = ()) -> pd.DataFrame:
    return pd.read_sql_query(query, conn, params=params)


def get_latest_batch_id(conn, import_type: str) -> Optional[int]:
    row = conn.execute('''SELECT batch_id FROM import_batches
        WHERE import_type = ? AND status = 'success'
        ORDER BY batch_id DESC LIMIT 1''', (import_type,)).fetchone()
    return int(row["batch_id"]) if row else None
