from __future__ import annotations
from datetime import datetime
from dateutil.relativedelta import relativedelta
from typing import Dict, Any
import pandas as pd
import sqlite3
from .db import read_df, get_latest_batch_id
from .utils import format_money


def default_billing_period(config: Dict[str, Any]) -> str:
    mode = config.get("billing", {}).get("default_period_mode", "previous_month")
    fmt = config.get("billing", {}).get("display_format", "%m/%Y")
    now = datetime.now()
    if mode == "previous_month":
        now = now - relativedelta(months=1)
    return now.strftime(fmt)


def get_operational_view(conn: sqlite3.Connection, config: Dict[str, Any]) -> pd.DataFrame:
    latest_tn08 = get_latest_batch_id(conn, "tn08")
    if not latest_tn08:
        return pd.DataFrame()
    query = '''
    WITH latest_debt AS (
        SELECT ma_tt, MAX(billing_period) AS billing_period, SUM(debt_amount) AS debt_amount,
               MAX(source_customer_name) AS source_customer_name, MAX(source_address) AS source_address
        FROM debt_snapshots WHERE batch_id = ? GROUP BY ma_tt
    ),
    paid AS (
        SELECT ma_tt, SUM(paid_amount) AS paid_amount, MAX(created_at) AS last_paid_import_at
        FROM paid_updates GROUP BY ma_tt
    ),
    last_interaction AS (
        SELECT i.* FROM interactions i
        INNER JOIN (SELECT ma_tt, MAX(created_at) AS max_created_at FROM interactions GROUP BY ma_tt) x
        ON i.ma_tt = x.ma_tt AND i.created_at = x.max_created_at
    ),
    latest_contact AS (
        SELECT c.* FROM customer_contacts c
        INNER JOIN (SELECT ma_tt, MAX(created_at) AS max_created_at FROM customer_contacts GROUP BY ma_tt) x
        ON c.ma_tt = x.ma_tt AND c.created_at = x.max_created_at
    )
    SELECT c.ma_tt, c.representative_code,
           COALESCE(c.customer_name, d.source_customer_name, '') AS customer_name,
           COALESCE(c.address, d.source_address, '') AS address,
           COALESCE(lc.contact_value, c.current_phone, '') AS phone,
           COALESCE(c.assigned_staff, '') AS assigned_staff,
           d.billing_period, COALESCE(d.debt_amount, 0) AS debt_amount,
           COALESCE(p.paid_amount, 0) AS paid_amount,
           li.result AS last_result, li.note AS last_note,
           li.promised_payment_date, li.created_at AS last_contacted_at
    FROM latest_debt d
    LEFT JOIN customers c ON c.ma_tt = d.ma_tt
    LEFT JOIN paid p ON p.ma_tt = d.ma_tt
    LEFT JOIN last_interaction li ON li.ma_tt = d.ma_tt
    LEFT JOIN latest_contact lc ON lc.ma_tt = d.ma_tt
    ORDER BY d.debt_amount DESC
    '''
    df = read_df(conn, query, (latest_tn08,))
    if df.empty:
        return df
    labels = config.get("status_rules", {}).get("labels", {})
    colors = config.get("status_rules", {}).get("colors", {})
    def compute_status(row):
        if float(row.get("paid_amount") or 0) > 0:
            return "paid"
        if row.get("last_contacted_at"):
            return "contacted_unpaid"
        return "unpaid_uncontacted"
    df["status_code"] = df.apply(compute_status, axis=1)
    df["status_label"] = df["status_code"].map(labels).fillna(df["status_code"])
    df["status_color"] = df["status_code"].map(colors).fillna("")
    df["debt_amount_display"] = df["debt_amount"].apply(format_money)
    return df


def render_message(template_name: str, row: dict, config: Dict[str, Any]) -> str:
    templates = config.get("message_templates", {})
    if template_name not in templates:
        raise ValueError(f"Không tìm thấy mẫu tin nhắn: {template_name}")
    app_cfg = config.get("app", {})
    context = {
        "company_unit": app_cfg.get("company_unit", ""),
        "website": app_cfg.get("default_website", ""),
        "payment_deadline": app_cfg.get("default_payment_deadline_text", ""),
        "staff_name": row.get("assigned_staff") or app_cfg.get("default_staff_name", ""),
        "staff_phone": app_cfg.get("default_staff_phone", ""),
        "billing_period": row.get("billing_period") or default_billing_period(config),
        "customer_name": row.get("customer_name", ""),
        "ma_tt": row.get("ma_tt", ""),
        "debt_amount": format_money(row.get("debt_amount", 0)),
    }
    return templates[template_name].format(**context)


def summarize(df: pd.DataFrame) -> Dict[str, Any]:
    if df.empty:
        return {"total_ma_tt":0, "total_debt":0.0, "paid_count":0, "contacted_unpaid_count":0, "unpaid_uncontacted_count":0}
    return {
        "total_ma_tt": int(len(df)),
        "total_debt": float(df["debt_amount"].sum()),
        "paid_count": int((df["status_code"] == "paid").sum()),
        "contacted_unpaid_count": int((df["status_code"] == "contacted_unpaid").sum()),
        "unpaid_uncontacted_count": int((df["status_code"] == "unpaid_uncontacted").sum()),
    }
