from __future__ import annotations
import streamlit as st
from src.config import load_config
from src.db import get_connection, init_db
from src.ui import render_upload_page, render_customer_workbench, render_report_page, render_settings_page

st.set_page_config(page_title="Mini-CRM Thu Cước KHDN", page_icon="📞", layout="wide")

@st.cache_resource
def bootstrap():
    config = load_config("config/settings.yaml")
    conn = get_connection(config["database"]["path"])
    init_db(conn)
    return config, conn

config, conn = bootstrap()
st.title(config["app"]["title"])
st.caption("Kiến trúc: Configuration → Data/SQLite → ETL Validation → Business Logic → Streamlit UI")

tabs = st.tabs(["1. Cập nhật dữ liệu", "2. Khách cần xử lý", "3. Báo cáo", "4. Cấu hình"])
with tabs[0]:
    render_upload_page(conn, config)
with tabs[1]:
    render_customer_workbench(conn, config)
with tabs[2]:
    render_report_page(conn, config)
with tabs[3]:
    render_settings_page(config)
