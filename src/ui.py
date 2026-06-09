from __future__ import annotations
import io
import sqlite3
from typing import Dict, Any
import pandas as pd
import streamlit as st
from .etl import read_excel_flexible, standardize_dataframe, validate_standardized, ImportValidationError
from .db import create_batch, upsert_customers, insert_debt_snapshots, insert_paid_updates, add_interaction, add_contact, add_sent_message, read_df
from .services import get_operational_view, render_message, summarize, default_billing_period
from .utils import format_money, normalize_phone


def import_file_widget(conn: sqlite3.Connection, config: Dict[str, Any], import_type: str, billing_period: str | None = None):
    import_cfg = config["imports"][import_type]
    uploaded = st.file_uploader(f"Upload {import_cfg.get('label', import_type)}", type=["xlsx", "xls"], key=f"upload_{import_type}")
    if not uploaded:
        return
    try:
        raw_df, sheet_name = read_excel_flexible(uploaded, import_cfg)
        std_df = standardize_dataframe(raw_df, import_cfg, import_type)
        quality = validate_standardized(std_df, import_type)
        st.info(f"Đọc được {quality['row_count']} dòng. Tổng tiền nhận diện: {format_money(quality['total_amount'])}.")
        for warning in quality["warnings"]:
            st.warning(warning)
        with st.expander("Xem trước dữ liệu sau chuẩn hóa"):
            st.dataframe(std_df.head(30), use_container_width=True)
        if st.button(f"Xác nhận import {import_cfg.get('label', import_type)}", key=f"confirm_{import_type}"):
            batch_id = create_batch(conn, import_type, uploaded.name, sheet_name, quality["row_count"], quality["total_amount"])
            if import_type in {"tn08", "assignment"}:
                upsert_customers(conn, std_df.to_dict("records"))
            if import_type == "tn08":
                insert_debt_snapshots(conn, batch_id, std_df.to_dict("records"), billing_period or default_billing_period(config))
            if import_type == "paid":
                insert_paid_updates(conn, batch_id, std_df.to_dict("records"))
            st.success(f"Import thành công. Batch ID: {batch_id}")
    except ImportValidationError as exc:
        st.error(str(exc))
    except Exception as exc:
        st.error(f"Lỗi không xác định khi import: {exc}")


def render_upload_page(conn: sqlite3.Connection, config: Dict[str, Any]):
    st.subheader("1. Cập nhật dữ liệu")
    st.write("Dữ liệu Excel được nạp vào SQLite theo từng batch, không ghi đè lịch sử. Tên cột được nhận diện theo `config/settings.yaml`.")
    billing_period = st.text_input("Kỳ cước áp dụng cho file TN08", value=default_billing_period(config))
    col1, col2 = st.columns(2)
    with col1:
        import_file_widget(conn, config, "tn08", billing_period=billing_period)
        import_file_widget(conn, config, "paid")
    with col2:
        import_file_widget(conn, config, "assignment")
    st.divider()
    st.subheader("Lịch sử import gần nhất")
    batches = read_df(conn, """
        SELECT batch_id, import_type, file_name, sheet_name, row_count, total_amount, status, imported_at
        FROM import_batches ORDER BY batch_id DESC LIMIT 20
    """)
    st.dataframe(batches, use_container_width=True)


def render_customer_workbench(conn: sqlite3.Connection, config: Dict[str, Any]):
    st.subheader("2. Khách cần xử lý")
    df = get_operational_view(conn, config)
    if df.empty:
        st.warning("Chưa có dữ liệu TN08. Hãy import TN08 trước.")
        return
    summary = summarize(df)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Tổng mã theo dõi", summary["total_ma_tt"])
    c2.metric("Tổng tiền còn nợ", format_money(summary["total_debt"]))
    c3.metric("Đã liên hệ chưa đóng", summary["contacted_unpaid_count"])
    c4.metric("Chưa liên hệ / thiếu số", summary["unpaid_uncontacted_count"])
    status_options = ["Tất cả"] + list(df["status_label"].dropna().unique())
    selected_status = st.selectbox("Lọc trạng thái", status_options)
    view_df = df.copy()
    if selected_status != "Tất cả":
        view_df = view_df[view_df["status_label"] == selected_status]
    search = st.text_input("Tìm theo MA_TT / tên khách / số điện thoại")
    if search:
        s = search.lower()
        view_df = view_df[view_df["ma_tt"].str.lower().str.contains(s, na=False) | view_df["customer_name"].str.lower().str.contains(s, na=False) | view_df["phone"].str.lower().str.contains(s, na=False)]
    display_cols = ["status_color", "status_label", "ma_tt", "customer_name", "phone", "billing_period", "debt_amount_display", "last_result", "promised_payment_date", "last_contacted_at"]
    st.dataframe(view_df[display_cols], use_container_width=True, height=360)
    if view_df.empty:
        st.info("Không có khách phù hợp bộ lọc.")
        return
    ma_tt_selected = st.selectbox("Chọn MA_TT để thao tác", view_df["ma_tt"].tolist())
    row = df[df["ma_tt"] == ma_tt_selected].iloc[0].to_dict()
    st.markdown("### Chi tiết khách hàng")
    col_info, col_action = st.columns([1, 1])
    with col_info:
        st.write(f"**MA_TT:** {row.get('ma_tt', '')}")
        st.write(f"**Tên khách:** {row.get('customer_name', '')}")
        st.write(f"**Số điện thoại/Zalo ưu tiên:** {row.get('phone', '') or 'Chưa có'}")
        st.write(f"**Kỳ cước:** {row.get('billing_period', '')}")
        st.write(f"**Tiền còn nợ:** {format_money(row.get('debt_amount', 0))}")
        st.write(f"**Trạng thái:** {row.get('status_color', '')} {row.get('status_label', '')}")
        st.markdown("#### Lịch sử tương tác")
        history = read_df(conn, """
            SELECT result, note, promised_payment_date, created_by, created_at
            FROM interactions WHERE ma_tt = ? ORDER BY created_at DESC LIMIT 10
        """, (ma_tt_selected,))
        st.dataframe(history, use_container_width=True)
    with col_action:
        st.markdown("#### Cập nhật kết quả cuộc gọi")
        result = st.selectbox("Kết quả", config.get("contact_results", []))
        promised_date = st.date_input("Ngày hẹn thanh toán", value=None)
        note = st.text_area("Ghi chú cuộc gọi")
        created_by = st.text_input("Người cập nhật", value=config["app"].get("default_staff_name", ""))
        if st.button("Lưu lịch sử gọi"):
            add_interaction(conn, ma_tt_selected, result, note, str(promised_date) if promised_date else "", created_by)
            st.success("Đã lưu lịch sử gọi.")
        st.markdown("#### Cập nhật số Zalo / số liên hệ mới")
        new_contact = st.text_input("Số Zalo / SĐT mới")
        contact_person_name = st.text_input("Tên người phụ trách mới")
        role = st.text_input("Vai trò", value="Kế toán")
        contact_note = st.text_area("Ghi chú số liên hệ", key="contact_note")
        if st.button("Lưu số liên hệ mới"):
            normalized_contact = normalize_phone(new_contact)
            if not normalized_contact:
                st.error("Số liên hệ chưa hợp lệ.")
            else:
                add_contact(conn, ma_tt_selected, normalized_contact, "zalo", contact_person_name, role, contact_note)
                st.success("Đã lưu số liên hệ mới và cập nhật làm số ưu tiên.")
    st.divider()
    st.markdown("### Tạo mẫu tin nhắn Zalo để copy")
    template_names = list(config.get("message_templates", {}).keys())
    template_name = st.selectbox("Chọn mẫu tin nhắn", template_names)
    message = render_message(template_name, row, config)
    st.code(message, language="text")
    if st.button("Ghi nhận đã copy/gửi mẫu này"):
        add_sent_message(conn, ma_tt_selected, template_name, message, config["app"].get("default_staff_name", ""))
        st.success("Đã ghi nhận lịch sử mẫu tin nhắn.")


def render_report_page(conn: sqlite3.Connection, config: Dict[str, Any]):
    st.subheader("3. Báo cáo")
    df = get_operational_view(conn, config)
    if df.empty:
        st.warning("Chưa có dữ liệu để báo cáo.")
        return
    summary = summarize(df)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Tổng MA_TT", summary["total_ma_tt"])
    c2.metric("Tổng tiền còn nợ", format_money(summary["total_debt"]))
    c3.metric("Đã đóng", summary["paid_count"])
    c4.metric("Chưa liên hệ / thiếu số", summary["unpaid_uncontacted_count"])
    grouped = df.groupby(["representative_code", "customer_name"], dropna=False).agg(
        so_ma_tt=("ma_tt", "count"), tong_no=("debt_amount", "sum"), so_da_lien_he=("last_contacted_at", lambda s: s.notna().sum())
    ).reset_index().sort_values("tong_no", ascending=False)
    grouped["tong_no_hien_thi"] = grouped["tong_no"].apply(format_money)
    st.markdown("### Gom nhóm theo công ty / mã đại diện")
    st.dataframe(grouped, use_container_width=True)
    st.markdown("### Danh sách chi tiết")
    st.dataframe(df, use_container_width=True)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Chi tiet")
        grouped.to_excel(writer, index=False, sheet_name="Gom nhom")
    output.seek(0)
    st.download_button("Tải báo cáo Excel", data=output, file_name="bao_cao_thu_cuoc_khdn.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def render_settings_page(config: Dict[str, Any]):
    st.subheader("4. Cấu hình đang dùng")
    st.write("Tên cột, mẫu tin nhắn, nhân viên, hotline, trạng thái đều nằm trong file cấu hình.")
    st.json(config)
