import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import re

# ==========================================
# KHỐI 1: CẤU HÌNH HỆ THỐNG
# ==========================================
DB_PATH = "crm_data.db"

def tao_mau_tin_nhan(ky_cuoc):
    return f"""VNPT Long Thành- Nhơn Trạch thông báo:
Đã có thông báo cước kỳ cước tháng {ky_cuoc}. Anh/chị truy cập trang https://vnptdongnai.vn/ để lấy thông báo cước và hóa đơn của công ty.
Anh/chị vui lòng thanh toán cước trước ngày 15. Sau ngày 15 hệ thống sẽ tự động đưa lưới khóa khi ghi nhận còn tồn nợ.
Liên hệ nhân viên kinh doanh: Thuận - 0837892579 để được hỗ trợ gạch nợ sớm nhất sau khi thanh toán.
Nếu đã thanh toán cước, anh/chị vui lòng bỏ qua tin nhắn trên.
Trân trọng."""

# ==========================================
# KHỐI 2: CƠ SỞ DỮ LIỆU (DATABASE)
# ==========================================
def init_db():
    conn = sqlite3.connect(DB_PATH)
    # Bảng khách hàng & Hóa đơn
    conn.execute('''CREATE TABLE IF NOT EXISTS danh_sach_no (
        ma_tt TEXT PRIMARY KEY,
        ten_khach_hang TEXT,
        dia_chi TEXT,
        so_dien_thoai TEXT,
        tien_no REAL,
        trang_thai TEXT DEFAULT 'Chưa liên hệ'
    )''')
    conn.commit()
    return conn

def update_database(conn, df):
    # Đẩy dữ liệu file excel vào SQL
    for _, row in df.iterrows():
        conn.execute('''
            INSERT INTO danh_sach_no (ma_tt, ten_khach_hang, dia_chi, so_dien_thoai, tien_no, trang_thai)
            VALUES (?, ?, ?, ?, ?, 'Chưa liên hệ')
            ON CONFLICT(ma_tt) DO UPDATE SET
            tien_no=excluded.tien_no
        ''', (str(row['ma_tt']), str(row['ten_khach_hang']), str(row['dia_chi']), str(row['so_dien_thoai']), float(row['tien_no'])))
    conn.commit()

def mark_as_paid(conn, df_paid):
    # Đánh dấu đã đóng cho các mã TT có trong file
    for _, row in df_paid.iterrows():
        conn.execute("UPDATE danh_sach_no SET trang_thai = 'Đã đóng tiền' WHERE ma_tt = ?", (str(row['ma_tt']),))
    conn.commit()

# ==========================================
# KHỐI 3: ETL (LÀM SẠCH DỮ LIỆU)
# ==========================================
def clean_data(df, file_type):
    # Chuẩn hóa tên cột để tránh lỗi viết hoa/thường hay có dấu cách
    df.columns = df.columns.str.strip().str.lower()
    
    # Mapping cột tùy theo file
    if file_type == 'tn08':
        col_map = {
            'ma_tt': 'ma_tt', 'mã thanh toán': 'ma_tt',
            'tên khách hàng': 'ten_khach_hang', 'tên thanh toán': 'ten_khach_hang',
            'total nợ thu vét': 'tien_no', 'tiền nợ': 'tien_no',
            'địa chỉ kh': 'dia_chi', 'địa chỉ thanh toán': 'dia_chi',
            'số dt liên hệ': 'so_dien_thoai'
        }
    else:
        col_map = {'ma_tt': 'ma_tt', 'mã thanh toán': 'ma_tt'}

    df = df.rename(columns=col_map)
    
    # Lọc bỏ các cột không cần thiết
    if 'ma_tt' in df.columns:
        df['ma_tt'] = df['ma_tt'].astype(str).str.strip().str.upper()
    
    if file_type == 'tn08':
        if 'tien_no' in df.columns:
            df['tien_no'] = pd.to_numeric(df['tien_no'], errors='coerce').fillna(0)
        
        # Tạo các cột còn thiếu nếu file nguồn không có
        for req_col in ['ten_khach_hang', 'dia_chi', 'so_dien_thoai', 'tien_no']:
            if req_col not in df.columns:
                df[req_col] = ""
                
        return df[['ma_tt', 'ten_khach_hang', 'dia_chi', 'so_dien_thoai', 'tien_no']].dropna(subset=['ma_tt'])
    else:
        return df[['ma_tt']].dropna(subset=['ma_tt'])

# ==========================================
# KHỐI 4: GIAO DIỆN (UI - STREAMLIT)
# ==========================================
st.set_page_config(page_title="CRM Thu Cước KHDN", layout="wide")
conn = init_db()

st.title("📞 Hệ thống Quản lý Thu Cước KHDN")

tab1, tab2, tab3 = st.tabs(["1. Nạp Dữ Liệu", "2. Bàn Làm Việc", "3. Báo Cáo KPI"])

# --- TAB 1: NẠP DỮ LIỆU ---
with tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1. Nạp File Nợ Cước (TN08)")
        file_tn08 = st.file_uploader("Kéo thả file TN08 vào đây", type=['xlsx', 'xls', 'csv'], key="tn08")
        if file_tn08:
            df_raw = pd.read_csv(file_tn08) if file_tn08.name.endswith('.csv') else pd.read_excel(file_tn08)
            df_clean = clean_data(df_raw, 'tn08')
            update_database(conn, df_clean)
            st.success(f"Đã cập nhật {len(df_clean)} khách hàng đang nợ vào hệ thống!")

    with col2:
        st.subheader("2. Nạp File Khách Đã Đóng")
        file_paid = st.file_uploader("Kéo thả file Đã đóng tiền vào đây", type=['xlsx', 'xls', 'csv'], key="paid")
        if file_paid:
            df_paid_raw = pd.read_csv(file_paid) if file_paid.name.endswith('.csv') else pd.read_excel(file_paid)
            df_paid_clean = clean_data(df_paid_raw, 'paid')
            mark_as_paid(conn, df_paid_clean)
            st.success(f"Đã quét và gạch nợ cho {len(df_paid_clean)} khách hàng!")

# --- TAB 2: BÀN LÀM VIỆC ---
with tab2:
    st.subheader("Danh sách khách hàng cần xử lý")
    
    # Đọc dữ liệu từ SQL
    df_sql = pd.read_sql_query("SELECT * FROM danh_sach_no WHERE trang_thai != 'Đã đóng tiền'", conn)
    
    if not df_sql.empty:
        # Hiển thị bảng dữ liệu với màu sắc
        def color_status(val):
            color = 'red' if val == 'Chưa liên hệ' else 'orange' if val == 'Đã gọi - Hẹn đóng' else 'green'
            return f'color: {color}; font-weight: bold'
            
        st.dataframe(df_sql.style.map(color_status, subset=['trang_thai']), use_container_width=True)
        
        st.divider()
        st.markdown("### 💬 Trợ lý Zalo & Cập nhật trạng thái")
        
        c1, c2 = st.columns([1, 2])
        with c1:
            kh_chon = st.selectbox("Chọn Mã TT để xử lý:", df_sql['ma_tt'].tolist())
            trang_thai_moi = st.selectbox("Cập nhật kết quả:", ["Đã gọi - Hẹn đóng", "Sai số ĐT", "Khách không nghe máy"])
            if st.button("Lưu cập nhật"):
                conn.execute("UPDATE danh_sach_no SET trang_thai = ? WHERE ma_tt = ?", (trang_thai_moi, kh_chon))
                conn.commit()
                st.rerun() # Tải lại trang để nhảy màu ngay lập tức
                
        with c2:
            ky_cuoc = st.text_input("Nhập kỳ cước (VD: 05/2026):", value=datetime.now().strftime("%m/%Y"))
            st.markdown("**Mẫu tin nhắn tự động (Bấm icon góc phải để Copy):**")
            st.code(tao_mau_tin_nhan(ky_cuoc), language="text")
    else:
        st.info("Tuyệt vời! Tuyến của bạn hiện tại không còn ai nợ cước cần gọi.")

# --- TAB 3: BÁO CÁO ---
with tab3:
    st.subheader("Tổng quan năng suất")
    df_report = pd.read_sql_query("SELECT * FROM danh_sach_no", conn)
    
    if not df_report.empty:
        tong_tien_no = df_report[df_report['trang_thai'] != 'Đã đóng tiền']['tien_no'].sum()
        da_dong = len(df_report[df_report['trang_thai'] == 'Đã đóng tiền'])
        chua_goi = len(df_report[df_report['trang_thai'] == 'Chưa liên hệ'])
        
        metric1, metric2, metric3 = st.columns(3)
        metric1.metric("Tổng tiền cần thu hồi", f"{tong_tien_no:,.0f} VNĐ")
        metric2.metric("Số khách đã gạch nợ", da_dong)
        metric3.metric("Số khách CHƯA LIÊN HỆ", chua_goi)
