# Mini-CRM Thu Cước KHDN

Project Streamlit hỗ trợ theo dõi thu cước Khách hàng Doanh nghiệp theo kiến trúc tách lớp:

```text
config/            Cấu hình mapping cột, trạng thái, mẫu tin nhắn
src/db.py          Database SQLite
src/etl.py         Đọc, chuẩn hóa, validate file Excel
src/services.py    Logic nghiệp vụ: trạng thái, báo cáo, lịch sử gọi
src/ui.py          Thành phần giao diện Streamlit
app.py             Entry point chạy ứng dụng
```

## Cài đặt

```bash
pip install -r requirements.txt
```

## Chạy app

```bash
streamlit run app.py
```

## Quy trình dùng hằng ngày

1. Vào tab **1. Cập nhật dữ liệu**
2. Upload file **TN08 hóa đơn chưa thu**
3. Upload file **DS giao kỳ cước**
4. Nếu có, upload file **DS đã đóng**
5. Vào tab **2. Khách cần xử lý**
6. Chọn khách, gọi điện, cập nhật kết quả, copy mẫu tin nhắn
7. Vào tab **3. Báo cáo** để xem KPI và xuất Excel

## Nguyên tắc thiết kế

- Không hardcode tên cột trong logic xử lý.
- Tên cột được khai báo trong `config/settings.yaml`.
- Mẫu tin nhắn được khai báo trong config.
- Lịch sử gọi và số liên hệ mới được lưu riêng, không ghi đè mất lịch sử.
- Mỗi lần import file được lưu thành một batch để truy vết.
