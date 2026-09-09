# Google Sheets Customer Data Entry Form (Google Apps Script)

Ứng dụng Form nhập liệu thông tin khách hàng trực tiếp vào Google Sheets có tự động ghi nhận thời gian lưu (Timestamp) và STT.

## 📌 Các trường thông tin
1. **STT**: Tự động tăng dần theo số dòng tiếp theo.
2. **Khách hàng**: Họ tên khách hàng.
3. **SĐT**: Số điện thoại (tự động giữ số 0 ở đầu).
4. **Sản phẩm khách xem**: Danh sách / ghi chú sản phẩm khách quan tâm.
5. **Nhân viên**: Tên nhân viên tư vấn / tiếp khách.
6. **Ngày giờ lưu**: Tự động ghi nhận thời gian lúc bấm Lưu theo múi giờ Việt Nam (`dd/MM/yyyy HH:mm:ss`).

## 📁 Cấu trúc thư mục
- `Code_AllInOne.gs`: Bản tích hợp tất cả trong 1 file duy nhất (khuyên dùng để không bao giờ bị lỗi thiếu file HTML).
- `Code.gs`: Mã backend Google Apps Script.
- `Index.html`: Giao diện Form hiện đại (nếu muốn tách riêng HTML).

## 🚀 Cách sử dụng
1. Mở trang Google Sheet của bạn.
2. Chọn **Tiện ích mở rộng** > **Apps Script**.
3. Dán nội dung trong tệp `Code_AllInOne.gs` vào `Code.gs`.
4. Nhấn **Lưu** (Ctrl+S) rồi F5 lại trang tính để thấy menu **📋 Quản Lý Nhập Liệu**.
