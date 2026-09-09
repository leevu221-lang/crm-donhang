# Google Sheets Customer Data Entry Form (GitHub Pages & Google Apps Script)

Form nhập liệu khách hàng trực tuyến chạy trên **GitHub Pages** (hoặc trực tiếp trong **Google Sheets**), tự động ghi nhận STT và ngày giờ lưu (GMT+7).

## 🌐 Link truy cập trực tiếp trên GitHub Pages
Sau khi bật GitHub Pages, link truy cập trực tuyến của bạn sẽ là:
👉 **`https://leevu221-lang.github.io/crm-donhang/`**

---

## 🚀 Hướng dẫn kết nối 2 phút:

### Bước 1: Bật GitHub Pages (để có link web)
1. Mở repository: [https://github.com/leevu221-lang/crm-donhang](https://github.com/leevu221-lang/crm-donhang)
2. Nhấn vào **Settings** (Cài đặt) trên menu GitHub.
3. Chọn mục **Pages** ở cột menu bên trái.
4. Tại phần **Build and deployment** ➔ **Branch**:
   - Chọn nhánh **`main`**
   - Thư mục: **`/ (root)`**
   - Nhấn **Save**.
5. Đợi 1 phút, GitHub sẽ cấp link trang web: **`https://leevu221-lang.github.io/crm-donhang/`**.

---

### Bước 2: Cập nhật mã nguồn trên Google Sheets & Lấy link Web App
1. Mở Google Sheet của bạn ➔ Chọn **Tiện ích mở rộng** ➔ **Apps Script**.
2. Sao chép toàn bộ mã trong file [`Code_AllInOne.gs`](./Code_AllInOne.gs) và dán đè vào `Code.gs`.
3. Nhấn nút **Lưu** (Ctrl + S).
4. Ở góc trên bên phải màn hình Apps Script, nhấn nút xanh **Triển khai (Deploy)** ➔ **Triển khai mới (New deployment)**:
   - Nhấp vào biểu tượng bánh răng ⚙️ ➔ Chọn **Ứng dụng web (Web app)**.
   - **Mô tả**: `v1`
   - **Thực thi dưới dạng**: `Tôi (Email của bạn)`
   - **Ai có quyền truy cập**: Chọn **Bất kỳ ai (Anyone)** *(Quan trọng: để web từ GitHub gửi dữ liệu vào được)*.
   - Nhấn nút **Triển khai (Deploy)** ➔ Cấp quyền ủy quyền cho tài khoản Google của bạn.
5. Sao chép đường dẫn **URL ứng dụng web** (dạng `https://script.google.com/macros/s/.../exec`).

---

### Bước 3: Dán link kết nối vào Form trên GitHub Pages
1. Mở link web GitHub: `https://leevu221-lang.github.io/crm-donhang/`
2. Bấm vào biểu tượng bánh răng **⚙️** ở góc trên bên phải form.
3. Dán link Web App vừa sao chép ở Bước 2 vào ô và bấm **Lưu Cấu Hình**.
4. Xong! Từ nay bất kỳ ai nhập dữ liệu trên link GitHub này, dữ liệu sẽ ngay lập tức được chuyển thẳng vào Google Sheet của bạn kèm ngày giờ chính xác!
