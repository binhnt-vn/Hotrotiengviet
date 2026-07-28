# 📝 Word Web Add-in: Tiện Ích Xử Lý Tiếng Việt

Bộ công cụ **Word Web Add-in** hỗ trợ 3 nhóm tính năng xử lý văn bản tiếng Việt tối ưu cho Microsoft Word 365 (Desktop & Web):

---

## 🌟 Tính Năng Nổi Bật

### 1. 📅 Rút Ngắn Định Dạng Ngày Tháng
Tự động tìm kiếm và quy đổi định dạng ngày tháng dài sang dạng viết tắt gọn gàng (hỗ trợ cả các số có số `0` ở đầu):
* `ngày 20 tháng 07 năm 2026` ➔ `20/7/2026`
* `ngày 20 tháng 07` ➔ `ngày 20/7`
* `tháng 07 năm 2026` ➔ `T7/2026`
* `tháng 07` ➔ `T7`

### 2. 🔤 Quản Lý & Thay Thế Từ Viết Tắt
* **Quản lý linh hoạt:** Thêm mới, **Chỉnh sửa trực tiếp (Inline Edit)**, Xóa quy tắc viết tắt.
* **Xuất & Nhập từ điển (Export / Import):** Dễ dàng xuất danh sách từ điển ra tệp `tu_dien_viet_tat.json` hoặc nhập từ điển có sẵn để chia sẻ giữa các máy tính.
* **Tự động sắp xếp A-Z:** Danh sách hiển thị tự động sắp xếp theo bảng chữ cái tiếng Việt giúp dễ dàng tra cứu.
* **Thay thế thông minh:**
  * Tìm kiếm không phân biệt chữ hoa/thường (`thành phố` hay `THÀNH PHỐ` đều khớp).
  * Giữ nguyên chữ hoa/thường bản gốc của từ viết tắt (VD: luôn ra `TP.HCM`, `CHXHCNVN`).
  * Tự động ưu tiên thay thế cụm từ dài trước cụm từ ngắn để tránh lỗi hỏng văn bản.

### 3. 🔍 Rà Soát & Sửa Lỗi Chính Tả Tiếng Việt
* Phát hiện các lỗi chính tả phổ biến (dấu hỏi/ngã, phụ âm x/s, ch/tr, từ ghép sai).
* Hỗ trợ **sửa từng lỗi** hoặc **sửa tất cả lỗi** với 1 cú nhấp chuột.

---

## 📁 Cấu Trúc Thư Mục Dự Án

```
Word Add-in/
├── manifest.xml           ← File khai báo Add-in (dùng để sideload vào Word)
├── app_standalone.html    ← Giao diện chính (Standalone single-file web app)
├── app.js                 ← Mã nguồn xử lý JavaScript chính
├── index.html             ← Trang HTML ứng dụng
├── style.css              ← Giao diện CSS
├── commands.html          ← Trang khởi tạo Office.js Commands
├── package.json           ← Cấu hình dự án npm & scripts
├── .gitignore             ← Cấu hình loại bỏ node_modules và tệp tạm
├── assets/
│   ├── icon-16.png        ← Icon 16x16 cho thanh Ribbon
│   ├── icon-32.png        ← Icon 32x32 cho thanh Ribbon
│   └── icon-80.png        ← Icon 80x80 cho thanh Ribbon
└── README.md              ← Hướng dẫn sử dụng
```

---

## 🚀 CÁCH TRIỂN KHAI VÀ CÀI ĐẶT

### 1. Đưa Dự Án Lên GitHub (Hosting)

1. Tải toàn bộ mã nguồn dự án lên GitHub Repository của bạn.
2. Bật tính năng **GitHub Pages** (Khuyên dùng):
   * Vào Repo của bạn trên GitHub ➔ **Settings** ➔ **Pages**.
   * Tại **Build and deployment**, chọn Branch `main` ➔ nhấn **Save**.
   * Bạn sẽ nhận được đường dẫn dạng: `https://<your-username>.github.io/<your-repo>/app_standalone.html`.
3. Cập nhật đường dẫn URL trên vào tệp `manifest.xml` tại ô `<SourceLocation>` và `<bt:Url id="Taskpane.Url">`.

---

### 2. Sideload Vào Word 365 Desktop (Windows)

#### Cách 1: Sử dụng Thư Mục Chia Sẻ (Shared Folder Catalog) — *Khuyên dùng trên Windows*
1. Chuột phải vào thư mục chứa dự án trên máy (hoặc thư mục trên ổ mạng) ➔ Chọn **Properties** ➔ **Sharing** ➔ **Share...** ➔ Chọn tài khoản của bạn hoặc `Everyone` ➔ Nhấn **Share**.
2. Sao chép đường dẫn mạng UNC (dạng: `\\localhost\Word Add-in` hoặc `\\ServerName\ShareName`).
3. Mở Word Desktop ➔ Vào **File** ➔ **Options** ➔ **Trust Center** ➔ **Trust Center Settings...** ➔ **Trusted Add-in Catalogs**.
4. Dán đường dẫn UNC vào ô **Catalog URL** ➔ Nhấn **Add catalog** ➔ Tích chọn ô **Show in Menu**.
5. Nhấn **OK** và khởi động lại ứng dụng Word.
6. Vào tab **Insert** (Chèn) ➔ **Get Add-ins** (hoặc **My Add-ins**) ➔ Chuyển sang tab **SHARED FOLDER** ➔ Chọn Add-in và nhấn **Add**.

---

### 3. Sideload Vào Word Online (Trình duyệt Web)

1. Truy cập [Word Online](https://word.office.com) và mở một văn bản.
2. Vào **Insert** ➔ **Add-ins** ➔ **My Add-ins**.
3. Bấm **Upload My Add-in** ở góc trên bên phải.
4. Chọn tệp `manifest.xml` từ máy tính của bạn và chọn **Upload**.

---

## 🛠️ Hướng Dẫn Dành Cho Lập Trình Viên (Development)

Tải thư viện phụ thuộc và chạy chế độ thử nghiệm:

```bash
# Cài đặt thư viện dev
npm install

# Khởi chạy server HTTPS thử nghiệm tại máy cục bộ (localhost:3000)
npm run serve
```

---

## 📄 Giấy Phép
Dự án được phát hành theo giấy phép **MIT License**.
