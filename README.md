# Hướng Dẫn Cài Đặt & Sử Dụng Word Add-In (Xử Lý Tiếng Việt)

Bộ công cụ **Word Web Add-in** hỗ trợ 3 tính năng quan trọng:
1. **Rút ngắn định dạng ngày tháng tiếng Việt**:
   - `ngày 20 tháng 7 năm 2026` ➔ `20/7/2026`
   - `ngày 20 tháng 7` ➔ `ngày 20/7`
   - `tháng 7 năm 2026` ➔ `T7/2026`
   - `tháng 7` ➔ `T7`
2. **Thay thế từ khóa thành từ viết tắt** (Rút gọn văn bản):
   - Cho phép lưu và quản lý bảng rule viết tắt (VD: `Cộng hòa Xã hội Chủ nghĩa Việt Nam` ➔ `CHXHCNVN`, `Thành phố` ➔ `TP.`).
   - Hỗ trợ thay thế toàn bộ hoặc vùng chọn.
   - Lưu rule vào localStorage (không mất khi đóng Word).
3. **Rà soát chính tả Tiếng Việt**:
   - Nhận diện các lỗi chính tả phổ biến (dấu hỏi/ngã, phụ âm x/s, ch/tr, từ ghép sai).
   - Hỗ trợ **sửa từng lỗi** hoặc **sửa tất cả** với một click.

---

## 📁 Cấu trúc thư mục

```
Word Add-in/
├── manifest.xml           ← File khai báo add-in (upload vào Word)
├── app_standalone.html    ← Giao diện chính (single-file, host trên CDN)
├── app.js                 ← JavaScript cho bản local development
├── index.html             ← HTML cho bản local development
├── style.css              ← CSS cho bản local development
├── commands.html           ← Commands page (Office.js)
├── package.json           ← npm config
├── assets/
│   ├── icon-16.png        ← Icon 16x16 cho Ribbon
│   ├── icon-32.png        ← Icon 32x32 cho Ribbon
│   └── icon-80.png        ← Icon 80x80 cho Ribbon
└── README.md              ← File này
```

---

## 🖥️ CÁCH CÀI ĐẶT TRÊN WORD 365 DESKTOP

### Bước 1: Push toàn bộ file lên GitHub

Đảm bảo repo `binhnt-vn/Hotrotiengviet` đã có đầy đủ các file:
- `app_standalone.html`
- `assets/icon-16.png`
- `assets/icon-32.png`
- `assets/icon-80.png`

```bash
cd "C:\Users\binhnt.MARKET_CLEARING\Dropbox\NGHIEN CUU\Word Add-in"
git add .
git commit -m "Fix add-in: syntax errors, manifest, icons"
git push origin main
```

### Bước 2: Đợi CDN cập nhật (1-2 phút)

File được host qua jsDelivr CDN:
```
https://cdn.jsdelivr.net/gh/binhnt-vn/Hotrotiengviet@main/app_standalone.html
```

> **Lưu ý:** Nếu đã push trước đó, có thể cần purge cache CDN bằng cách truy cập:
> `https://purge.jsdelivr.net/gh/binhnt-vn/Hotrotiengviet@main/app_standalone.html`

### Bước 3: Sideload manifest vào Word 365 Desktop

**Cách 1: Upload trực tiếp (đơn giản nhất)**
1. Mở **Word 365 Desktop**
2. Vào tab **Insert** (Chèn) trên Ribbon
3. Bấm **Get Add-ins** (Nhận Add-in) hoặc **My Add-ins**
4. Chọn tab **My Add-ins** (Add-in của tôi)
5. Bấm **Upload My Add-in** (Tải lên Add-in của tôi)
6. Chọn file `manifest.xml` từ thư mục dự án
7. Bấm **Upload**

**Cách 2: Sideload qua thư mục Wef (nếu Cách 1 không khả dụng)**
1. Mở File Explorer, vào đường dẫn:
   ```
   %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
   ```
   (Nếu thư mục `Wef` chưa có, tạo mới)
2. Copy file `manifest.xml` vào thư mục này
3. Khởi động lại Word

**Cách 3: Dùng Share folder catalog**
1. Tạo một thư mục chia sẻ mạng (VD: `\\localhost\AddIns`)
2. Copy `manifest.xml` vào thư mục đó
3. Trong Word: **File** → **Options** → **Trust Center** → **Trust Center Settings** → **Trusted Add-in Catalogs**
4. Thêm URL thư mục chia sẻ, tick **Show in Menu**
5. Khởi động lại Word

### Bước 4: Sử dụng

Sau khi cài đặt thành công:
- Nút **"Mở Tiện Ích"** sẽ xuất hiện trong nhóm **"Xử lý Tiếng Việt"** trên tab **Home**
- Bấm nút để mở Task Pane bên phải với 3 tab tính năng

---

## 🌐 CÀI ĐẶT TRÊN WORD ONLINE (Web)

1. Truy cập [https://word.office.com](https://word.office.com)
2. Mở hoặc tạo mới một tài liệu
3. Vào **Insert** → **Add-ins** → **Upload My Add-in**
4. Chọn file `manifest.xml`
5. Add-in sẽ mở ngay lập tức

---

## ⚠️ Xử lý sự cố

| Vấn đề | Giải pháp |
|--------|----------|
| Task Pane trắng | Kiểm tra CDN URL đã push file chưa. Thử mở URL trực tiếp trên trình duyệt. |
| Lỗi "manifest invalid" | Đảm bảo manifest.xml đúng định dạng XML. Kiểm tra icon URLs có truy cập được. |
| Không thấy nút trên Ribbon | Thử đóng và mở lại Word. Hoặc dùng Cách 2 (Wef folder). |
| Lỗi JavaScript | Mở DevTools (F12) trong Task Pane để xem Console log. |
| CDN không cập nhật | Purge cache: `https://purge.jsdelivr.net/gh/binhnt-vn/Hotrotiengviet@main/app_standalone.html` |
