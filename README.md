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
3. **Rà soát chính tả Tiếng Việt**:
   - Nhận diện các lỗi chính tả phổ biến (dấu hỏi/ngã, phụ âm x/s, ch/tr, từ ghép sai như *chuẩn đoán* ➔ *chẩn đoán*, *thăm quan* ➔ *tham quan*...).

---

## 🌟 CÁCH ĐẮC ĐỊA NHẤT: Thử Nghiệm Ngay Trên Word Online (Web)

Do chính sách máy tính cơ quan/công ty khóa cả Registry lẫn quyền Chia sẻ (Share) trên Word Desktop, cách **nhanh nhất 100% thành công** là thử nghiệm qua **Word Online (Trình duyệt)**:

### Các bước thực hiện:

1. **Mở Word Online**:
   Truy cập [https://word.office.com](https://word.office.com) (Đăng nhập tài khoản Microsoft cá nhân hoặc công ty).

2. **Mở một văn bản bất kỳ**:
   Tạo mới một Tài liệu trống (Blank document).

3. **Chèn Add-in**:
   - Nhấp tab **Insert** (Chèn) trên thanh công cụ.
   - Nhấp chọn nút **Add-ins** (Tệp bổ sung) ➔ Chọn **More Add-ins** (Add-in khác).
   - Ở góc trên cùng bên phải cửa sổ pop-up, bấm nút **Upload My Add-in** (Tải lên Add-in của tôi).
   - Chọn đường dẫn file `manifest.xml` tại:
     `C:\Users\binhnt.MARKET_CLEARING\Dropbox\NGHIEN CUU\Word Add-in\manifest.xml`

4. **Sử dụng**:
   Tab **"Xử lý Tiếng Việt"** và nút **"Mở Tiện Ích"** sẽ xuất hiện ngay lập tức trên thanh Ribbon!


