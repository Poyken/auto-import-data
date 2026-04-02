# PHẦN MỀM TỰ ĐỘNG NẠP DỮ LIỆU ĐO (V6 ULTIMATE)

## 🌟 Tính năng chính:
1. **Đồng bộ tự động:** Quét thư mục ngày hôm nay (`yyyy-MM-dd`) và nạp file Excel mới sinh vào SQL Server.
2. **Ánh xạ thông minh (Smart Mapping):** Tự động nhận diện các cột Excel có tên biến thể như `Equipment Number`, `LotNo`, `Capacity(mAh)`,...
3. **Chống nạp trùng thông minh:** Nhận diện tệp dựa trên Tên file và Dung lượng. Dù bạn di chuyển tệp sang thư mục khác, phần mềm vẫn biết nó đã nạp rồi.
4. **Cập nhật Path:** Tự động cập nhật đường dẫn mới vào lịch sử nạp dữ liệu nếu file được di chuyển.

## 📦 Cách cài đặt & Chạy:
1. Giải nén toàn bộ 6 file vào một thư mục.
2. Chỉnh sửa `appsettings.json` để khai báo:
   - `ConnectionString`: Chuỗi kết nối SQL Server của bạn.
   - `BaseFolder`: Thư mục mẹ chứa các thư mục ngày đo (`yyyy-MM-dd`).
3. Chạy file `ImportData.exe`.

## ⚙️ Cấu hình SQL:
- Bảng dữ liệu: `SortingDataImportExcel`
- Bảng lịch sử: `ExcelImportHistory`

---
*Phát triển bởi Antigravity AI Assistant.*
