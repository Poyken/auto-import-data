# HỆ THỐNG TỰ ĐỘNG NHẬP DỮ LIỆU TỤ ĐIỆN (AUTO IMPORT DATA)

Đây là ứng dụng chạy ngầm trên Windows nhằm tự động quét và nhập dữ liệu từ các tệp Excel kết quả vào cơ sở dữ liệu SQL Server.

## 📦 Hướng dẫn mang sang máy khác (Deployment)

Đây là cách nhanh nhất và ổn định nhất để cài đặt trên máy tính mới:

1.  **Chuẩn bị file**: Copy file duy nhất [ImportData.exe](dist/SingleFile/ImportData.exe) (Bản Single-File) và file `appsettings.json` vào máy tính mới.
2.  **Cơ sở dữ liệu**: Chạy file `Setup_Database.sql` trên SQL Server của máy đích để tạo bảng.
3.  **Cấu hình**: Mở `appsettings.json` và chỉnh:
    - `DefaultConnection`: IP/Tên SQL Server của máy mới.
    - `BaseFolder`: Thư mục chứa file Excel của máy đo mới.
4.  **Chạy**: Nhấp đúp vào `ImportData.exe`. Ứng dụng sẽ tự động chạy và thu nhỏ xuống khay đồng hồ.

---

## 🧹 Cách dọn dẹp Source Code (Cleanup)

Để có bộ mã nguồn sạch nhất, bạn chỉ cần giữ lại các thư mục/file sau:
- **Thư mục**: `Core`, `Services`.
- **File mã nguồn**: `Program.cs`, `Form1.cs`, `Form1.Designer.cs`, `Form1.resx`.
- **File cấu hình & DB**: `ImportData.csproj`, `appsettings.json`, `Setup_Database.sql`, `app_icon.ico`.

*Bạn có thể xóa hoàn toàn các thư mục `bin`, `obj`, `.vs` và `dist` khi muốn lưu trữ mã nguồn sạch.*

---

## 🛠 Tính năng kỹ thuật
- **Log Xanh (Hacker Style)**: Giờ đây đã có khoảng cách 110px giúp bạn dễ đọc giờ và nội dung log.
- **Trạng thái tức thì**: Khi bạn bấm "Đổi thư mục", lỗi "Lỗi đường dẫn" sẽ biến mất ngay lập tức nếu đường dẫn đúng.
- **Tự lưu cấu hình**: Bạn đổi thư mục trong giao diện, ứng dụng sẽ tự lưu vào file để không bị mất khi khởi động lại.
