# 📦 Hướng Dẫn Cài Đặt — Auto Import Capacitor Data

## Yêu Cầu Hệ Thống

- Windows 10/11 (64-bit)
- SQL Server đã cài đặt và chạy
- Không cần cài .NET Runtime (đã đóng gói sẵn trong file .exe)

---

## Bước 1: Giải Nén

1. Copy file `ImportData_v1.0.zip` vào máy tính cần cài
2. Chuột phải → **Extract All** (Giải nén tất cả)
3. Chọn thư mục cài đặt, ví dụ: `C:\ImportData\`

Sau khi giải nén sẽ có 2 file:

```
C:\ImportData\
├── ImportData.exe          (120 MB - File chạy chính)
└── appsettings.json        (File cấu hình)
```

---

## Bước 2: Chuẩn Bị Database

Mở **SQL Server Management Studio (SSMS)** và chạy file `Setup_Database.sql` trên server thật.

Script này sẽ:
- Thêm các cột cần thiết vào bảng `CapacitorLogs` (nếu chưa có)
- Tạo bảng `ImportHistory` để theo dõi lịch sử nạp file

---

## Bước 3: Chỉnh Cấu Hình

Mở file `appsettings.json` bằng **Notepad** và chỉnh các thông số:

```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=ĐỊA_CHỈ_SQL_SERVER;Database=CapacitorDB;Integrated Security=True;TrustServerCertificate=True;"
  },
  "FolderSettings": {
    "BaseFolder": "ĐƯỜNG_DẪN_THƯ_MỤC_MÁY_ĐO"
  },
  "SyncSettings": {
    "ScanIntervalSeconds": 600
  },
  "HealthCheckSettings": {
    "ConnectionTimeoutSeconds": 5
  }
}
```

### Giải thích các thông số:

| Thông số | Ý nghĩa | Ví dụ |
|----------|---------|-------|
| `DefaultConnection` | Chuỗi kết nối SQL Server | `Server=192.168.1.100;Database=CapacitorDB;...` |
| `BaseFolder` | Thư mục chứa dữ liệu máy đo | `C:\Data\CapacitorLogs` |
| `ConnectionTimeoutSeconds` | Thời gian chờ kết nối DB (giây) | `5` (mạng LAN), `10` (mạng VPN) |

### Nếu SQL Server dùng tài khoản SQL (không dùng Windows Auth):

```json
"DefaultConnection": "Server=192.168.1.100;Database=CapacitorDB;User Id=sa;Password=MatKhau123;TrustServerCertificate=True;"
```

---

## Bước 4: Chạy Ứng Dụng

1. Nhấp đúp `ImportData.exe` để khởi động
2. Ứng dụng sẽ tự động:
   - Kiểm tra kết nối SQL Server
   - Theo dõi thư mục máy đo
   - Nạp dữ liệu Excel mới vào Database
3. Khi đóng cửa sổ, ứng dụng **chạy ngầm** ở khay hệ thống (góc đồng hồ)
4. Nhấp đúp icon ở khay để mở lại giao diện

---

## Bước 5: Tự Khởi Động Cùng Windows

Ứng dụng **tự đăng ký** khởi động cùng Windows khi chạy lần đầu.
Không cần cấu hình thêm.

---

## Ghi Chú

- File `appsettings.json` có thể chỉnh **khi app đang chạy** — app tự cập nhật mỗi 10 giây
- Ứng dụng chỉ nạp file Excel trong thư mục ngày hôm nay (định dạng `yyyy-MM-dd`)
- File đã nạp sẽ không bị nạp lại lần thứ 2
