using Microsoft.Win32;      // Khai báo cấp quyền can thiệp vào Registry O.S, khu vực lưu cấu hình hệ thống máy ẩn.
using System;               // Thư viện các toán tử và hàm cơ bản xử lý chuỗi của C#.
using System.Windows.Forms; // Thư viện WinForms dùng lệnh Application.ExecutablePath để hệ thống tự tìm được đường dẫn của mình.

namespace ImportData.Helpers 
{
    /// <summary>
    /// Lớp SystemHelper: Cung cấp các thao tác công cụ liên quan đến xử lý tương tác trực tiếp đối với Hệ điều hành Windows.
    /// static: Cho phép nhà phát triển dùng trực tiếp các hàm chạy nhanh mà không cần gọi hàm Khởi tạo New.
    /// </summary>
    public static class SystemHelper 
    {
        /// <summary>
        /// Hàm SetStartup: Gọi tính năng thêm tiến trình tự chạy ngầm ứng dụng ngay lặp tức khi Windows khởi hoàn tất khi thao tác khởi nguồn Desktop xong.
        /// </summary>
        public static void SetStartup(Action<string> logger) 
        {
            try 
            {
                // Gọi vào Registry.CurrentUser mục đích là để cập nhật quyền riêng của máy tài khoản hiện có.
                // Tránh lỗi bảo vệ thư mục từ Virus quét ở một số Hệ Điều Hành O.S. (Lưu ở 'Run').
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    // Khởi tạo phím khóa "AutoImportData" và gán giá trị mang "Đường dẫn tuyệt đối .exe vật lý File Ứng dụng này đang chạy".
                    // Máy O.S sẽ nắm lấy path file chạy của C# và gọi tệp khi có điện.
                    key.SetValue("AutoImportData", Application.ExecutablePath);

                } // Tự động Đóng cấu hình khi chui ra khỏi nhịp Hàm using.
            }
            catch (Exception ex) 
            {
                // Một vài hệ thống bảo mật cực mạnh như Anti-Virus đôi lúc chặn việc này.
                // Chúng ta sẽ bỏ qua không làm ứng dụng sập hỏng, mà chỉ rớt log báo lên để người dùng Admin vào gỡ hệ máy Security thủ công.
                logger?.Invoke($"Lỗi đăng ký thiết lập tự khởi động cùng hệ thống: Thông số {ex.Message}");
            }
        }
    }
}
