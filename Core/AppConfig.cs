using Microsoft.Extensions.Configuration;
using System;
using System.IO;

namespace ImportData.Core
{
    /// <summary>
    /// Lớp AppConfig: Quản lý cấu hình hệ thống.
    /// Giúp App biết "Kết nối DB nào?" và "Quét thư mục nào?".
    /// </summary>
    public class AppConfig
    {
        // Default connection string if appsettings.json is missing.
        private const string DEFAULT_DB_CONN = @"Server=.;Database=CapacitorDB;Integrated Security=True;TrustServerCertificate=True;";
        
        // Default folder name.
        private const string DEFAULT_FOLDER = @"task";

        public string ConnectionString { get; set; }
        public string BaseFolder { get; set; }

        /// <summary>
        /// Khởi tạo mặc định ban đầu. Hàm này sẽ tự chạy đầu tiên khi gõ "new AppConfig()".
        /// </summary>
        public AppConfig() 
        {
            ConnectionString = DEFAULT_DB_CONN;
            // Set default base folder to Desktop/task.
            BaseFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), DEFAULT_FOLDER); 
        }

        /// <summary>
        /// Hàm này dùng để nạp file cấu hình appsettings.json lên.
        /// Cho phép nạp lại khi App đang chạy (nhờ quét timer ngầm bên ngoài).
        /// </summary>
        public void Load(Action<string> logger)
        {
            try
            {
                // Lấy đường dẫn thư mục gốc nơi ứng dụng đang chạy (thường là thư mục bin\Debug hoặc bin\Release).
                string runDir = AppDomain.CurrentDomain.BaseDirectory; 
                
                // Tìm thư mục chứa mã nguồn (Source Directory) bằng cách lùi lại 3 cấp từ thư mục chạy.
                // Việc này giúp App tìm thấy file appsettings.json trong thư mục Project khi đang chạy ở chế độ Debug trên Visual Studio.
                string sourceDir = Path.GetFullPath(Path.Combine(runDir, @"..\..\..\")); 
                
                // Xác định đường dẫn đầy đủ của file cấu hình trong cả thư mục chạy (Run) và thư mục nguồn (Source).
                string settingsPath = Path.Combine(runDir, "appsettings.json"); 
                string sourceSettingsPath = Path.Combine(sourceDir, "appsettings.json"); 

                // Kiểm tra xem file ở đâu tồn tại thì dùng file đó (Ưu tiên file ở thư mục Source để tiện chỉnh sửa khi Code).
                string targetFile = File.Exists(sourceSettingsPath) ? sourceSettingsPath : settingsPath; 

                // Nếu tìm thấy file cấu hình thì bắt đầu nạp.
                if (File.Exists(targetFile))
                {
                    string json;
                    // Mở file với chế độ cho phép các chương trình khác cũng có thể đọc/ghi (FileShare.ReadWrite).
                    // Giúp tránh lỗi "File is in use" nếu người dùng đang mở file appsettings.json bằng Notepad.
                    using (var fs = new FileStream(targetFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) 
                    using (var sr = new StreamReader(fs)) 
                    {
                        json = sr.ReadToEnd(); // Đọc toàn bộ nội dung file Json vào biến chuỗi.
                    }

                    // Nếu nội dung đọc được không bị trống.
                    if (!string.IsNullOrWhiteSpace(json))
                    {
                        // Sử dụng thư viện System.Text.Json để phân tích (Parse) chuỗi văn bản thành đối tượng Json.
                        using (var doc = System.Text.Json.JsonDocument.Parse(json)) 
                        {
                            var root = doc.RootElement; // Lấy phần tử gốc của file Json.
                            
                            // 1. Tìm mục "ConnectionStrings" và lấy giá trị "DefaultConnection".
                            if (root.TryGetProperty("ConnectionStrings", out var connStrings) && 
                                connStrings.TryGetProperty("DefaultConnection", out var defaultConn))
                            {
                                // Gán chuỗi kết nối Database, nếu trong file bị null thì giữ nguyên giá trị mặc định cũ.
                                ConnectionString = defaultConn.GetString() ?? ConnectionString; 
                            }

                            // 2. Tìm mục "FolderSettings" và lấy giá trị "BaseFolder".
                            if (root.TryGetProperty("FolderSettings", out var folderSettings) && 
                                folderSettings.TryGetProperty("BaseFolder", out var baseFld))
                            {
                                // Gán đường dẫn thư mục quét, nếu null thì giữ nguyên mặc định.
                                BaseFolder = baseFld.GetString() ?? BaseFolder; 
                            }
                        }
                    }
                    
                    // Nếu có truyền hàm logger vào thì in thông báo nạp thành công tên file nào.
                    if (logger != null) logger.Invoke($"Cấu hình đã nạp: {Path.GetFileName(targetFile)}");
                }
            }
            catch (Exception) // Hứng trọi mọi lỗi nảy sinh (ví dụ file Json viết sai cú pháp ngoặc nhọn).
            {
                // Bỏ qua lỗi: Nếu file cấu hình hỏng, App sẽ dùng thông số cũ hoặc mặc định để tiếp tục sống.
                // Ứng dụng sẽ tự động thử nạp lại sau một chu kỳ 10 giây khác.
            }
        }
    }
}
