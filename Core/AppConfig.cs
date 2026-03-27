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
        // --- CẤU HÌNH FIX CỨNG (SỬA TẠI ĐÂY) ---
        // 1. Dán chuỗi kết nối Database có mã hóa của bạn vào đây.
        private const string HARDCODED_CONN = @"Server=.;Database=CapacitorDB;Integrated Security=True;TrustServerCertificate=True;";
        // 2. Đổi thành 'true' để ép buộc App chỉ dùng chuỗi trên, bỏ qua file appsettings.json.
        private const bool USE_HARDCODED_CONN = false; 

        // Default connection string if appsettings.json is missing and hardcode is disabled.
        private const string DEFAULT_DB_CONN = @"Server=.;Database=CapacitorDB;Integrated Security=True;TrustServerCertificate=True;";
        
        // Default folder name.
        private const string DEFAULT_FOLDER = @"task";

        public string ConnectionString { get; set; }
        public string BaseFolder { get; set; }
        private string _loadedFilePath; // Đường dẫn file đã nạp.

        // Thời gian chờ kết nối tối đa (giây) khi health check kiểm tra SQL.
        // Mặc định 5 giây, user có thể chỉnh trong appsettings.json mục HealthCheckSettings.ConnectionTimeoutSeconds.
        public int HealthCheckTimeoutSeconds { get; set; } = 5;

        /// <summary>
        /// Khởi tạo mặc định ban đầu. Hàm này sẽ tự chạy đầu tiên khi gõ "new AppConfig()".
        /// </summary>
        public AppConfig() 
        {
            // Nếu chọn Fix cứng thì dùng luôn, không thì dùng mặc định tạm thời.
            ConnectionString = USE_HARDCODED_CONN ? HARDCODED_CONN : DEFAULT_DB_CONN;
            
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
                _loadedFilePath = targetFile; // Ghi nhớ đường dẫn để xíu nữa Save đè vào đúng chỗ này.

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
                            if (!USE_HARDCODED_CONN && 
                                root.TryGetProperty("ConnectionStrings", out var connStrings) && 
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

                            // 3. Tìm mục "HealthCheckSettings" và lấy giá trị "ConnectionTimeoutSeconds".
                            if (root.TryGetProperty("HealthCheckSettings", out var healthSettings) && 
                                healthSettings.TryGetProperty("ConnectionTimeoutSeconds", out var timeoutVal))
                            {
                                int val = timeoutVal.GetInt32();
                                // Giới hạn trong khoảng hợp lý 1-30 giây, tránh user ghi nhầm số âm hoặc quá lớn.
                                HealthCheckTimeoutSeconds = Math.Max(1, Math.Min(30, val));
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

        /// <summary>
        /// Lưu cấu hình hiện tại xuống file appsettings.json.
        /// </summary>
        public void Save()
        {
            try
            {
                if (string.IsNullOrEmpty(_loadedFilePath)) return;

                // Tạo đối tượng JSON để lưu.
                var configData = new
                {
                    ConnectionStrings = new
                    {
                        DefaultConnection = ConnectionString
                    },
                    FolderSettings = new
                    {
                        BaseFolder = BaseFolder
                    },
                    HealthCheckSettings = new
                    {
                        ConnectionTimeoutSeconds = HealthCheckTimeoutSeconds
                    }
                };

                // Chuyển đối tượng thành chuỗi JSON đẹp mắt.
                string json = System.Text.Json.JsonSerializer.Serialize(configData, new System.Text.Json.JsonSerializerOptions 
                { 
                    WriteIndented = true 
                });

                // Ghi đè vào file.
                File.WriteAllText(_loadedFilePath, json);
            }
            catch (Exception)
            {
                // Lờ đi nếu không lưu được (ví dụ file đang bị khóa cứng).
            }
        }
    }
}
