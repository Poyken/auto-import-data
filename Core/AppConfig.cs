using Microsoft.Extensions.Configuration;
using System;
using System.IO;

namespace ImportData.Core
{
    /// <summary>
    /// Lớp quản lý cấu hình hệ thống
    /// </summary>
    public class AppConfig
    {
        private const string DEFAULT_DB_CONN = @"Server=.;Database=CapacitorDB;Integrated Security=True;TrustServerCertificate=True;";
        private const string DEFAULT_FOLDER = @"task";

        public string ConnectionString { get; set; }
        public string BaseFolder { get; set; }

        public AppConfig()
        {
            ConnectionString = DEFAULT_DB_CONN;
            BaseFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), DEFAULT_FOLDER);
        }

        public void Load(Action<string> logger)
        {
            try
            {
                // Thư mục chạy thực tế (Thường là bin/Debug/...)
                string runDir = AppDomain.CurrentDomain.BaseDirectory;
                // Thư mục Source Code (Lùi lại 3 cấp: \bin\Debug\net10.0-windows\)
                string sourceDir = Path.GetFullPath(Path.Combine(runDir, @"..\..\..\"));

                string settingsPath = Path.Combine(runDir, "appsettings.json");
                string sourceSettingsPath = Path.Combine(sourceDir, "appsettings.json");

                // Ưu tiên đọc file appsettings.json ở thư mục Code Gốc nếu đang chạy Debug
                string targetFile = File.Exists(sourceSettingsPath) ? sourceSettingsPath : settingsPath;

                if (File.Exists(targetFile))
                {
                    // LỚP BẢO VỆ 2: NẠP NÓNG CẤU HÌNH BẰNG RAW FILE STREAM MÀ KHÔNG DÙNG CACHE
                    string json;
                    using (var fs = new FileStream(targetFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var sr = new StreamReader(fs))
                    {
                        json = sr.ReadToEnd();
                    }

                    if (!string.IsNullOrWhiteSpace(json))
                    {
                        using (var doc = System.Text.Json.JsonDocument.Parse(json))
                        {
                            var root = doc.RootElement;
                            if (root.TryGetProperty("ConnectionStrings", out var connStrings) && 
                                connStrings.TryGetProperty("DefaultConnection", out var defaultConn))
                            {
                                ConnectionString = defaultConn.GetString() ?? ConnectionString;
                            }

                            if (root.TryGetProperty("FolderSettings", out var folderSettings) && 
                                folderSettings.TryGetProperty("BaseFolder", out var baseFld))
                            {
                                BaseFolder = baseFld.GetString() ?? BaseFolder;
                            }
                        }
                    }
                    if (logger != null) logger.Invoke($"Đã tải cấu hình từ file appsettings.json");
                }
                else
                {
                    if (logger != null) logger.Invoke("Không tìm thấy appsettings.json, dùng cấu hình mặc định.");
                }
            }
            catch (Exception ex)
            {
                // Không báo lỗi khi người dùng đang gõ dở file JSON gây sai cú pháp (sẽ chờ chu kỳ sau)
            }
        }
    }
}
