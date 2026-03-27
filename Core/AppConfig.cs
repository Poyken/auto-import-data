using System;
using System.IO;
using System.Text.Json;

namespace ImportData.Core
{
    /// <summary>
    /// Manages application configuration and persistence.
    /// Handles database connection strings, folder paths, and synchronization settings.
    /// </summary>
    public class AppConfig
    {
        private const string HardcodedConn = @"Server=dbserver.hycap.co.kr,5398;Database=SmartFactoryV2;User ID=vinaadmin;Password=vina1234%6&8;TrustServerCertificate=True;";
        private const bool UseHardcodedConn = true; 
        private const string DefaultFolderName = "task";

        public string ConnectionString { get; set; }
        public string BaseFolder { get; set; }
        public int ScanIntervalSeconds { get; set; } = 600;
        public int HealthCheckTimeoutSeconds { get; set; } = 5;

        private string _configFilePath;

        public AppConfig() 
        {
            ConnectionString = HardcodedConn;
            BaseFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), DefaultFolderName); 
        }

        /// <summary>
        /// Loads configuration from appsettings.json, with fallbacks to defaults.
        /// </summary>
        public void Load(Action<string> logger = null)
        {
            try
            {
                string runDir = AppDomain.CurrentDomain.BaseDirectory; 
                string sourceDir = Path.GetFullPath(Path.Combine(runDir, @"..\..\..\")); 
                
                string settingsPath = Path.Combine(runDir, "appsettings.json"); 
                string sourceSettingsPath = Path.Combine(sourceDir, "appsettings.json"); 

                _configFilePath = File.Exists(sourceSettingsPath) ? sourceSettingsPath : settingsPath; 

                if (File.Exists(_configFilePath))
                {
                    string json;
                    using (var fs = new FileStream(_configFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) 
                    using (var sr = new StreamReader(fs)) 
                    {
                        json = sr.ReadToEnd();
                    }

                    if (!string.IsNullOrWhiteSpace(json))
                    {
                        using (var doc = JsonDocument.Parse(json)) 
                        {
                            var root = doc.RootElement; 
                            
                            if (!UseHardcodedConn && root.TryGetProperty("ConnectionStrings", out var connSection))
                            {
                                ConnectionString = connSection.GetProperty("DefaultConnection").GetString() ?? ConnectionString;
                            }

                            if (root.TryGetProperty("FolderSettings", out var folderSection))
                            {
                                BaseFolder = folderSection.GetProperty("BaseFolder").GetString() ?? BaseFolder;
                            }

                            if (root.TryGetProperty("SyncSettings", out var syncSection))
                            {
                                ScanIntervalSeconds = syncSection.GetProperty("ScanIntervalSeconds").GetInt32();
                            }

                            if (root.TryGetProperty("HealthCheckSettings", out var healthSection))
                            {
                                HealthCheckTimeoutSeconds = healthSection.GetProperty("ConnectionTimeoutSeconds").GetInt32();
                            }
                        }
                    }
                    logger?.Invoke($"[CONFIG] Configuration loaded from: {Path.GetFileName(_configFilePath)}");
                }
            }
            catch (Exception ex)
            {
                logger?.Invoke($"[CONFIG-WARN] Failed to load config, using defaults: {ex.Message}");
            }
        }

        /// <summary>
        /// Saves the current configuration properties back to the appsettings.json file.
        /// </summary>
        public void Save()
        {
            try
            {
                if (string.IsNullOrEmpty(_configFilePath)) return;

                var configData = new
                {
                    ConnectionStrings = new { DefaultConnection = ConnectionString },
                    FolderSettings = new { BaseFolder = BaseFolder },
                    SyncSettings = new { ScanIntervalSeconds = ScanIntervalSeconds },
                    HealthCheckSettings = new { ConnectionTimeoutSeconds = HealthCheckTimeoutSeconds }
                };

                string json = JsonSerializer.Serialize(configData, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_configFilePath, json);
            }
            catch (Exception)
            {
                // Silently fail if save is not possible (e.g. file locked)
            }
        }
    }
}
