using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using ImportData.Core;
using ImportData.Services;
using ImportData.Helpers;

namespace ImportData
{
    public partial class Form1 : Form
    {
        private const int MAX_LOG_LINES = 1000;
        
        private readonly AppConfig _config;
        private DatabaseService _dbService;
        private ExcelService _excelService;
        
        private FileSystemWatcher _watcher;
        private bool _isProcessing;
        
        // --- HỆ THỐNG TỰ PHỤC HỒI (SELF-HEALING) ---
        private System.Windows.Forms.Timer _healthTimer;
        private string _lastState = "";
        private bool _isSystemHealthy = false;

        // --- HỆ THỐNG SYSTEM TRAY (CHẠY NGẦM) ---
        private NotifyIcon _trayIcon;

        public Form1()
        {
            InitializeComponent();
            
            // Bật chế độ tự vẽ (OwnerDraw) để đổi màu từng dòng Log
            lstLogs.DrawMode = DrawMode.OwnerDrawFixed;
            lstLogs.DrawItem += LstLogs_DrawItem;
            
            _config = new AppConfig();
            _config.Load(Log);

            _dbService = new DatabaseService(_config, Log);
            _excelService = new ExcelService(Log);

            SystemHelper.SetStartup(Log);

            // Cấu hình Icon chạy ngầm ở góc dưới màn hình (Không có menu để cấm thoát)
            _trayIcon = new NotifyIcon()
            {
                Icon = SystemIcons.Information,
                Text = "Hệ thống Import Data (Đang chạy ngầm)",
                Visible = true
            };

            // Nháy đúp chuột để mở lại cửa sổ
            _trayIcon.DoubleClick += (s, e) => {
                this.Show();
                this.WindowState = FormWindowState.Normal;
            };

            this.ShowInTaskbar = true;
            this.WindowState = FormWindowState.Normal;

            this.Shown += Form1_Shown;
            this.FormClosing += Form1_FormClosing;
        }

        private async void Form1_Shown(object sender, EventArgs e)
        {
            Log($">>> HỆ THỐNG KHỞI CHẠY - ĐANG THEO DÕI: {_config.BaseFolder} <<<");
            
            // Khởi tạo và chạy vòng lặp kiểm tra sức khỏe mỗi 10 giây
            _healthTimer = new System.Windows.Forms.Timer();
            _healthTimer.Interval = 10000;
            _healthTimer.Tick += async (s, ev) => await PerformHealthCheckLoopAsync();
            
            // Chạy ngay lần đầu tiên
            await PerformHealthCheckLoopAsync();
            _healthTimer.Start();
        }

        private async Task PerformHealthCheckLoopAsync()
        {
            // Nạp lại cấu hình (Hot Reload)
            _config.Load(null); // Không truyền tham số Log để tránh spam

            string currentState = "HEALTHY";
            bool currentPathOk = true;
            bool currentDbOk = true;

            // 1. Kiểm tra Folder
            if (!Directory.Exists(_config.BaseFolder))
            {
                currentPathOk = false;
                currentState = "PATH_ERROR";
            }

            // 2. Kiểm tra Database
            if (currentPathOk)
            {
                currentDbOk = await _dbService.TestConnectionAsync();
                if (!currentDbOk) currentState = "DB_ERROR";
            }

            _isSystemHealthy = currentPathOk && currentDbOk;

            // Xử lý báo Log khi chuyển trạng thái (tránh spam)
            if (currentState != _lastState)
            {
                if (_isSystemHealthy)
                {
                    Log($"[OK] Hệ thống sẵn sàng. Đang theo dõi: {_config.BaseFolder}");
                    UpdateStatus("Hệ thống Sẵn sàng", Color.Green);
                    
                    InitWatcher();
                    await PerformSyncAsync();
                }
                else
                {
                    StopWatcher(); // Dừng quét file khi có lỗi
                    
                    if (!currentPathOk)
                    {
                        Log($"[LỖI THƯ MỤC] Không tìm thấy đường dẫn: {_config.BaseFolder}");
                        UpdateStatus("Lỗi: Sai đường dẫn thư mục", Color.Red);
                    }
                    else if (!currentDbOk)
                    {
                        Log("[LỖI DATABASE] Không kết nối được SQL Server (Mất mạng hoặc sai cấu hình).");
                        UpdateStatus("Lỗi: Không kết nối được Database", Color.Red);
                    }
                }
                _lastState = currentState;
            }
        }

        private void StopWatcher()
        {
            if (_watcher != null)
            {
                _watcher.EnableRaisingEvents = false;
                _watcher.Dispose();
                _watcher = null;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true; 
                this.Hide(); // Ẩn hẳn Form, không để dính trên Taskbar nữa
                
                // Hiển thị bong bóng thông báo góc System Tray
                _trayIcon.ShowBalloonTip(3000, "Thông báo", "Ứng dụng vẫn đang chạy ngầm để nạp dữ liệu.", ToolTipIcon.Info);
                
                Log("Đã thu nhỏ ứng dụng xuống khay hệ thống (System Tray).");
            }
        }

        private void UpdateStatus(string message, Color color)
        {
            lblStatus.Text = message;
            lblStatus.ForeColor = color;
        }

        private void LstLogs_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            e.DrawBackground();

            string text = lstLogs.Items[e.Index].ToString();
            
            // Mặc định chữ màu xanh lá cây, nếu có mác LỖI thì tô màu đỏ
            Brush textBrush = Brushes.LimeGreen; 
            if (text.Contains("[LỖI") || text.Contains("[Lỗi]"))
            {
                textBrush = Brushes.Red;
            }
            else if (text.Contains("[OK]"))
            {
                textBrush = Brushes.Cyan;
            }

            e.Graphics.DrawString(text, e.Font, textBrush, e.Bounds, StringFormat.GenericDefault);
            e.DrawFocusRectangle();
        }


        private void Log(string message)
        {
            if (lstLogs.InvokeRequired)
            {
                lstLogs.Invoke(new Action(() => Log(message)));
                return;
            }
            
            string time = DateTime.Now.ToString("HH:mm:ss");
            lstLogs.Items.Add($"[{time}] {message}");
            lstLogs.SelectedIndex = lstLogs.Items.Count - 1; 
            
            if (lstLogs.Items.Count > MAX_LOG_LINES) 
                lstLogs.Items.RemoveAt(0);
        }

        private void InitWatcher()
        {
            if (!Directory.Exists(_config.BaseFolder))
            {
                Directory.CreateDirectory(_config.BaseFolder);
            }

            _watcher = new FileSystemWatcher(_config.BaseFolder)
            {
                IncludeSubdirectories = true,
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,
                Filter = "*.*",
                EnableRaisingEvents = true
            };

            _watcher.Created += OnFileChanged;
            _watcher.Changed += OnFileChanged;
            _watcher.Renamed += OnFileChanged;
        }

        private async void OnFileChanged(object sender, FileSystemEventArgs e)
        {
            if (!_isSystemHealthy) return; // Bảo vệ: Không xử lý nếu đang mất mạng

            string ext = Path.GetExtension(e.FullPath).ToLower();

            if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm") return;

            string todayFolder = DateTime.Now.ToString("yyyy-MM-dd");
            if (!e.FullPath.Contains(todayFolder)) return;

            if (_isProcessing) return;
            
            try 
            {
                _isProcessing = true;
                
                if (await WaitForFileReadyAsync(e.FullPath))
                {
                    await PerformSyncAsync();
                }
                else
                {
                    Log($"Bỏ qua file {Path.GetFileName(e.FullPath)} do tiến trình khác đang chiếm dụng.");
                }
            }
            finally 
            {
                _isProcessing = false;
            }
        }

        private async Task<bool> WaitForFileReadyAsync(string filePath)
        {
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        return true;
                    }
                }
                catch (IOException)
                {
                    await Task.Delay(1000);
                }
            }
            return false;
        }

        private async Task PerformSyncAsync()
        {
            if (!_isSystemHealthy) return; // Bảo vệ: Không đồng bộ khi lỗi

            string sourceFolder = Path.Combine(_config.BaseFolder, DateTime.Now.ToString("yyyy-MM-dd"));
            
            if (!Directory.Exists(sourceFolder))
            {
                UpdateStatus($"Đang chờ thư mục: {DateTime.Now:yyyy-MM-dd}", lblStatus.ForeColor);
                return;
            }

            this.UseWaitCursor = true;
            UpdateStatus("Hệ thống đang đồng bộ...", lblStatus.ForeColor);

            try
            {
                string[] files = Directory.GetFiles(sourceFolder, "*.*")
                    .Where(s => s.EndsWith(".xlsx") || s.EndsWith(".xls") || s.EndsWith(".xlsm"))
                    .ToArray();

                foreach (string filePath in files)
                {
                    string fileName = Path.GetFileName(filePath);
                    
                    if (await _dbService.IsFileImportedAsync(fileName)) continue;

                    Log($"Phát hiện mới: {fileName} - Đang nạp dữ liệu...");
                    
                    bool success = await ProcessSingleFileAsync(filePath, fileName);
                    if (success)
                    {
                        Log($"[Xong] File {fileName} nạp thành công.");
                    }
                    else
                    {
                        Log($"[Lỗi] File {fileName} nạp thất bại.");
                    }
                }
                
                UpdateStatus("Hệ thống Sẵn sàng", Color.Green);
            }
            catch (Exception ex)
            {
                UpdateStatus("Lỗi đồng bộ dữ liệu", Color.Red);
                Log($"LỖI ĐỒNG BỘ: {ex.Message}");
            }
            finally
            {
                this.UseWaitCursor = false;
            }
        }

        private async Task<bool> ProcessSingleFileAsync(string filePath, string fileName)
        {
            try
            {
                var dataTable = _excelService.ReadExcelFile(filePath);
                if (dataTable == null || dataTable.Rows.Count == 0) return false;

                int rowsImported = await _dbService.ExecuteImportBatchAsync(dataTable);
                if (rowsImported > 0)
                {
                    await _dbService.MarkFileAsImportedAsync(fileName, filePath);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                Log($"Lỗi xử lý file {fileName}: {ex.Message}");
                return false;
            }
        }
    }
}
