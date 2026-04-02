using Microsoft.Win32;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using ImportData.Core;
using ImportData.Services;

namespace ImportData
{
    /// <summary>
    /// Giao diện chính Form1: Trung tâm điều phối của ứng dụng Auto Import.
    /// Quản lý việc theo dõi thư mục thời gian thực, đồng bộ dữ liệu và hiển thị trạng thái hệ thống.
    /// </summary>
    public partial class Form1 : Form
    {
        // Giới hạn 1000 dòng nhật ký trên màn hình để tiết kiệm RAM.
        private const int MaxLogLines = 1000; 
        
        private readonly AppConfig _config;           // Thông số cấu hình (Thư mục, SQL).
        private readonly DatabaseService _dbService; // Dịch vụ SQL Server.
        private readonly ExcelService _excelService; // Dịch vụ đọc file Excel đo đạc.
        private FileSystemWatcher _watcher;          // "Cảm biến" cảm nhận file mới sinh.
        private bool _isProcessing;                  // Cờ ngăn việc quét tệp chồng chéo.
        private System.Windows.Forms.Timer _healthTimer; // Đồng hồ 10 giây khám sức khỏe app.
        private string _lastHealthState = "";        // Ghi lại lỗi lần cuối để tránh spam log.
        private bool _isSystemHealthy = false;       // App đang ổn (True) hay đang lỗi (False).
        private NotifyIcon _trayIcon;                 // Biểu tượng nhỏ chạy dưới góc khay Windows.

        // HÀM KHỞI TẠO: Chạy đầu tiên khi bật phần mềm.
        public Form1()
        {
            InitializeComponent(); 
            
            // 1. Cài đặt chế độ tự vẽ màu cho bảng nhật ký (Matrix-Style).
            lstLogs.DrawMode = DrawMode.OwnerDrawFixed;
            lstLogs.DrawItem += LstLogs_DrawItem; 
            
            // 2. Nạp cấu hình từ tệp appsettings.json.
            _config = new AppConfig();
            _config.Load();

            // 3. Khởi động các dịch vụ phụ trợ.
            _dbService = new DatabaseService(_config, Log); 
            _excelService = new ExcelService(Log);          

            // 4. Tạo biểu tượng chạy ngầm dưới khay đồng hồ.
            _trayIcon = new NotifyIcon()
            {
                Icon = this.Icon,
                Text = "Dịch vụ Auto Import (Đang chạy ngầm)",
                Visible = true
            };

            // Nhấp đúp và icon khay để hiện lại ứng dụng.
            _trayIcon.DoubleClick += (s, e) => {
                this.Show(); 
                this.WindowState = FormWindowState.Normal; 
            };

            this.ShowInTaskbar = true; 
            this.WindowState = FormWindowState.Normal; 

            // Gán sự kiện khi hiện giao diện và khi người dùng muốn tắt app.
            this.Shown += Form1_Shown; 
            this.FormClosing += Form1_FormClosing;  
        }

        // Khi Form đã hiện lên: Bắt đầu canh gác!
        private async void Form1_Shown(object sender, EventArgs e)
        {
            Log($"[KHỞI TẠO] Theo dõi thư mục: {_config.BaseFolder}");
            
            // Cài đặt đồng hồ 10 giây lặp lại.
            _healthTimer = new System.Windows.Forms.Timer { Interval = 10000 };
            _healthTimer.Tick += async (s, ev) => await PerformHealthCheckAsync(); 
            
            // Đăng ký app cùng Windows khởi động máy tính.
            RegisterAutoStart();
            
            // Khám bệnh lần 1 và bắt đầu đập nhịp đồng hồ.
            await PerformHealthCheckAsync(); 
            _healthTimer.Start();
        }

        /// <summary>
        /// Hàm khám sức khỏe hệ thống: Chạy lặp lại mỗi 10 giây.
        /// </summary>
        private async Task PerformHealthCheckAsync()
        {
            string previousFolder = _config.BaseFolder;
            
            // Không thực hiện Load() ở đây để tránh ghi đè giá trị Folder người dùng vừa chọn bằng tay.
            // Chỉ Load cấu hình định kỳ khi hệ thống đang ở trạng thái nhàn rỗi.

            string currentState = "HEALTHY";
            bool isDirectoryOk = Directory.Exists(_config.BaseFolder); // Thư mục còn sống không?
            bool isDatabaseOk = isDirectoryOk && await _dbService.TestConnectionAsync(); // SQL còn sống không?

            if (!isDirectoryOk) currentState = "PATH_ERROR";
            else if (!isDatabaseOk) currentState = "DB_ERROR";

            _isSystemHealthy = isDirectoryOk && isDatabaseOk;

            // Xử lý khi người dùng đổi sang thư mục canh gác mới.
            if (previousFolder != _config.BaseFolder)
            {
                Log($"[THÔNG BÁO] Hệ thống nhận diện thư mục quét mới: {_config.BaseFolder}"); 
                RestartWatcher(); 
                await SynchronizeAsync(); 
            }
            else if (_isSystemHealthy && _lastHealthState != "HEALTHY")
            {
                // Trường hợp phục hồi từ lỗi sang ổn định thì cũng cần sync lại.
                await SynchronizeAsync();
            }

            // Xử lý thay đổi tình trạng bệnh tật của hệ thống.
            if (currentState != _lastHealthState)
            {
                if (_isSystemHealthy) 
                {
                    if (_lastHealthState != "") Log("[OK] Hệ thống đã phục hồi kết nối.");
                    UpdateStatus("Hệ thống Sẵn sàng", Color.Green);
                    RestartWatcher(); // Kích hoạt lại cảm biến.
                    await SynchronizeAsync(); // Quét đồng bộ các tệp cũ còn sót lại.
                }
                else 
                {
                    StopWatcher(); // Hệ thống lỗi thì ngưng cảm biến cho nhẹ máy.
                    if (!isDirectoryOk) 
                    {
                        Log($"[LỖI] Không tìm thấy đường dẫn: {_config.BaseFolder}"); 
                        UpdateStatus("Lỗi Thư mục", Color.Red); 
                    }
                    else if (!isDatabaseOk) 
                    {
                        Log("[LỖI] Kết nối SQL Server thất bại.");
                        UpdateStatus("Lỗi kết nối SQL", Color.Red);
                    }
                }
                _lastHealthState = currentState;
            }
        }

        // Hàm khởi động lại cảm biến (Watcher) canh gác file mới.
        private void RestartWatcher()
        {
            StopWatcher();
            if (!Directory.Exists(_config.BaseFolder)) 
            {
                Directory.CreateDirectory(_config.BaseFolder); 
            }

            _watcher = new FileSystemWatcher(_config.BaseFolder) 
            {
                IncludeSubdirectories = true, // Canh gác cả thư mục con bên trong.
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite, // Nhận báo khi có file mới hoặc ghi thêm.
                Filter = "*.*",
                EnableRaisingEvents = true 
            };

            // Gắn sự kiện khi Watcher phát hiện sự thay đổi.
            _watcher.Created += OnFileEvent;
            _watcher.Changed += OnFileEvent;
            _watcher.Renamed += OnFileEvent;
        }

        // Hàm dừng cảm biến.
        private void StopWatcher()
        {
            if (_watcher != null) 
            {
                _watcher.EnableRaisingEvents = false;
                _watcher.Dispose();
                _watcher = null; 
            }
        }

        // Sự kiện khi người dùng nhấn dấu [X] tắt app.
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true; // Chặn lệnh tắt hẳn của Windows.
                this.Hide();     // Ẩn cửa sổ thôi.
                // Hiện bóng thông báo báo hiệu app vẫn sống ngầm dưới khay.
                _trayIcon.ShowBalloonTip(2000, "Import Data", "Ứng dụng vẫn đang chạy ngầm để canh file mới.", ToolTipIcon.Info); 
            }
        }

        // Sự kiện khi nhấn nút bấm "Thay đổi thư mục".
        private async void BtnChangeFolder_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Chọn thư mục máy đo sinh ra dữ liệu";
                dialog.SelectedPath = _config.BaseFolder;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    _config.BaseFolder = dialog.SelectedPath;
                    _config.Save(); // Lưu đè vào appsettings.json
                    Log($"[CHỌN THƯ MỤC] Đã đổi đường dẫn: {_config.BaseFolder}");
                    await PerformHealthCheckAsync();
                }
            }
        }

        // Hàm cập nhật dòng chữ Trạng thái (Xanh lá, Đỏ, Vàng).
        private void UpdateStatus(string message, Color color)
        {
            // Bảo vệ đa luồng (Cross-Thread safety).
            if (this.InvokeRequired) 
            {
                this.Invoke(new Action(() => UpdateStatus(message, color))); 
                return;
            }
            lblStatus.Text = message; 
            lblStatus.ForeColor = color; 
        }

        // Hàm Tự vẽ màu cho bảng nhật ký để có hiệu ứng phát sáng.
        private void LstLogs_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            string text = lstLogs.Items[e.Index].ToString(); 
            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

            // 1. VẼ NỀN (Màu đen hoặc màu Highlight xanh).
            if (isSelected)
            {
                using (var backBrush = new SolidBrush(Color.FromArgb(0, 120, 215))) 
                    e.Graphics.FillRectangle(backBrush, e.Bounds);
            }
            else
            {
                e.Graphics.FillRectangle(Brushes.Black, e.Bounds);
            }

            // 2. VẼ CHỮ.
            int timeEndIndex = text.IndexOf(']');
            string timePart = timeEndIndex > 0 ? text.Substring(0, timeEndIndex + 1) : "";
            string msgPart = timeEndIndex > 0 ? text.Substring(timeEndIndex + 1) : text;
            
            Brush textBrush = isSelected ? Brushes.White : Brushes.Lime;
            // Vẽ Thời gian và Nội dung nhật ký ở các tọa độ lệch nhau để tạo hàng lối.
            e.Graphics.DrawString(timePart, e.Font, textBrush, new PointF(e.Bounds.X + 5, e.Bounds.Y + 2));
            e.Graphics.DrawString(msgPart, e.Font, textBrush, new PointF(e.Bounds.X + 110, e.Bounds.Y + 2));
        }

        // Hàm đẩy một dòng tin nhắn vào bảng Log.
        private void Log(string message)
        {
            if (lstLogs.InvokeRequired) 
            {
                lstLogs.Invoke(new Action(() => Log(message))); 
                return;
            }
            
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            lstLogs.Items.Add($"[{timestamp}] {message}"); 
            lstLogs.SelectedIndex = lstLogs.Items.Count - 1; 
            
            // Xóa dòng Log cũ nhất nếu vượt quá 1000 dòng.
            if (lstLogs.Items.Count > MaxLogLines) 
                lstLogs.Items.RemoveAt(0); 
        }

        // Khi có tiếng chuông báo động 'Có File Mới' từ Cảm biến OS.
        private async void OnFileEvent(object sender, FileSystemEventArgs e)
        {
            if (!_isSystemHealthy) return; 

            // Chỉ quan tâm tệp Excel (.xlsx, .xls, .xlsm).
            string ext = Path.GetExtension(e.FullPath).ToLower();
            if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm") return;

            // Kiểm tra trạng thái nạp file thông minh
            int status = await _dbService.CheckImportStatusAsync(e.FullPath);
            if (status == 1) return; // Đã nạp rồi.
            if (status == 2) 
            {
                await _dbService.UpdateHistoryPathAsync(e.FullPath);
                return;
            }

            // Quan tâm tệp nằm trong thư mục ngày hôm nay (yyyy-MM-dd).
            string todayFolder = DateTime.Now.ToString("yyyy-MM-dd");
            if (!e.FullPath.Contains(todayFolder)) return; 

            if (_isProcessing) return; 
            
            try 
            {
                _isProcessing = true; 
                await ProcessSingleFileAsync(e.FullPath); // Nhảy vào xử lý nạp ngay!
            }
            finally 
            {
                _isProcessing = false; 
            }
        }

        // Đội quân Thu Vén: Quét sạch các file còn sót trong ngày hôm nay.
        private async Task SynchronizeAsync()
        {
            if (!_isSystemHealthy) return; 

            string todayFolder = DateTime.Now.ToString("yyyy-MM-dd");
            string sourcePath = Path.Combine(_config.BaseFolder, todayFolder);
            
            if (!Directory.Exists(sourcePath)) 
            {
                UpdateStatus($"Đang chờ dữ liệu: {todayFolder}", Color.White);
                return;
            }

            UpdateStatus("Đang quét tệp...", Color.Yellow); 

            try
            {
                // Sử dụng Task.Run để quét tệp ngầm, tránh treo giao diện (UI) khi thư mục có hàng nghìn file.
                string[] files = await Task.Run(() => 
                    Directory.GetFiles(sourcePath, "*.*", SearchOption.AllDirectories)
                        .Where(f => f.EndsWith(".xlsx") || f.EndsWith(".xls") || f.EndsWith(".xlsm"))
                        .ToArray()
                );

                if (files.Length == 0)
                {
                    UpdateStatus("Hệ thống Sẵn sàng", Color.Green);
                    return;
                }

                int count = 0;
                foreach (string file in files) 
                {
                    UpdateStatus($"Đồng bộ {++count}/{files.Length}", Color.Orange);
                    await ProcessSingleFileAsync(file); 
                }
                
                UpdateStatus("Hệ thống Sẵn sàng", Color.Green); 
            }
            catch (Exception ex)
            {
                Log($"[LỖI] Đồng bộ hóa thất bại: {ex.Message}"); 
                UpdateStatus("Lỗi Đồng bộ", Color.Red); 
            }
        }

        /// <summary>
        /// TRUNG TÂM XỬ LÝ: 'Giải cứu' tệp Excel từ ổ cứng để đưa vào SQL.
        /// </summary>
        private async Task ProcessSingleFileAsync(string filePath)
        {
            string fileName = Path.GetFileName(filePath);

            try
            {
                // 1. Kiểm tra trạng thái nạp file thông minh (V6-Ultimate)
                int status = await _dbService.CheckImportStatusAsync(filePath);
                if (status == 1) return; // Đã nạp thành công ở chính path này rồi.
                if (status == 2) 
                {
                    await _dbService.UpdateHistoryPathAsync(filePath);
                    return;
                }

                // 2. Đàm phán O.S: Chờ máy đo buông tay khỏi tệp (Ready to Read).
                if (!await IsFileReadyAsync(filePath))
                {
                    Log($"[BỎ QUA] Tệp đang bận hoặc bị khóa: {fileName}");
                    return;
                }

                Log($"[ĐANG NẠP] Xử lý tệp: {fileName}");

                // 3. Đọc dữ liệu Excel.
                var data = _excelService.ReadExcelFile(filePath);
                if (data == null || data.Rows.Count == 0) return; 

                // 4. Nhập hàng loạt vào SQL.
                await _dbService.ExecuteImportBatchAsync(data, fileName, filePath);
            }
            catch (Exception ex)
            {
                Log($"[LỖI] Xử lý {fileName} bị dừng giữa chừng: {ex.Message}"); 
            }
        }

        // Hàm giúp app chờ đợi file khi nó đang bị máy đo 'rặn' nốt dữ liệu.
        private async Task<bool> IsFileReadyAsync(string filePath)
        {
            // Thử 7 lần với khoảng nghỉ 1.5s (tổng cộng ~10 giây) để chờ máy đo nhả file.
            for (int i = 0; i < 7; i++)
            {
                try
                {
                    // Thử mượn file với quyền độc quyền (None).
                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                        return true; // Thành công - file không còn bị máy đo giữ.
                }
                catch (IOException) 
                {
                    // Chờ thêm 1.5 giây rồi gõ cửa lại lượt tiếp theo.
                    await Task.Delay(1500); 
                }
            }
            return false;
        }

        // Tự đăng ký app vào thư mục Startup của Registry Windows.
        private void RegisterAutoStart()
        {
            try 
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (key != null) key.SetValue("AutoImportData", Application.ExecutablePath);
                }
            }
            catch (Exception ex) 
            {
                Log($"[CẢNH BÁO] Không thể đặt tự khởi động: {ex.Message}");
            }
        }
    }
}
