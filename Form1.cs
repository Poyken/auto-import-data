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
    /// <summary>
    /// Giao diện chính (Form1): Đóng vai trò là trung tâm điều khiển của toàn bộ phần mềm.
    /// Nó quản lý các dịch vụ (Database, Excel) và theo dõi thư mục để tự động nạp dữ liệu.
    /// </summary>
    public partial class Form1 : Form
    {
        private const int MAX_LOG_LINES = 1000; 
        
        private readonly AppConfig _config;
        private DatabaseService _dbService;
        private ExcelService _excelService;
        
        private FileSystemWatcher _watcher;
        private bool _isProcessing;
        
        private System.Windows.Forms.Timer _healthTimer;
        private string _lastState = "";
        private bool _isSystemHealthy = false;

        private NotifyIcon _trayIcon;

        public Form1()
        {
            InitializeComponent();
            
            lstLogs.DrawMode = DrawMode.OwnerDrawFixed;
            lstLogs.DrawItem += LstLogs_DrawItem; 
            
            _config = new AppConfig();
            _config.Load(null);

            _dbService = new DatabaseService(_config, Log); // Khởi tạo dịch vụ Database.
            _excelService = new ExcelService(Log);          // Khởi tạo dịch vụ Excel.

            // Yêu cầu Windows: Mỗi khi bật máy tính lên, hãy tự động chạy phần mềm này ẩn ở dưới nền.
            SystemHelper.SetStartup(Log);

            // Cài đặt biểu tượng nhỏ chạy dưới góc khay đồng hồ.
            _trayIcon = new NotifyIcon()
            {
                Icon = SystemIcons.Information,
                Text = "Import Data (Running)",
                Visible = true
            };

            // Bắt sự kiện: Nếu người dùng nhấp đúp máy chuột vào biểu tượng thì mở to giao diện phần mềm ra.
            _trayIcon.DoubleClick += (s, e) => {
                this.Show(); 
                this.WindowState = FormWindowState.Normal; 
            };

            this.ShowInTaskbar = true; // Hiện phần mềm trên thanh ngang dưới cùng của Windows.
            this.WindowState = FormWindowState.Normal; 

            // Cài đặt sự kiện: Ngay khi bảng Form1 vừa hiện lên màn hình thì chạy hàm "Form1_Shown".
            this.Shown += Form1_Shown; 
            
            // Cài đặt sự kiện: Khi người dùng bấm dấu X tắt app thì chạy hàm "Form1_FormClosing".
            this.FormClosing += Form1_FormClosing;  
        }

        // Hàm này tự động chạy ngay sau khi phần mềm mở lên. Chữ 'async' giúp app không bị đơ khi thực thi lệnh bên trong.
        private async void Form1_Shown(object sender, EventArgs e)
        {
            Log($"Monitoring: {_config.BaseFolder}");
            
            _healthTimer = new System.Windows.Forms.Timer(); 
            _healthTimer.Interval = 10000; // Đặt thời gian kiểm tra định kỳ là 10.000 milliseconds (10 giây).
            
            // Cứ mỗi 10 giây, đồng hồ kêu Tick, ta gọi hàm PerformHealthCheckLoopAsync để khám sức khỏe tổng chẩn.
            _healthTimer.Tick += async (s, ev) => await PerformHealthCheckLoopAsync(); 
            
            await PerformHealthCheckLoopAsync(); // Chạy khám sức khỏe lần đầu tiên ngay lúc này.
            _healthTimer.Start(); // Khởi động đồng hồ.
        }

        /// <summary>
        /// Hàm khám sức khỏe hệ thống: Chạy ngầm mỗi 10 giây.
        /// Kiểm tra xem người dùng có đổi cấu hình file không, đường dẫn thư mục còn tồn tại không, và SQL có kết nối được không.
        /// </summary>
        private async Task PerformHealthCheckLoopAsync()
        {
            string oldFolder = _config.BaseFolder; // Lưu lại thư mục cũ để xíu nữa so sánh xem có đổi thư mục không.

            // Đọc lại file appsettings.json. Điểm hay là người dùng sửa file txt, app lập tức cập nhật mà không cần khởi động lại.
            _config.Load(null); 

            string currentState = "HEALTHY"; // Chuẩn bị một biến gán trạng thái Tốt (HEALTHY).
            bool currentPathOk = true; 
            bool currentDbOk = true;   

            // 1. Khám bệnh thư mục gốc: Xem thư mục quét của máy đo còn tồn tại trên ổ cứng không? Cẩn thận lỡ ai xóa nhầm.
            if (!Directory.Exists(_config.BaseFolder))
            {
                currentPathOk = false; 
                currentState = "PATH_ERROR"; 
            }

            // 2. Khám bệnh mạng SQL: Nếu thư mục ổn, ta mới đi thử mở cửa cơ sở dữ liệu xem có thành công không.
            if (currentPathOk)
            {
                currentDbOk = await _dbService.TestConnectionAsync();
                if (!currentDbOk) currentState = "DB_ERROR"; 
            }

            // Kết luận tổng thể: Cấu hình chung là 'Khỏe' khi cả đường dẫn Tốt VÀ SQL bắt tay Tốt.
            _isSystemHealthy = currentPathOk && currentDbOk;

            // Kịch bản 1: Nếu ứng dụng Khỏe, mà người dùng lại vừa vào appsettings.json đổi tên thư mục Máy Đo sang thư mục khác.
            if (_isSystemHealthy && oldFolder != _config.BaseFolder)
            {
                Log($"Monitoring folder: {_config.BaseFolder}"); 
                InitWatcher(); 
                await PerformSyncAsync(); 
            }

            // Kịch bản 2: Báo Lỗi hoặc Phục Hồi. So sánh trạng thái lỗi hiện tại với 10 giây trước đó. 
            // Ta chỉ in ra màn hình khi có sự thay đổi bệnh tật, tránh việc cứ 10s lại spam 1 dòng báo lỗi giống nhau lên ứng dụng.
            if (currentState != _lastState)
            {
                if (_isSystemHealthy) 
                {
                    UpdateStatus("Hệ thống Sẵn sàng", Color.Green);
                    
                    InitWatcher(); 
                    await PerformSyncAsync(); 
                }
                else 
                {
                    // Lỗi gì đó. Dừng theo dõi bằng Watcher để tiết kiệm CPU cho máy.
                    StopWatcher(); 
                    if (!currentPathOk) 
                    {
                        Log($"[ERROR] Path not found: {_config.BaseFolder}"); 
                        UpdateStatus("Lỗi đường dẫn", Color.Red); 
                    }
                    else if (!currentDbOk) 
                    {
                        Log("[ERROR] Database connection failed.");
                        UpdateStatus("Lỗi kết nối SQL", Color.Red);
                    }
                }
                _lastState = currentState; // Ghi nhớ lại tình trạng lỗi cho chu kỳ 10 giây mẻ sau.
            }
        }

        // Hàm tắt mắt theo dõi. Giải phóng hệ thống quan sát OS.
        private void StopWatcher()
        {
            if (_watcher != null) 
            {
                _watcher.EnableRaisingEvents = false; // Ngưng kích hoạt báo động.
                _watcher.Dispose(); // Gỡ bỏ cảm biến khỏi hệ thống RAM để không nặng Windows.
                _watcher = null; 
            }
        }

        // Sự kiện: Khi người dùng thao tác bấm dấu [X] đỏ góc trên ứng dụng.
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Kiểm tra xem chính người dùng đang chủ động thao tác.
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true; // Chặn lệnh tắt của Windows. Hủy diệt phần mềm là không hợp lệ!
                this.Hide();     // Ẩn bảng màn hình ứng dụng đi, che mất khỏi thanh công việc.
                
                // Hiển thị một bóng thông báo nhỏ xíu góc Màn góc đồng hồ 3 giây. Cho anh em biết App vẫn còn sống ngầm.
                _trayIcon.ShowBalloonTip(2000, "Import Data", "Ứng dụng đang chạy ngầm.", ToolTipIcon.Info); 
            }
        }

        // Hàm cập nhật dòng chữ Trạng Thái (Màu vàng, đỏ, xanh).
        private void UpdateStatus(string message, Color color)
        {
            // Bảo vệ đa luồng: Do Timer thường đẻ ra một nhánh công việc mới, nếu nhánh này cố vẽ chữ lên Form của nhánh chính sẽ bị Windows báo Lỗi văng.
            // Lệnh Invoke giúp xin phép Form được vẽ chữ hợp pháp mà không sợ treo app (Cross-Thread Exception).
            if (this.InvokeRequired) 
            {
                this.Invoke(new Action(() => UpdateStatus(message, color))); 
                return;
            }
            lblStatus.Text = message; 
            lblStatus.ForeColor = color; 
        }

        // Tự tay vẽ màu cho bảng Log chữ.
        private void LstLogs_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            string text = lstLogs.Items[e.Index].ToString(); // Bốc dòng chữ tại vị trí dòng số mấy ra.
            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

            // --- PHẦN 1: VẼ NỀN ---
            if (isSelected)
            {
                // Màu nền khi được chọn (Focus): Dùng màu Xám Slate tối (Gần với Đen) để làm nổi bật dòng mà không gây chói.
                using (SolidBrush backBrush = new SolidBrush(Color.FromArgb(40, 40, 45)))
                {
                    e.Graphics.FillRectangle(backBrush, e.Bounds);
                }
                
                // Vẽ một "Thanh Nhấn" (Accent Bar) màu xanh sáng ở mép trái để nhận diện dòng đang chọn.
                using (SolidBrush accentBrush = new SolidBrush(Color.FromArgb(0, 255, 255)))
                {
                    e.Graphics.FillRectangle(accentBrush, e.Bounds.X, e.Bounds.Y, 4, e.Bounds.Height);
                }
            }
            else
            {
                // Màu nền mặc định (Đen sâu thẳm).
                e.Graphics.FillRectangle(Brushes.Black, e.Bounds);
            }

            // --- PHẦN 2: QUYẾT ĐỊNH MÀU CHỮ ---
            // GIỮ NGUYÊN MÀU GỐC: Không đổi sang màu Trắng khi focus để đảm bảo dữ liệu quan trọng ([LỖI], [OK]) luôn dễ nhận biết.
            Brush textBrush = Brushes.LimeGreen; // Mặc định dùng Xanh Lime.

            // Tìm từ khóa để đổi màu nhấn mạnh theo đúng ý nghĩa thông báo.
            if (text.Contains("[LỖI")) textBrush = Brushes.Red;
            else if (text.Contains("[OK]")) textBrush = Brushes.Cyan;
            else if (text.Contains("[THAY ĐỔI") || text.Contains("Phát hiện mới")) textBrush = Brushes.Yellow;

            // --- PHẦN 3: VẼ CHỮ ---
            // Căn lề trái cách ra một chút (X + 8) để không bị đè lên Thanh Nhấn.
            Rectangle textBounds = new Rectangle(e.Bounds.X + 8, e.Bounds.Y, e.Bounds.Width - 8, e.Bounds.Height);
            
            // Sử dụng StringFormat để căn giữa chữ theo chiều dọc cho cân đối.
            using (StringFormat sf = new StringFormat { LineAlignment = StringAlignment.Center })
            {
                e.Graphics.DrawString(text, e.Font, textBrush, textBounds, sf);
            }
        }

        // Hàm Ghi Log, đẩy dòng chữ mới vào danh sách ListBox.
        private void Log(string message)
        {
            if (lstLogs.InvokeRequired) // Kiểm soát đa luồng giống hàm UpdateStatus ở trên.
            {
                lstLogs.Invoke(new Action(() => Log(message))); 
                return;
            }
            
            string time = DateTime.Now.ToString("HH:mm:ss"); // Cắt gọn giờ phút biểu trưng.
            lstLogs.Items.Add($"[{time}] {message}"); 
            
            // Kéo thanh cuộn chuột chạy tuột xuống dòng mới sinh ra để khách dễ đọc tin tức cuối.
            lstLogs.SelectedIndex = lstLogs.Items.Count - 1; 
            
            // Thuật toán cuộn vòng (Circular buffer): Tránh tràn bộ nhớ
            // Nếu Log vọt qua 1000 dòng, ta sẽ tiến hành cứt cái dòng trên cùng rác cũ bỏ đi. Giữ nguyên 1000 dòng.
            if (lstLogs.Items.Count > MAX_LOG_LINES) 
                lstLogs.Items.RemoveAt(0); 
        }

        // Hàm kích hoạt cắm cảm biến trực tiếp vào hệ thống File của HDH Windows.
        private void InitWatcher()
        {
            // Tránh thư mục lõi mất thì Windows sẽ bắn lỗi với Watcher. Create Directory bảo toàn trước.
            if (!Directory.Exists(_config.BaseFolder)) 
            {
                Directory.CreateDirectory(_config.BaseFolder); 
            }

            _watcher = new FileSystemWatcher(_config.BaseFolder) 
            {
                IncludeSubdirectories = true, // Quét luôn cả những thư mục con bên trong nó (Ví dụ thư mục ngày tháng).
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite, // Đảo bảo chỉ nhận cảnh báo khi có Tên Mới sinh, và khi nó Vừa ghi file dữ liệu thêm vào.
                Filter = "*.*", // Lọc tất cả mọi loại thể file tên không chừa ai cả.
                EnableRaisingEvents = true // Khởi động nút kích hoạt nghe nhìn.
            };

            // Gắn hành động xử trí OnFileChanged (hàm bên dưới) ứng nghiệm khi có các việc này:
            _watcher.Created += OnFileChanged; // Vừa sinh tệp con.
            _watcher.Changed += OnFileChanged; // Đang thay chữ cập nhật file này.
            _watcher.Renamed += OnFileChanged; // Sửa tên tệp của nó.
        }

        // Hàm nhận tính hiệu báo động từ O.S và nhảy vào lấy file về.
        private async void OnFileChanged(object sender, FileSystemEventArgs e)
        {
            // Nếu hệ thống đang đứt SQL lỗi thì báo động cũng phớt lờ, vì cố mở ra cũng vô tác dụng, không đổ data đi được.
            if (!_isSystemHealthy) return; 

            // Cắt đuôi loại file: Nhận chỉ định xlsx, hay xls... Tránh rác pdf, ảnh img bị quăng lầm vào kho.
            string ext = Path.GetExtension(e.FullPath).ToLower();
            if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm") return;

            // Xử lý bộ máy đo lường hay ngốc nghếch viết đè dòng ngày hôm qua của dự án cũ. 
            // Cắt ra chỉ theo dõi thư mục yyyy-MM-dd là ĐÚNG ngày hôm Nay. Còn không phải thì Không cất công lọc nó.
            string todayFolder = DateTime.Now.ToString("yyyy-MM-dd");
            string folderName = Path.GetFileName(Path.GetDirectoryName(e.FullPath));
            if (folderName != todayFolder) return;

            // Biến bảo vệ Cờ _isProcessing chống hiện tượng Race Condition (Tranh đồ ăn).
            // Do máy đo sẽ nổ phát nhiều event Changed liên tục 1 mili giây. Phải khóa lại để Hàm xử lý an tâm nuốt trọn xong 1 cục tệp.
            if (_isProcessing) return; 
            
            try 
            {
                _isProcessing = true; // Khóa cửa chặn luồng. Không cho bất kỳ ai báo động khác làm phiền nhánh quét này.
                
                // Kỹ thuật "Chờ File rảnh" - Máy đo đang ghi data vào file O.S, hệ thống bị nghẽn Lock. 
                // Ta đứng chờ gõ cửa. Mở được cửa file thì chèn quét gom bộ Đồng vớt (Sync).
                if (await WaitForFileReadyAsync(e.FullPath)) 
                {
                    await PerformSyncAsync(); 
                }
            }
            finally 
            {
                // finally chắc chắn cờ sẽ được mở khóa xả ra DÙ cho bên trong Try có văng Error Exception nổ đi chăng nữa. Tránh kẹt App vĩnh cửu.
                _isProcessing = false; 
            }
        }

        // Hàm đàm phán "Khoảng lùi": Lịch sự chờ đợi Tệp máy đo buông tay khỏi hệ điều hành, tránh ghi lật gây hư File.
        private async Task<bool> WaitForFileReadyAsync(string filePath)
        {
            for (int i = 0; i < 5; i++) // Thử gõ cửa 5 lượt.
            {
                try
                {
                    // Lệnh quyền FileShare.None (Không cho ai đụng). Ép xin HĐH là người chủ độc quyền của File. 
                    // Nếu máy Đo đang ghi, Windows sẽ báo Lỗi (IOException). Lọt được lệnh này tức là file Đã Ghi Hoàn Toàn Xong Rảnh Tay. Nhả True!
                    using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        return true; 
                    }
                }
                catch (IOException) // Tường lửa O.S đánh văng vì Máy đo đang rặn Ghi xuất.
                {
                    await Task.Delay(1000); // Lùi bộ đếm chờ máy đo xuất hết thư thả, 1000mS (1 giây) sau hàm For vòng lặp lên gõ cửa tiếp tục.
                }
            }
            return false; // Đợi 5 giây (Tức 5 lần cản bộ gõ) mà vẫn khóa, từ bỏ luôn! Chắc file bị lỗi chết nên treo.
        }

        // Đội quân thu vén Tàn tệp Đồng bộ (Nạp vào lúc mới mở Màn Hình phần mềm, hay sau khi có cảnh báo File Rảnh O.S).
        private async Task PerformSyncAsync()
        {
            if (!_isSystemHealthy) return; 

            // Cấu trúc chuỗi dẫn nối sâu vô trong thư mục Hữu hiệu Của Hiện Tại.
            string todayFolder = DateTime.Now.ToString("yyyy-MM-dd");
            string sourceFolder = Path.Combine(_config.BaseFolder, todayFolder);
            
            // Xóa sổ, ngày hôm nay chưa mở máy Lập trình Đo đạc thì thư mục chắc chắn chưa tồn. Ngưng quét Tốn Nhịp.
            if (!Directory.Exists(sourceFolder)) return; 

            UpdateStatus("Đang đồng bộ...", Color.Yellow); 

            try
            {
                // Quét rải rác thu lấy toàn bộ File Mọi Kiểu (*.*) ở tận hang hốc thư mục, sau đó dùng bộ Lưới Linq lọc giặt (Where) ra riêng loại tệp chữ Excel rớt gọn cho Mảng list.
                string[] files = Directory.GetFiles(sourceFolder, "*.*")
                    .Where(s => s.EndsWith(".xlsx") || s.EndsWith(".xls") || s.EndsWith(".xlsm"))
                    .ToArray();

                foreach (string filePath in files) // Giải trình quét 1 mẻ các loại file List hộp mới vừa gắp.
                {
                    string fileName = Path.GetFileName(filePath); // Cắt giật đuôi dẫn thư đằng trước, lấy đúng tên cục trơn "Result_1.xlsx" cho đẹp Data DB.
                    
                    // Nhắc Thợ Database Hỏi Trong Sổ Lịch SQL có cái Tên file nào Trùng chưa?
                    // Trùng (True) thì chạy câu lệnh (continue) -> Lờ Bỏ nó đi lấy Tệp Tên Kế Tiếp vòng lại. Tiết kiệm không phải Load Lại tệp Xã Hội rác Trùng Cũ!
                    if (await _dbService.IsFileImportedAsync(fileName)) continue; 

                    Log($"Importing: {fileName}");
                    
                    bool success = await ProcessSingleFileAsync(filePath, fileName);
                    
                    if (success) Log($"[OK] {fileName}");
                    else Log($"[ERROR] {fileName} failed.");
                }
                
                UpdateStatus("Hệ thống Sẵn sàng", Color.Green); 
            }
            catch (Exception)
            {
                UpdateStatus("Lỗi đồng bộ", Color.Red); 
                Log("[ERROR] Sync failed."); 
            }
        }

        // Tác Vụ Đặc Biệt Nhất: Bắt Trọn Giải Cứu Kết Nối Excel Ram Đẩy Cho Chở SQL Thả.
        private async Task<bool> ProcessSingleFileAsync(string filePath, string fileName)
        {
            try
            {
                // Triệu hồi ExcelService: Nuốt chửng nguyên file vật lý châm thành 1 bộ Data Lưới Table Phẳng lì đẹp mắt!.
                var dataTable = _excelService.ReadExcelFile(filePath);
                
                // Trường hợp file khuyết Null, Mất Nàng hay rỗng không hàng cột nào. Dừng Cuộc Chơi Cho Nó Bãi!
                if (dataTable == null || dataTable.Rows.Count == 0) return false; 
  
                // Triệu hồi Thợ SQL (DatabaseService): Gấp bộ DataTable vuông vức này chuyển nhồi Tống vào miệng rãnh BulkCopy nạp tốc độ hàng giây!
                int rowsImported = await _dbService.ExecuteImportBatchAsync(dataTable, fileName, filePath);
                
                // Hồi phản số dòng (Ví dụ trả 10). Tức là > 0 thì Trình Báo Vui Lên Sếp Cờ Hàm Rằng Bọt Đã Đổ True Lên Ngon Lành Cành Quất Rụng Ngay!.
                return rowsImported > 0; 
            }
            catch (Exception)
            {
                Log($"[ERROR] Processing {fileName} failed."); 
                return false;
            }
        }
    }
}
