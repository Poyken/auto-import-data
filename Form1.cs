using ExcelDataReader;
using Microsoft.Data.SqlClient; // Thư viện kết nối SQL Server chính thức của Microsoft
using System.Data;
using Microsoft.Win32; // Cho phép can thiệp vào Registry của Windows
using Microsoft.Extensions.Configuration; // Thư viện để đọc file cấu hình JSON

namespace ImportData
{
    // Lớp xử lý chính của ứng dụng Windows Forms
    public partial class Form1 : Form
    {
        // --- 1. CÁC HẰNG SỐ CẤU HÌNH HỆ THỐNG ---
        
        // Giới hạn số dòng lịch sử hiển thị trên ListBox để tránh tốn RAM
        private const int MAX_LOG_LINES = 1000;
        
        // Độ trễ chờ cho file được giải phóng hoàn toàn từ tiến trình khác (máy đo, copy file)
        private const int FILE_RELEASE_DELAY_MS = 2000; 
        
        // Chuỗi kết nối Database mặc định nếu không nạp được từ file appsettings.json
        private const string DEFAULT_DB_CONN = @"Server=.;Database=CapacitorDB;Integrated Security=True;TrustServerCertificate=True;";
        
        // --- 2. BIẾN THÀNH VIÊN VÀ ĐỐI TƯỢNG ĐIỀU KHIỂN ---
        
        // Đối tượng thực hiện việc "lắng nghe" sự thay đổi trong thư mục
        private FileSystemWatcher watcher;
        
        // Cờ kiểm soát quá trình xử lý, ngăn chặn việc xử lý nhiều file cùng lúc gây lỗi
        private bool isProcessing = false;

        // Lưu trữ chuỗi kết nối Database hiện tại
        private string connectionString = DEFAULT_DB_CONN;
        
        // Lưu trữ đường dẫn thư mục gốc cần theo dõi (Desktop\task)
        private string baseFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "task");

        /// <summary>
        /// Nạp dữ liệu cấu hình từ file appsettings.json nằm cùng thư mục chạy ứng dụng
        /// </summary>
        private void LoadSettings()
        {
            try
            {
                // Đường dẫn đầy đủ đến file cấu hình
                string settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
                
                if (File.Exists(settingsPath))
                {
                    // Thiết lập builder để đọc file JSON
                    var builder = new ConfigurationBuilder()
                        .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                        .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

                    var config = builder.Build();

                    // Lấy chuỗi kết nối từ mục "ConnectionStrings"
                    var loadedConn = config.GetConnectionString("DefaultConnection");
                    if (!string.IsNullOrEmpty(loadedConn)) connectionString = loadedConn;

                    // Lấy đường dẫn thư mục từ mục "FolderSettings"
                    var loadedFolder = config["FolderSettings:BaseFolder"];
                    if (!string.IsNullOrWhiteSpace(loadedFolder)) baseFolder = loadedFolder;
                    
                    Log("Đã tải cấu hình từ file appsettings.json thành công.");
                }
                else
                {
                    Log("Không tìm thấy appsettings.json, sử dụng cấu hình mặc định.");
                }
            }
            catch (Exception ex)
            {
                // Ghi lại lỗi nếu quá trình nạp cấu hình thất bại
                Log($"Lỗi khi tải cấu hình: {ex.Message}");
            }
        }

        /// <summary>
        /// Hàm khởi tạo chính của Form
        /// </summary>
        public Form1()
        {
            // Khởi tạo các thành phần giao diện (định nghĩa trong Designer)
            InitializeComponent();
            
            // Đăng ký bộ Encoding để đọc được file Excel chứa mã ký tự đặc biệt/tiếng Việt
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Nạp cấu hình ứng dụng
            LoadSettings();

            // Thiết lập ứng dụng tự chạy khi bật máy tính
            SetStartup();

            // Cấu hình giao diện ban đầu
            this.ShowInTaskbar = true;
            this.WindowState = FormWindowState.Normal;

            // Đăng ký sự kiện xảy ra sau khi Form hiển thị
            this.Shown += async (s, e) => {
                Log($">>> HỆ THỐNG KHỞI CHẠY - ĐANG THEO DÕI: {baseFolder} <<<");
                
                // Đồng bộ dữ liệu lần đầu (xử lý các file đã có sẵn trong folder)
                await PerformSyncAsync(); 
                
                // Bắt đầu chế độ theo dõi thời gian thực
                InitWatcher();            
            };

            // Sự kiện khi người dùng nhấn nút đóng [X]
            this.FormClosing += (s, e) => {
                if (e.CloseReason == CloseReason.UserClosing)
                {
                    // Không đóng ứng dụng mà chỉ thu nhỏ để chạy ngầm dưới Taskbar
                    e.Cancel = true; 
                    this.WindowState = FormWindowState.Minimized;
                    Log("Ứng dụng đang chạy ngầm trên thanh Taskbar.");
                }
            };
        }

        /// <summary>
        /// Ghi log hoạt động vào ListBox hiển thị trên Form
        /// </summary>
        private void Log(string message)
        {
            // Kiểm tra nếu luồng gọi hàm không phải luồng UI chính
            if (lstLogs.InvokeRequired)
            {
                lstLogs.Invoke(new Action(() => Log(message)));
                return;
            }
            
            // Định dạng thời gian: Giờ:Phút:Giây
            string time = DateTime.Now.ToString("HH:mm:ss");
            lstLogs.Items.Add($"[{time}] {message}");
            
            // Tự động cuộn danh sách xuống dòng cuối cùng
            lstLogs.SelectedIndex = lstLogs.Items.Count - 1; 
            
            // Giới hạn số lượng dòng log hiển thị để bảo toàn bộ nhớ
            if (lstLogs.Items.Count > MAX_LOG_LINES) 
                lstLogs.Items.RemoveAt(0);
        }

        /// <summary>
        /// Cấu hình và bắt đầu theo dõi thư mục nguồn
        /// </summary>
        private void InitWatcher()
        {
            // Tự động tạo thư mục nếu chưa tồn tại
            if (!Directory.Exists(baseFolder)) Directory.CreateDirectory(baseFolder);

            // Khởi tạo đối tượng theo dõi thư mục
            watcher = new FileSystemWatcher(baseFolder);
            watcher.IncludeSubdirectories = true; // Theo dõi cả các thư mục con theo ngày
            watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite;
            watcher.Filter = "*.*";

            // Đăng ký các sự kiện thay đổi file
            watcher.Created += OnFileChanged;
            watcher.Changed += OnFileChanged;
            watcher.Renamed += OnFileChanged;

            // Bắt đầu chế độ "lắng nghe"
            watcher.EnableRaisingEvents = true;
        }

        /// <summary>
        /// Xử lý khi có file Excel mới được đưa vào thư mục theo dõi
        /// </summary>
        private async void OnFileChanged(object sender, FileSystemEventArgs e)
        {
            // Bước 1: Kiểm tra phần mở rộng file (chỉ nhận file Excel)
            string ext = Path.GetExtension(e.FullPath).ToLower();
            if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm") return;

            // Bước 2: Chỉ xử lý trong thư mục của ngày hiện tại (yyyy-MM-dd)
            string todayFolder = DateTime.Now.ToString("yyyy-MM-dd");
            if (!e.FullPath.Contains(todayFolder)) return;

            // Bước 3: Ngăn chặn xử lý chồng chéo (Race Condition)
            if (isProcessing) return;
            
            try 
            {
                isProcessing = true;
                
                // Bước 4: Chờ file được lưu hoàn tất từ các thiết bị đo (máy tính đo, copy...)
                await Task.Delay(FILE_RELEASE_DELAY_MS); 
                
                // Bước 5: Thực hiện quét và đồng bộ dữ liệu
                await PerformSyncAsync();
            }
            finally 
            {
                // Giải phóng cờ để cho phép lần xử lý tiếp theo
                isProcessing = false;
            }
        }

        /// <summary>
        /// Thiết lập Key trong Registry để ứng dụng tự khởi động cùng Windows
        /// </summary>
        private void SetStartup()
        {
            try
            {
                // Mở nhánh Registry tự chạy của người dùng hiện tại
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    // Ghi đường dẫn file .exe vào Registry
                    key.SetValue("AutoImportData", Application.ExecutablePath);
                }
            }
            catch (Exception ex)
            {
                Log($"Lỗi thiết lập khởi động cùng Windows: {ex.Message}");
            }
        }

        /// <summary>
        /// Quy trình chính: Duyệt tất cả file trong folder ngày và nhập vào cơ sở dữ liệu
        /// </summary>
        private async Task PerformSyncAsync()
        {
            // Xác định thư mục đích theo định dạng ngày: baseFolder\yyyy-MM-dd
            string sourceFolder = Path.Combine(baseFolder, DateTime.Now.ToString("yyyy-MM-dd"));
            
            // Nếu hôm nay chưa có thiết bị nào sinh file (chưa có folder) thì thoát
            if (!Directory.Exists(sourceFolder))
            {
                lblStatus.Text = $"Đang chờ thư mục: {DateTime.Now:yyyy-MM-dd}";
                return;
            }

            // Đổi con trỏ chuột sang trạng thái chờ
            this.UseWaitCursor = true;
            lblStatus.Text = "Hệ thống đang kiểm tra...";

            try
            {
                // Lấy toàn bộ danh sách file Excel đang có trong thư mục xử lý
                string[] files = Directory.GetFiles(sourceFolder, "*.*")
                    .Where(s => s.EndsWith(".xlsx") || s.EndsWith(".xls") || s.EndsWith(".xlsm"))
                    .ToArray();

                int newlyImported = 0; // Biến đếm số file nạp thành công

                foreach (string filePath in files)
                {
                    string fileName = Path.GetFileName(filePath);
                    
                    // KIỂM TRA: File này đã thực hiện nạp trước đó chưa?
                    if (await IsFileImported(fileName)) continue;

                    Log($"Phát hiện mới: {fileName} - Đang nạp dữ liệu...");
                    
                    // THỰC THI: Đọc và chèn dữ liệu vào bảng CapacitorLogs
                    bool success = await ProcessSingleFile(filePath);
                    
                    if (success)
                    {
                        // Ghi lại lịch sử nạp thành công vào bảng ImportHistory
                        await MarkFileAsImported(fileName, filePath);
                        Log($"[Xong] File {fileName} đã nạp thành công.");
                        newlyImported++;
                    }
                    else
                    {
                        Log($"[Lỗi] File {fileName} nạp thất bại.");
                    }
                }

                // Cập nhật trạng thái hiển thị trên giao diện
                string summary = newlyImported > 0 ? $"Đã hoàn thành nạp {newlyImported} file mới!" : "Dữ liệu hiện tại đã cũ (đã nạp từ trước).";
                lblStatus.Text = summary;
                Log($">>> Trạng thái: {summary}");
            }
            catch (Exception ex)
            {
                lblStatus.Text = "Lỗi nghiêm trọng";
                Log($"LỖI: {ex.Message}");
            }
            finally
            {
                // Trả con trỏ chuột về trạng thái bình thường
                this.UseWaitCursor = false;
            }
        }

        /// <summary>
        /// Kiểm tra xem tên file đã tồn tại trong lịch sử (Database) chưa
        /// </summary>
        /// <param name="fileName">Tên file cần kiểm tra</param>
        private async Task<bool> IsFileImported(string fileName)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    await conn.OpenAsync();
                    
                    // Truy vấn đếm số lượng bản ghi trùng tên file trong bảng lịch sử nạp
                    string sql = "SELECT COUNT(*) FROM ImportHistory WHERE FileName = @name";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@name", fileName);
                        int count = (int)await cmd.ExecuteScalarAsync();
                        
                        // Nếu lớn hơn 0 nghĩa là file này đã được nạp rồi
                        return count > 0;
                    }
                }
            }
            catch (Exception ex)
            { 
                Log($"Lỗi kết nối DB (Kiểm tra file): {ex.Message}");
                return false; 
            } 
        }

        /// <summary>
        /// Ghi nhận thông tin file nạp thành công vào lịch sử để tránh nạp trùng lần sau
        /// </summary>
        private async Task MarkFileAsImported(string fileName, string filePath)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    await conn.OpenAsync();
                    
                    // Lưu dấu file, đường dẫn và thời gian nạp vào bảng ImportHistory
                    string sql = "INSERT INTO ImportHistory (FileName, FilePath, ImportTime) VALUES (@name, @path, GETDATE())";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@name", fileName);
                        cmd.Parameters.AddWithValue("@path", filePath);
                        await cmd.ExecuteNonQueryAsync();
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Lỗi lưu lịch sử (Database): {ex.Message}");
            }
        }

        /// <summary>
        /// Xử lý một file Excel đơn lẻ: Đọc nội dung vào DataTable và chèn vào Database
        /// </summary>
        private async Task<bool> ProcessSingleFile(string filePath)
        {
            try
            {
                // Bước 1: Đọc nội dung Excel vào bảng tạm (DataTable)
                DataTable dt = ReadExcelFile(filePath);

                if (dt == null || dt.Rows.Count == 0) return false;

                // Bước 2: Đẩy nội dung từ bảng tạm vào SQL Server qua cơ chế Batch (theo đợt)
                int importedCount = await ExecuteImportBatch(dt);
                
                // Trả về kết quả thành công nếu nạp được ít nhất 1 dòng dữ liệu
                return importedCount > 0;
            }
            catch (Exception ex)
            {
                Log($"Lỗi xử lý luồng dữ liệu file: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Chèn dữ liệu từ bảng DataTable vào Database sử dụng Transaction (đảm bảo tính toàn vẹn)
        /// </summary>
        private async Task<int> ExecuteImportBatch(DataTable dt)
        {
            // Lọc ra danh sách các dòng dữ liệu hợp lệ (VALID)
            var validRows = dt.AsEnumerable()
                .Where(r => r.Field<string>("Internal_Status") == "VALID")
                .ToList();

            if (validRows.Count == 0) return 0;

            int processed = 0; // Biến đếm số dòng đã chèn thành công
            
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                await conn.OpenAsync();
                
                // Sử dụng Transaction: Nếu một dòng bị lỗi, toàn bộ file sẽ không được nạp (tránh rác dữ liệu)
                using (SqlTransaction trans = conn.BeginTransaction())
                {
                    try
                    {
                        foreach (var row in validRows)
                        {
                            // Câu lệnh SQL Insert vào bảng chính CapacitorLogs
                            string sql = @"INSERT INTO CapacitorLogs (
                                EquipmentNumber, SorterNum, StartTime, WorkflowCode, LotNo, 
                                Barcode, Slot, Position, Channel, Capacity_mAh, 
                                Capacitance_F, BeginVoltageSD_mV, ChargeEndCurrent_mA, EndVoltage_mV, EndCurrent_mA, 
                                DischargeVoltage1_mV, DischargeVal1_Time, DischargeVoltage2_mV, DischargeVal2_Time, DischargeBeginVoltage_mV, 
                                DischargeBeginCurrent_mA, NGInfo, EndTime) 
                                VALUES (@c0, @c1, @c2, @c3, @c4, @c5, @c6, @c7, @c8, @c9, @c10, @c11, @c12, @c13, @c14, @c15, @c16, @c17, @c18, @c19, @c20, @c21, @c22)";

                            using (SqlCommand cmd = new SqlCommand(sql, conn, trans))
                            {
                                // Ánh xạ 23 cột dữ liệu từ file Excel vào các tham số SQL ứng với tên @c0 -> @c22
                                for (int i = 0; i < 23; i++)
                                {
                                    MapParameter(cmd, i, row[i]);
                                }
                                await cmd.ExecuteNonQueryAsync();
                            }
                            processed++;
                        }
                        
                        // Nếu chạy hết mọi dòng mà không có lỗi thì mới chính thức ghi vào Disk
                        trans.Commit();
                    }
                    catch (Exception ex)
                    {
                        // Hủy bỏ toàn bộ các dòng nạp dở nếu có bất kỳ lỗi nào phát sinh
                        trans.Rollback();
                        Log($"Lỗi Transaction (Ghi dữ liệu): {ex.Message}");
                        throw; 
                    }
                }
            }
            return processed;
        }

        /// <summary>
        /// Phân tích dữ liệu từ Excel và chuyển đổi sang kiểu dữ liệu SQL phù hợp (DateTime, Int, Double...)
        /// </summary>
        private void MapParameter(SqlCommand cmd, int index, object value)
        {
            string paramName = $"@c{index}"; // Tên tham số tương ứng @c0..@c22
            string valStr = value?.ToString(); // Lấy giá trị chuỗi từ Excel

            // Các cột ngày tháng (Dòng 2 và 22)
            if (index == 2 || index == 22)
            {
                cmd.Parameters.AddWithValue(paramName, DateTime.TryParse(valStr, out var dt) ? dt : (object)DBNull.Value);
            }
            // Các cột số nguyên (Dòng 7 và 8)
            else if (index == 7 || index == 8)
            {
                cmd.Parameters.AddWithValue(paramName, int.TryParse(valStr, out var iv) ? iv : (object)DBNull.Value);
            }
            // Các cột số thực (Điện áp, Điện trở, Dung lượng...)
            else if ((index >= 9 && index <= 15) || index == 17 || index == 19 || index == 20)
            {
                cmd.Parameters.AddWithValue(paramName, double.TryParse(valStr, out var dv) ? dv : (object)DBNull.Value);
            }
            // Các cột văn bản (String)
            else
            {
                cmd.Parameters.AddWithValue(paramName, valStr ?? "");
            }
        }

        /// <summary>
        /// Dùng thư viện ExcelDataReader để nạp toàn bộ dữ liệu Excel vào một DataTable cục bộ
        /// </summary>
        private DataTable ReadExcelFile(string filePath)
        {
            try
            {
                DataTable dt;
                
                // Mở luồng đọc file Excel
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // Chuyển đổi dữ liệu Excel thành DataSet
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            // Cấu hình: lấy dòng đầu tiên của Excel làm tên tiêu đề cột
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        dt = result.Tables[0]; // Lấy sheet đầu tiên
                    }
                }

                // Thêm các cột trạng thái kỹ thuật để phần mềm tự kiểm soát nội bộ
                if (!dt.Columns.Contains("Internal_Status"))
                    dt.Columns.Add("Internal_Status", typeof(string));
                if (!dt.Columns.Contains("Reject_Reason"))
                    dt.Columns.Add("Reject_Reason", typeof(string));

                // Thực hiện kiểm duyệt từng dòng ngay sau khi đọc xong
                foreach (DataRow row in dt.Rows)
                {
                    ValidateRow(row);
                }

                return dt;
            }
            catch (Exception ex)
            {
                Log($"Lỗi khi đọc file Excel: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Logic kiểm tra tính hợp lệ cho dữ liệu: Hiện tại được cài đặt mặc định mọi dòng đều VALID.
        /// Bạn có thể bổ sung các luật kiểm tra (Ví dụ: Dung lượng không được < 0...) tại đây.
        /// </summary>
        private void ValidateRow(DataRow row)
        {
            row["Internal_Status"] = "VALID";
            row["Reject_Reason"] = "Dữ liệu hợp lệ";
        }

        // --- HẾT PHẦN LOGIC XỬ LÝ ---
    }
}
