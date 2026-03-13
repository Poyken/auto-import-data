using ExcelDataReader;
using Microsoft.Data.SqlClient; // Thư viện kết nối SQL Server mới nhất
using System.Data;
using System.Drawing;
using System.IO;
using Microsoft.Win32; // Cho phép can thiệp vào Registry để khởi động cùng Windows
using System.Runtime.InteropServices;
using Microsoft.Extensions.Configuration;

namespace ImportData
{
    // Lớp Form1: Chứa giao diện người dùng và logic chính để tự động nhập dữ liệu
    public partial class Form1 : Form
    {
        // --- 1. BIẾN THÀNH VIÊN VÀ ĐỐI TƯỢNG ĐIỀU KHIỂN ---
        
        // Bảng tạm lưu trữ dữ liệu từ Excel trước khi đưa vào Database
        private DataTable dataToImport;
        
        // Đối tượng theo dõi sự thay đổi của thư mục (thêm, sửa, đổi tên file)
        private FileSystemWatcher watcher;
        
        // Cờ đánh dấu hệ thống đang trong quá trình xử lý file, dùng để tránh xử lý trùng lặp
        private bool isProcessing = false;

        // Chuỗi kết nối tới cơ sở dữ liệu SQL Server (Giá trị mặc định)
        private string connectionString = @"Server=.;Database=CapacitorDB;Integrated Security=True;TrustServerCertificate=True;";
        
        // Đường dẫn thư mục gốc cần theo dõi để lấy file Excel
        private string baseFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "task");

        /// <summary>
        /// Nạp các thiết lập từ file cấu hình appsettings.json
        /// </summary>
        private void LoadSettings()
        {
            try
            {
                string settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
                if (File.Exists(settingsPath))
                {
                    var builder = new ConfigurationBuilder()
                        .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                        .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

                    var config = builder.Build();

                    var loadedConn = config.GetConnectionString("DefaultConnection");
                    if (!string.IsNullOrEmpty(loadedConn)) connectionString = loadedConn;

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
                Log($"Lỗi khi tải cấu hình: {ex.Message}");
            }
        }


        /// <summary>
        /// Hàm khởi tạo của Form: Thiết lập các giá trị ban đầu và đăng ký sự kiện
        /// </summary>
        public Form1()
        {
            // Khởi tạo các thành phần giao diện (được định nghĩa trong Form1.Designer.cs)
            InitializeComponent();
            
            // Đăng ký bộ mã hóa cho thư viện ExcelDataReader 
            // Cần thiết để đọc được các bảng mã tiếng Việt hoặc ký tự đặc biệt trong Excel
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Tải cấu hình từ file bên ngoài
            LoadSettings();

            // Thiết lập để ứng dụng tự động chạy khi Windows khởi động
            SetStartup();

            // Cấu hình hiển thị: Cho phép hiện ở Taskbar và trạng thái cửa sổ bình thường
            this.ShowInTaskbar = true;
            this.WindowState = FormWindowState.Normal;

            // Đăng ký sự kiện: Sau khi Form hiển thị lần đầu tiên
            this.Shown += async (s, e) => {
                Log($">>> HỆ THỐNG KHỞI CHẠY - ĐANG THEO DÕI FOLDER: {baseFolder} <<<");
                // Quét lần đầu tiên để xử lý các file cũ có sẵn trong folder
                await PerformSyncAsync(); 
                // Khởi tạo bộ theo dõi thư mục để bắt kịp các file mới phát sinh
                InitWatcher();            
            };

            // Sự kiện khi người dùng nhấn nút X tắt ứng dụng
            this.FormClosing += (s, e) => {
                if (e.CloseReason == CloseReason.UserClosing)
                {
                    // Hủy lệnh đóng: Thay vì đóng, chúng ta thu nhỏ ứng dụng xuống Taskbar
                    e.Cancel = true;
                    this.WindowState = FormWindowState.Minimized;
                    Log("Ứng dụng được thu nhỏ xuống Taskbar để tiếp tục chạy ngầm.");
                }
            };
        }

        /// <summary>
        /// Ghi nhật ký hoạt động vào danh sách hiển thị trên màn hình (ListBox)
        /// </summary>
        private void Log(string message)
        {
            // Kiểm tra nếu luồng gọi khác với luồng giao diện (UI Thread)
            if (lstLogs.InvokeRequired)
            {
                lstLogs.Invoke(new Action(() => Log(message)));
                return;
            }
            
            string time = DateTime.Now.ToString("HH:mm:ss");
            lstLogs.Items.Add($"[{time}] {message}");
            
            // Tự động cuộn xuống dòng mới nhất
            lstLogs.SelectedIndex = lstLogs.Items.Count - 1; 
            
            // Giới hạn 1000 dòng log để tránh tình trạng tốn bộ nhớ khi chạy lâu ngày
            if (lstLogs.Items.Count > 1000) lstLogs.Items.RemoveAt(0);
        }

        /// <summary>
        /// Khởi tạo đối tượng theo dõi thư mục. 
        /// Bất kỳ thay đổi nào trong folder này sẽ kích hoạt code xử lý.
        /// </summary>
        private void InitWatcher()
        {
            if (!Directory.Exists(baseFolder)) Directory.CreateDirectory(baseFolder);

            watcher = new FileSystemWatcher(baseFolder);
            watcher.IncludeSubdirectories = true; // Để bắt được các folder theo ngày như 2026-03-13
            watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite;
            watcher.Filter = "*.*";

            // Sự kiện khi có file mới hoặc file bị ghi đè/thay đổi nội dung
            watcher.Created += OnFileChanged;
            watcher.Changed += OnFileChanged;
            watcher.Renamed += OnFileChanged;

            watcher.EnableRaisingEvents = true;
        }

        /// <summary>
        /// Xử lý khi có file Excel mới được tạo ra hoặc chỉnh sửa
        /// </summary>
        private async void OnFileChanged(object sender, FileSystemEventArgs e)
        {
            // 1. Chỉ quan tâm đến file Excel (.xlsx, .xls, .xlsm)
            string ext = Path.GetExtension(e.FullPath).ToLower();
            if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm") return;

            // 2. Chỉ xử lý các file nằm trong folder tương ứng với ngày hôm nay (yyyy-MM-dd)
            string todayFolder = DateTime.Now.ToString("yyyy-MM-dd");
            if (!e.FullPath.Contains(todayFolder)) return;

            // 3. Nếu đang xử lý dở một file khác thì bỏ qua để tránh xung đột
            if (isProcessing) return;
            
            try 
            {
                isProcessing = true;
                // Nghỉ 2 giây để chắc chắn rằng file đã được hệ thống/Excel ghi xong hoàn toàn
                await Task.Delay(2000); 
                // Tiến hành đồng bộ
                await PerformSyncAsync();
            }
            finally 
            {
                // Sau khi xong (dù thành công hay lỗi) thì mở khóa cho lần xử lý tiếp theo
                isProcessing = false;
            }
        }

        private void SetStartup()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    key.SetValue("AutoImportData", Application.ExecutablePath);
                }
            }
            catch { }
        }

        private async void btnSync_Click(object sender, EventArgs e)
        {
            await PerformSyncAsync();
        }

        /// <summary>
        /// Quy trình chính: Quét folder, kiểm tra file mới và nhập vào database
        /// </summary>
        private async Task PerformSyncAsync()
        {
            // Đường dẫn folder của ngày hôm nay
            string sourceFolder = Path.Combine(baseFolder, DateTime.Now.ToString("yyyy-MM-dd"));
            
            // Nếu hôm nay chưa có folder nào được tạo thì dừng lại
            if (!Directory.Exists(sourceFolder))
            {
                lblStatus.Text = $"Chờ thư mục: {DateTime.Now:yyyy-MM-dd}";
                return;
            }

            // Hiệu ứng chuột quay tròn khi xử lý
            this.UseWaitCursor = true;
            lblStatus.Text = "Đang kiểm tra thư mục...";

            try
            {
                // Lấy danh sách toàn bộ các file Excel có trong thư mục
                string[] files = Directory.GetFiles(sourceFolder, "*.*")
                    .Where(s => s.EndsWith(".xlsx") || s.EndsWith(".xls") || s.EndsWith(".xlsm"))
                    .ToArray();

                int newlyImported = 0;

                foreach (string filePath in files)
                {
                    string fileName = Path.GetFileName(filePath);
                    
                    // KIỂM TRA: Nếu file này đã nằm trong bảng ImportHistory (đã nhập rồi) thì bỏ qua
                    if (await IsFileImported(fileName)) continue;

                    Log($"Phát hiện file mới: {fileName} - Bắt đầu xử lý...");
                    
                    // THỰC HIỆN: Đọc file Excel và Insert vào DB
                    bool success = await ProcessSingleFile(filePath);
                    
                    if (success)
                    {
                        // ĐÁNH DẤU: Lưu tên file vào lịch sử để lần sau không quét lại nữa
                        await MarkFileAsImported(fileName, filePath);
                        Log($"THÀNH CÔNG: Đã nhập dữ liệu từ file {fileName}");
                        newlyImported++;
                    }
                    else
                    {
                        Log($"THẤT BẠI: Không thể nhập dữ liệu từ file {fileName}");
                    }
                }

                // Cập nhật trạng thái tổng quát
                string summary = newlyImported > 0 ? $"Đồng bộ xong {newlyImported} file mới!" : "Không có file nào mới.";
                lblStatus.Text = summary;
                Log($">>> {summary}");
            }
            catch (Exception ex)
            {
                lblStatus.Text = "Lỗi hệ thống";
                Log($"LỖI NGHIÊM TRỌNG: {ex.Message}");
            }
            finally
            {
                this.UseWaitCursor = false;
            }
        }

        /// <summary>
        /// Kiểm tra xem một file đã từng được nhập vào database chưa
        /// </summary>
        /// <param name="fileName">Tên file cần kiểm tra</param>
        /// <returns>True nếu đã tồn tại trong lịch sử, ngược lại False</returns>
        private async Task<bool> IsFileImported(string fileName)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    await conn.OpenAsync();
                    // Truy vấn đếm số dòng có cùng tên file trong bảng ImportHistory
                    string sql = "SELECT COUNT(*) FROM ImportHistory WHERE FileName = @name";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@name", fileName);
                        int count = (int)await cmd.ExecuteScalarAsync();
                        return count > 0;
                    }
                }
            }
            catch { return false; } 
        }

        /// <summary>
        /// Ghi lại thông tin file đã nhập thành công vào bảng lịch sử
        /// </summary>
        private async Task MarkFileAsImported(string fileName, string filePath)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    await conn.OpenAsync();
                    // Lưu cả tên file, đường dẫn đầy đủ và thời gian thực hiện
                    string sql = "INSERT INTO ImportHistory (FileName, FilePath, ImportTime) VALUES (@name, @path, GETDATE())";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@name", fileName);
                        cmd.Parameters.AddWithValue("@path", filePath);
                        await cmd.ExecuteNonQueryAsync();
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Xử lý một file Excel đơn lẻ: Đọc nội dung vào bộ nhớ và sau đó đẩy dữ liệu hợp lệ vào Database.
        /// </summary>
        /// <param name="filePath">Đường dẫn đầy đủ đến file Excel cần xử lý</param>
        /// <returns>True nếu quá trình xử lý và nhập dữ liệu thành công, ngược lại False</returns>
        private async Task<bool> ProcessSingleFile(string filePath)
        {
            try
            {
                // Bước 1: Đọc dữ liệu từ file Excel vào bộ nhớ (DataTable)
                ReadExcelFile(filePath);

                if (dataToImport == null || dataToImport.Rows.Count == 0) return false;

                // Bước 2: Đẩy toàn bộ dữ liệu hợp lệ vào Database
                int importedCount = await ExecuteImportBatch();
                return importedCount > 0;
            }
            catch { return false; }
        }

        /// <summary>
        /// Thực hiện đưa dữ liệu từ DataTable vào SQL Server bằng Transaction.
        /// Chỉ các dòng được đánh dấu là "VALID" trong DataTable mới được nhập.
        /// </summary>
        /// <returns>Số dòng đã nhập thành công vào database</returns>
        private async Task<int> ExecuteImportBatch()
        {
            // Lọc ra các dòng được đánh dấu là VALID (Hợp lệ) sau khi quét Excel
            var validRows = dataToImport.AsEnumerable()
                .Where(r => r.Field<string>("Internal_Status") == "VALID")
                .ToList();

            if (validRows.Count == 0) return 0;

            int processed = 0;
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                await conn.OpenAsync();
                // Sử dụng Transaction để đảm bảo nếu lỗi ở giữa chừng thì sẽ hủy bỏ toàn bộ (tránh dữ liệu rác)
                using (SqlTransaction trans = conn.BeginTransaction())
                {
                    try
                    {
                        foreach (var row in validRows)
                        {
                            // Câu lệnh SQL Insert vào bảng CapacitorLogs
                            string sql = @"INSERT INTO CapacitorLogs (
                                EquipmentNumber, SorterNum, StartTime, WorkflowCode, LotNo, 
                                Barcode, Slot, Position, Channel, Capacity_mAh, 
                                Capacitance_F, BeginVoltageSD_mV, ChargeEndCurrent_mA, EndVoltage_mV, EndCurrent_mA, 
                                DischargeVoltage1_mV, DischargeVal1_Time, DischargeVoltage2_mV, DischargeVal2_Time, DischargeBeginVoltage_mV, 
                                DischargeBeginCurrent_mA, NGInfo, EndTime) 
                                VALUES (@c0, @c1, @c2, @c3, @c4, @c5, @c6, @c7, @c8, @c9, @c10, @c11, @c12, @c13, @c14, @c15, @c16, @c17, @c18, @c19, @c20, @c21, @c22)";

                            using (SqlCommand cmd = new SqlCommand(sql, conn, trans))
                            {
                                // Duyệt qua 23 cột dữ liệu tương ứng trong Excel
                                for (int i = 0; i < 23; i++)
                                {
                                    object val = row[i];
                                    
                                    // Xử lý chuyển đổi kiểu dữ liệu phù hợp với Database
                                    if (i == 2 || i == 22) // Cột ngày tháng
                                        cmd.Parameters.AddWithValue($"@c{i}", DateTime.TryParse(val?.ToString(), out var dt) ? dt : (object)DBNull.Value);
                                    else if (i == 7 || i == 8) // Cột số nguyên
                                        cmd.Parameters.AddWithValue($"@c{i}", int.TryParse(val?.ToString(), out var iv) ? iv : (object)DBNull.Value);
                                    else if ((i >= 9 && i <= 15) || i == 17 || i == 19 || i == 20) // Cột số thực (Double/Float)
                                        cmd.Parameters.AddWithValue($"@c{i}", double.TryParse(val?.ToString(), out var dv) ? dv : (object)DBNull.Value);
                                    else
                                        cmd.Parameters.AddWithValue($"@c{i}", val?.ToString() ?? "");
                                }
                                await cmd.ExecuteNonQueryAsync();
                            }
                            processed++;
                        }
                        trans.Commit();
                    }
                    catch { trans.Rollback(); throw; }
                }
            }
            return processed;
        }

        /// <summary>
        /// Đọc file Excel và chuyển đổi thành bảng dữ liệu (DataTable) trong bộ nhớ
        /// </summary>
        private void ReadExcelFile(string filePath)
        {
            try
            {
                // Mở luồng đọc file (Chế độ chỉ đọc để tránh xung đột nếu file đang mở bởi ứng dụng khác)
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // Cấu hình để lấy dòng đầu tiên của Excel làm tiêu đề cột (Header Row)
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        dataToImport = result.Tables[0];
                    }
                }

                // Thêm các cột kỹ thuật để quản lý trạng thái xử lý nội bộ của từng dòng
                if (!dataToImport.Columns.Contains("Internal_Status"))
                    dataToImport.Columns.Add("Internal_Status", typeof(string));
                if (!dataToImport.Columns.Contains("Reject_Reason"))
                    dataToImport.Columns.Add("Reject_Reason", typeof(string));

                // Thực hiện kiểm tra tính hợp lệ cho từng dòng dữ liệu ngay sau khi đọc
                foreach (DataRow row in dataToImport.Rows)
                {
                    ValidateRow(row);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi đọc file: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Kiểm tra dữ liệu của từng dòng trước khi nhập. 
        /// Hiện tại mặc định mọi dòng đều hợp lệ ("VALID"). Bạn có thể thêm các ràng buộc kiểm tra tại đây.
        /// </summary>
        private void ValidateRow(DataRow row)
        {
            row["Internal_Status"] = "VALID";
            row["Reject_Reason"] = "Hợp lệ";
        }


        // Các hàm cũ đã được gộp vào quy trình Sync tự động
    }
}
