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
    // Đây là file xử lý logic (Code Behind) của cái màn hình Form1.
    public partial class Form1 : Form
    {
        // 1. Biến lưu trữ dữ liệu Excel trong bộ nhớ
        private DataTable dataToImport;
        private FileSystemWatcher watcher;
        private bool isProcessing = false;

        // 2. Cấu hình (Sẽ được load từ appsettings.json)
        private string connectionString = @"Server=.;Database=CapacitorDB;Integrated Security=True;TrustServerCertificate=True;";
        private string baseFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "task");

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


        public Form1()
        {
            InitializeComponent();
            
            // Đăng ký bộ mã hóa cho ExcelDataReader (Hỗ trợ .NET mới)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Tải cấu hình bên ngoài
            LoadSettings();

            // 1. Cấu hình tự khởi động cùng Windows
            SetStartup();

            // 2. Thiết lập hiển thị Dashboard chuyên nghiệp
            this.ShowInTaskbar = true;
            this.WindowState = FormWindowState.Normal;

            // 3. Tự động quét một lần khi mở App, sau đó chuyển sang chế độ theo dõi
            this.Shown += async (s, e) => {
                Log($">>> HỆ THỐNG KHỞI CHẠY - ĐANG THEO DÕI FOLDER: {baseFolder} <<<");
                await PerformSyncAsync(); // Quét lần đầu để không sót file cũ
                InitWatcher();            // Bắt đầu theo dõi thay đổi thực tế
            };

            // 4. Ngăn việc tắt ứng dụng
            this.FormClosing += (s, e) => {
                if (e.CloseReason == CloseReason.UserClosing)
                {
                    e.Cancel = true;
                    this.WindowState = FormWindowState.Minimized;
                    Log("Ứng dụng được thu nhỏ xuống Taskbar để tiếp tục chạy ngầm.");
                }
            };
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
            lstLogs.SelectedIndex = lstLogs.Items.Count - 1; // Cuộn xuống dòng cuối
            
            // Giới hạn 1000 dòng log để tránh tràn bộ nhớ
            if (lstLogs.Items.Count > 1000) lstLogs.Items.RemoveAt(0);
        }

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

        private async void OnFileChanged(object sender, FileSystemEventArgs e)
        {
            // Kiểm tra định dạng file
            string ext = Path.GetExtension(e.FullPath).ToLower();
            if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm") return;

            // Kiểm tra xem có phải file nằm trong folder ngày hôm nay không
            string todayFolder = DateTime.Now.ToString("yyyy-MM-dd");
            if (!e.FullPath.Contains(todayFolder)) return;

            // Tránh việc nhiều sự kiện bắn ra cùng lúc cho cùng 1 file (Debounce)
            if (isProcessing) return;
            
            try 
            {
                isProcessing = true;
                // Đợi một chút để file kịp giải phóng khỏi Excel hoặc quá trình copy
                await Task.Delay(2000); 
                await PerformSyncAsync();
            }
            finally 
            {
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

        private async Task PerformSyncAsync()
        {
            string sourceFolder = Path.Combine(baseFolder, DateTime.Now.ToString("yyyy-MM-dd"));
            
            if (!Directory.Exists(sourceFolder))
            {
                lblStatus.Text = $"Chờ thư mục: {DateTime.Now:yyyy-MM-dd}";
                return; // Chỉ thoát khỏi hàm quét, không thoát App
            }

            this.UseWaitCursor = true;
            lblStatus.Text = "Đang kiểm tra thư mục...";

            try
            {
                string[] files = Directory.GetFiles(sourceFolder, "*.*")
                    .Where(s => s.EndsWith(".xlsx") || s.EndsWith(".xls") || s.EndsWith(".xlsm"))
                    .ToArray();

                int newlyImported = 0;

                foreach (string filePath in files)
                {
                    string fileName = Path.GetFileName(filePath);
                    
                    // 1. Kiểm tra xem file này đã từng được Import chưa (dựa trên bảng ImportHistory trong DB)
                    if (await IsFileImported(fileName)) continue;

                    Log($"Phát hiện file mới: {fileName} - Bắt đầu xử lý...");
                    
                    // 2. Xử lý Import dữ liệu
                    bool success = await ProcessSingleFile(filePath);
                    
                    if (success)
                    {
                        // 3. Ghi lại lịch sử kèm theo đường dẫn file
                        await MarkFileAsImported(fileName, filePath);
                        Log($"THÀNH CÔNG: Đã nhập dữ liệu từ file {fileName}");
                        newlyImported++;
                    }
                    else
                    {
                        Log($"THẤT BẠI: Không thể nhập dữ liệu từ file {fileName}");
                    }
                }

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
                // Không đóng ứng dụng nữa, để vòng lặp RunSyncCycle tiếp tục
            }
        }

        private async Task<bool> IsFileImported(string fileName)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    await conn.OpenAsync();
                    string sql = "SELECT COUNT(*) FROM ImportHistory WHERE FileName = @name";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@name", fileName);
                        int count = (int)await cmd.ExecuteScalarAsync();
                        return count > 0;
                    }
                }
            }
            catch { return false; } // Nếu DB chưa có bảng thì mặc định coi như chưa import
        }

        private async Task MarkFileAsImported(string fileName, string filePath)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    await conn.OpenAsync();
                    // Cập nhật lưu thêm đường dẫn và ngày giờ import
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

        // Xóa mã liên quan đến Table History cũ không còn dùng
        private async Task<bool> ProcessSingleFile(string filePath)
        {
            try
            {
                // Đọc file vào dataToImport
                ReadExcelFile(filePath);

                if (dataToImport == null || dataToImport.Rows.Count == 0) return false;

                // Tận dụng hàm Import sẵn có (đã được refactor một chút để nhận tham số hoặc dùng biến class)
                int importedCount = await ExecuteImportBatch();
                return importedCount > 0;
            }
            catch { return false; }
        }

        private async Task<int> ExecuteImportBatch()
        {
            // Tách phần logic Import từ btnImport_Click ra đây để dùng chung
            var validRows = dataToImport.AsEnumerable()
                .Where(r => r.Field<string>("Internal_Status") == "VALID")
                .ToList();

            if (validRows.Count == 0) return 0;

            int processed = 0;
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                await conn.OpenAsync();
                using (SqlTransaction trans = conn.BeginTransaction())
                {
                    try
                    {
                        foreach (var row in validRows)
                        {
                            string sql = @"INSERT INTO CapacitorLogs (
                                EquipmentNumber, SorterNum, StartTime, WorkflowCode, LotNo, 
                                Barcode, Slot, Position, Channel, Capacity_mAh, 
                                Capacitance_F, BeginVoltageSD_mV, ChargeEndCurrent_mA, EndVoltage_mV, EndCurrent_mA, 
                                DischargeVoltage1_mV, DischargeVal1_Time, DischargeVoltage2_mV, DischargeVal2_Time, DischargeBeginVoltage_mV, 
                                DischargeBeginCurrent_mA, NGInfo, EndTime) 
                                VALUES (@c0, @c1, @c2, @c3, @c4, @c5, @c6, @c7, @c8, @c9, @c10, @c11, @c12, @c13, @c14, @c15, @c16, @c17, @c18, @c19, @c20, @c21, @c22)";

                            using (SqlCommand cmd = new SqlCommand(sql, conn, trans))
                            {
                                for (int i = 0; i < 23; i++)
                                {
                                    object val = row[i];
                                    if (i == 2 || i == 22) // Date columns
                                        cmd.Parameters.AddWithValue($"@c{i}", DateTime.TryParse(val?.ToString(), out var dt) ? dt : (object)DBNull.Value);
                                    else if (i == 7 || i == 8) // Int columns
                                        cmd.Parameters.AddWithValue($"@c{i}", int.TryParse(val?.ToString(), out var iv) ? iv : (object)DBNull.Value);
                                    else if ((i >= 9 && i <= 15) || i == 17 || i == 19 || i == 20) // Float columns
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

        private void ReadExcelFile(string filePath)
        {
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        dataToImport = result.Tables[0];
                    }
                }

                // Thêm các cột trạng thái nếu chưa có
                if (!dataToImport.Columns.Contains("Internal_Status"))
                    dataToImport.Columns.Add("Internal_Status", typeof(string));
                if (!dataToImport.Columns.Contains("Reject_Reason"))
                    dataToImport.Columns.Add("Reject_Reason", typeof(string));

                // Thực hiện Validation tự động
                foreach (DataRow row in dataToImport.Rows)
                {
                    ValidateRow(row);
                }

                // Quy trình xử lý ngầm hoàn tất
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi đọc file: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ValidateRow(DataRow row)
        {
            // Hiện tại không còn cột nào bắt buộc phải có dữ liệu.
            // Mọi dòng đều được coi là hợp lệ (VALID) trừ khi bạn thêm quy tắc mới.
            row["Internal_Status"] = "VALID";
            row["Reject_Reason"] = "Hợp lệ";
        }


        // Các hàm cũ đã được gộp vào quy trình Sync tự động
    }
}
