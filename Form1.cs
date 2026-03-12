using ExcelDataReader;
using Microsoft.Data.SqlClient; // Thư viện kết nối SQL Server mới nhất
using System.Data;
using System.Drawing;
using System.IO;
using Microsoft.Win32; // Cho phép can thiệp vào Registry để khởi động cùng Windows
using System.Runtime.InteropServices;

namespace ImportData
{
    // Đây là file xử lý logic (Code Behind) của cái màn hình Form1.
    public partial class Form1 : Form
    {
        // 1. Biến lưu trữ dữ liệu Excel trong bộ nhớ
        private DataTable dataToImport;

        // 2. Cấu hình tự động
        private string connectionString = @"Server=.;Database=CapacitorDB;Integrated Security=True;TrustServerCertificate=True;";
        private string sourceFolder = @"c:\Users\User Vinatech.DESKTOP-RJJSEQU\Desktop\task\2026-03-11";


        public Form1()
        {
            InitializeComponent();
            
            // Đăng ký bộ mã hóa cho ExcelDataReader (Hỗ trợ .NET mới)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // 2. Tự động chạy đồng bộ ngay khi mở App
            this.Shown += (s, e) => btnSync_Click(null, null);
        }

        private async void btnSync_Click(object sender, EventArgs e)
        {
            string successFolder = Path.Combine(sourceFolder, "SUCCESS");

            if (!Directory.Exists(sourceFolder)) return;
            if (!Directory.Exists(successFolder)) Directory.CreateDirectory(successFolder);

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
                    string targetPath = Path.Combine(successFolder, fileName);

                    lblStatus.Text = $"Đang xử lý: {fileName}";
                    
                    // Xử lý Import
                    bool success = await ProcessSingleFile(filePath);
                    
                    if (success)
                    {
                        // Di chuyển file sang thư mục SUCCESS để không bị quét lại
                        if (File.Exists(targetPath)) File.Delete(targetPath); // Xóa bản cũ nếu trùng
                        File.Move(filePath, targetPath);
                        newlyImported++;
                    }
                }

                lblStatus.Text = newlyImported > 0 ? $"Đồng bộ xong {newlyImported} file!" : "Không có file mới.";
            }
            catch (Exception ex)
            {
                lblStatus.Text = "Lỗi: " + ex.Message;
            }
            finally
            {
                this.UseWaitCursor = false;
                
                // Tự động đóng ứng dụng khi đã hoàn tất công việc
                Application.Exit();
            }
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
