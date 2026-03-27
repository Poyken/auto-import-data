using Microsoft.Data.SqlClient;
using System;
using System.IO;
using System.Data;
using System.Threading.Tasks;
using ImportData.Core;

namespace ImportData.Services 
{
    /// <summary>
    /// Lớp DatabaseService: Phụ trách toàn bộ các thao tác với Cơ sở dữ liệu SQL Server.
    /// Bao gồm: Kiểm tra kết nối, Kiểm tra trùng lặp và Nhập dữ liệu hàng loạt tốc độ cao.
    /// </summary>
    public class DatabaseService
    {
        // Tên bảng chứa dữ liệu đo đạc.
        private const string TableData = "SortingDataImportExcel_test";
        // Tên bảng ghi lại lịch sử các file đã nạp (để tránh nạp trùng).
        private const string TableHistory = "ExcelImportHistory_test";

        // Danh sách ánh xạ các cột từ Excel sang SQL để App biết "Cái gì bỏ vào đâu?".
        private static readonly string[] SqlColumns = {
            "EquipmentNumber", "SorterNum", "StartTime", "WorkflowCode", "LotNo",
            "Barcode", "Slot", "Position", "Channel", "Capacity_mAh", "Capacitance_F", 
            "BeginVoltageSD_mV", "ChargeEndCurrent_mA", "EndVoltage_mV", "EndCurrent_mA", "DischargeVoltage1_mV", 
            "DischargeVoltage1_Time", "DischargeVoltage2_mV", "DischargeVoltage2_Time", "DischargeBeginVoltage_mV", "DischargeBeginCurrent_mA", 
            "NGInfo", "EndTime", "FilePath", "ImportDate"
        };

        private readonly AppConfig _config;   // Cấu hình (ConnectionString).
        private readonly Action<string> _logger; // Hàm để in nhật ký ra màn hình.
        private string _lastConnectionString; // Ghi nhớ chuỗi kết nối cũ để xử lý thay đổi server.

        // Hàm khởi tạo DatabaseService.
        public DatabaseService(AppConfig config, Action<string> logger)
        {
            _config = config; 
            _logger = logger; 
            _lastConnectionString = config.ConnectionString; 
        }

        /// <summary>
        /// Kiểm tra xem Server SQL có đang bật và cho phép kết nối không.
        /// </summary>
        public async Task<bool> TestConnectionAsync()
        {
            try 
            {
                // Nếu người dùng vừa đổi server trong file cấu hình, ta xóa cache kết nối cũ.
                if (_config.ConnectionString != _lastConnectionString)
                {
                    SqlConnection.ClearAllPools();
                    _lastConnectionString = _config.ConnectionString;
                }
                
                // Thiết lập thời gian chờ tối đa (Timeout) khi thử kết nối.
                var builder = new SqlConnectionStringBuilder(_config.ConnectionString) 
                { 
                    ConnectTimeout = _config.HealthCheckTimeoutSeconds 
                };
                
                // Mở thử 1 kết nối tạm thời đến SQL Server.
                using (var testConn = new SqlConnection(builder.ConnectionString)) 
                {
                    await testConn.OpenAsync(); // Mở kết nối bất đồng bộ.
                    return true; // Kết nối thành công.
                } 
            }
            catch (Exception)
            {
                // Lỗi kết nối sẽ được Form1 báo lại cho người dùng bằng thông báo màu đỏ.
                return false;
            }
        }

        /// <summary>
        /// Kiểm tra xem tệp Excel này đã được nạp thành công vào hệ thống trước đó chưa.
        /// </summary>
        public async Task<bool> IsFileImportedAsync(string filePath) 
        {
            try
            {
                using (var conn = new SqlConnection(_config.ConnectionString)) 
                {
                    await conn.OpenAsync(); 
                    // Câu lệnh SQL: Đếm số dòng có đường dẫn tệp trùng khớp và trạng thái thành công.
                    string sql = $"SELECT COUNT(*) FROM {TableHistory} WHERE FilePath = @path AND Status = 'Success'";
                    
                    using (var cmd = new SqlCommand(sql, conn)) 
                    {
                        cmd.Parameters.AddWithValue("@path", filePath); // Truyền tham số an toàn chống SQL Injection.
                        int count = (int)await cmd.ExecuteScalarAsync(); 
                        return count > 0; // Nếu lớn hơn 0 tức là tệp đã nạp rồi.
                    }
                } 
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[LỖI-SQL] Không kiểm tra được lịch sử nạp file: {ex.Message}");
                return false; 
            }
        }

        /// <summary>
        /// Nhập hàng ngàn dòng dữ liệu từ DataTable vào SQL Server bằng công nghệ SqlBulkCopy (Cực nhanh).
        /// </summary>
        public async Task<int> ExecuteImportBatchAsync(DataTable dt, string fileName, string filePath) 
        {
            if (dt == null || dt.Rows.Count == 0) return 0; // Nếu dữ liệu rỗng thì thoát.

            using (var conn = new SqlConnection(_config.ConnectionString)) 
            {
                await conn.OpenAsync(); 
                // Sử dụng Transaction (Giao dịch): Đảm bảo hoặc nạp hết, hoặc không nạp gì nếu có lỗi (để bảo toàn dữ liệu).
                using (var trans = conn.BeginTransaction())
                {
                    try
                    {
                        // Thêm 2 cột metadata "FilePath" và "ImportDate" vào bảng dữ liệu máy đo để dễ truy vết.
                        if (!dt.Columns.Contains("FilePath")) dt.Columns.Add("FilePath", typeof(string));
                        if (!dt.Columns.Contains("ImportDate")) dt.Columns.Add("ImportDate", typeof(DateTime));
                        
                        // Gán giá trị cụ thể cho từng dòng dữ liệu.
                        foreach (DataRow row in dt.Rows) 
                        {
                            row["FilePath"] = filePath; 
                            row["ImportDate"] = DateTime.Now; 
                        }

                        // Sử dụng SqlBulkCopy: Đổ hàng ngàn dòng dữ liệu trực tiếp vào SQL trong vài giây.
                        using (var bulkCopy = new SqlBulkCopy(conn, SqlBulkCopyOptions.Default, trans))
                        {
                            bulkCopy.DestinationTableName = TableData; // Tên bảng đích.
                            // Thực hiện ánh xạ (Mapping) từng cột trong RAM sang cột tương ứng trên SQL Server.
                            for (int i = 0; i < SqlColumns.Length; i++) 
                            {
                                bulkCopy.ColumnMappings.Add(i, SqlColumns[i]); 
                            }
                            await bulkCopy.WriteToServerAsync(dt); // Thực hiện đổ mẻ dữ liệu.
                        }

                        // Sau khi nạp dữ liệu máy đo xong, ta ghi lại "Biên bản" nạp file vào bảng lịch sử.
                        long fileSize = new FileInfo(filePath).Length;
                        string historySql = $"INSERT INTO {TableHistory} (FilePath, FileSize, ImportedAt, RowsInserted, Status) " +
                                            "VALUES (@path, @size, GETDATE(), @rows, 'Success')";
                        
                        using (var cmd = new SqlCommand(historySql, conn, trans)) 
                        {
                            cmd.Parameters.AddWithValue("@path", filePath);   
                            cmd.Parameters.AddWithValue("@size", fileSize);   
                            cmd.Parameters.AddWithValue("@rows", dt.Rows.Count);   
                            await cmd.ExecuteNonQueryAsync(); 
                        }

                        trans.Commit(); // Hoàn tất giao dịch (Lưu vĩnh viễn vào DB).
                        _logger?.Invoke($"[DB-OK] Đã nạp thành công {dt.Rows.Count} dòng từ tệp {fileName}");
                        return dt.Rows.Count; 
                    }
                    catch (Exception ex)
                    {
                        trans.Rollback(); // Nếu có bất kỳ lỗi gì phát sinh, hủy bỏ toàn bộ mẻ nạp này để tránh rác DB.
                        _logger?.Invoke($"[DB-FAIL] Nạp dữ liệu thất bại cho tệp {fileName}: {ex.Message}"); 
                        throw; // Bắn lỗi lên lớp trên để Form1 xử lý in ra màn hình.
                    }
                }
            }
        }
    }
}
