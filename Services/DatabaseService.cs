using Microsoft.Data.SqlClient;
using System.IO;
using System;
using System.Data;
using System.Threading.Tasks;
using ImportData.Core;

namespace ImportData.Services 
{
    public class DatabaseService
    {
        // --- CÀI ĐẶT CỐ ĐỊNH (Constants) ---
        // Tên các bảng trong SQL Server. Gom về đây để sau này đổi tên bảng chỉ cần sửa 1 chỗ.
        private const string TableData = "SortingDataImportExcel_test";
        private const string TableHistory = "ExcelImportHistory_test";

        // Danh sách các cột SQL tương ứng với file Excel máy đo.
        // Việc để static giúp App không phải tạo lại mảng này mỗi khi nạp file, tiết kiệm RAM.
        private static readonly string[] SqlColumns = {
            "EquipmentNumber", "SorterNum", "StartTime", "WorkflowCode", "LotNo",
            "Barcode", "Slot", "Position", "Channel", "Capacity_mAh", "Capacitance_F", 
            "BeginVoltageSD_mV", "ChargeEndCurrent_mA", "EndVoltage_mV", "EndCurrent_mA", "DischargeVoltage1_mV", 
            "DischargeVoltage1_Time", "DischargeVoltage2_mV", "DischargeVoltage2_Time", "DischargeBeginVoltage_mV", "DischargeBeginCurrent_mA", 
            "NGInfo", "EndTime", "FilePath", "ImportDate"
        };

        // _config: Nơi lưu trữ các thông số như Chuỗi kết nối SQL.
        private readonly AppConfig _config;
        
        // _logger: Hàm dùng để in các thông báo lỗi hoặc trạng thái ra màn hình ListBox.
        private readonly Action<string> _logger;

        // _lastConnectionString: Ghi nhớ chuỗi kết nối lần trước để phát hiện khi user đổi server trong appsettings.json.
        private string _lastConnectionString;

        // Hàm khởi tạo: Được gọi khi viết "new DatabaseService(...)". 
        public DatabaseService(AppConfig config, Action<string> logger)
        {
            _config = config; 
            _logger = logger; 
            _lastConnectionString = config.ConnectionString; 
        }

        /// <summary>
        /// Hàm TestConnectionAsync: Kiểm tra xem App có "nói chuyện" được với SQL Server không.
        /// </summary>
        public async Task<bool> TestConnectionAsync()
        {
            try 
            {
                if (_config.ConnectionString != _lastConnectionString)
                {
                    SqlConnection.ClearAllPools();
                    _lastConnectionString = _config.ConnectionString;
                }
                
                var builder = new SqlConnectionStringBuilder(_config.ConnectionString) { ConnectTimeout = _config.HealthCheckTimeoutSeconds };
                using (var testConn = new SqlConnection(builder.ConnectionString)) 
                {
                    await testConn.OpenAsync(); 
                    return true; 
                } 
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Hàm IsFileImportedAsync: Tra cứu xem cái tên file này đã có trong lịch sử nạp chưa.
        /// </summary>
        public async Task<bool> IsFileImportedAsync(string filePath) 
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(_config.ConnectionString)) 
                {
                    await conn.OpenAsync(); 
                    
                    // Sử dụng hằng số TableHistory thay vì viết cứng tên bảng.
                    string sql = $"SELECT COUNT(*) FROM {TableHistory} WHERE FilePath = @path AND Status = 'Success'";
                    
                    using (SqlCommand cmd = new SqlCommand(sql, conn)) 
                    {
                        cmd.Parameters.AddWithValue("@path", filePath); 
                        int count = (int)await cmd.ExecuteScalarAsync(); 
                        return count > 0; 
                    }
                } 
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[LỖI] Không thể kiểm tra lịch sử nạp file: {ex.Message}"); 
                return false; 
            }
        }

        /// <summary>
        /// Hàm ExecuteImportBatchAsync: Cỗ máy chính để đẩy hàng ngàn dòng dữ liệu vào SQL tốc độ cao.
        /// </summary>
        public async Task<int> ExecuteImportBatchAsync(DataTable dt, string fileName, string filePath) 
        {
            if (dt == null || dt.Rows.Count == 0) return 0; 

            using (SqlConnection conn = new SqlConnection(_config.ConnectionString)) 
            {
                await conn.OpenAsync(); 
                
                using (SqlTransaction trans = conn.BeginTransaction())
                {
                    try
                    {
                        // Thêm thông tin nguồn gốc file và ngày nạp.
                        if (!dt.Columns.Contains("FilePath")) dt.Columns.Add("FilePath", typeof(string));
                        if (!dt.Columns.Contains("ImportDate")) dt.Columns.Add("ImportDate", typeof(DateTime));
                        
                        foreach (DataRow row in dt.Rows) 
                        {
                            row["FilePath"] = filePath; 
                            row["ImportDate"] = DateTime.Now; 
                        }

                        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn, SqlBulkCopyOptions.Default, trans))
                        {
                            bulkCopy.DestinationTableName = TableData; // Sử dụng hằng số TableData.
                            
                            // Ánh xạ các cột từ DataTable sang SQL.
                            for (int i = 0; i < SqlColumns.Length; i++) 
                            {
                                bulkCopy.ColumnMappings.Add(i, SqlColumns[i]); 
                            }

                            await bulkCopy.WriteToServerAsync(dt); 
                        }

                        // Ghi lại lịch sử nạp file thành công.
                        long fileSize = new FileInfo(filePath).Length;
                        string historySql = $"INSERT INTO {TableHistory} (FilePath, FileSize, ImportedAt, RowsInserted, Status) " +
                                            "VALUES (@path, @size, GETDATE(), @rows, 'Success')";
                        
                        using (SqlCommand cmd = new SqlCommand(historySql, conn, trans)) 
                        {
                            cmd.Parameters.AddWithValue("@path", filePath);   
                            cmd.Parameters.AddWithValue("@size", fileSize);   
                            cmd.Parameters.AddWithValue("@rows", dt.Rows.Count);   
                            await cmd.ExecuteNonQueryAsync(); 
                        }

                        trans.Commit(); 
                        _logger?.Invoke($"[OK] Import success {dt.Rows.Count} rows from {fileName}");
                        return dt.Rows.Count; 
                    }
                    catch (Exception ex)
                    {
                        trans.Rollback(); 
                        _logger?.Invoke($"[LỖI] Nạp dữ liệu thất bại {dt.Rows.Count} dòng từ {fileName}: {ex.Message}"); 
                        throw; 
                    }
                }
            }
        }
    }
}
