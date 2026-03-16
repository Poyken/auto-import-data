using Microsoft.Data.SqlClient;
using System;
using System.Data;
using System.Threading.Tasks;
using ImportData.Core;

namespace ImportData.Services
{
    /// <summary>
    /// Service chuyên xử lý nghiệp vụ thao tác với lớp dữ liệu SQL Server
    /// </summary>
    public class DatabaseService
    {
        private readonly AppConfig _config;
        private readonly Action<string> _logger;

        public DatabaseService(AppConfig config, Action<string> logger)
        {
            _config = config;
            _logger = logger;
        }

        public async Task<bool> TestConnectionAsync()
        {
            try
            {
                // Ngắt Pool để kiểm tra nóng trong trường hợp mạng vừa có lại
                SqlConnection.ClearAllPools();
                
                var builder = new SqlConnectionStringBuilder(_config.ConnectionString) { ConnectTimeout = 2 };
                using (var testConn = new SqlConnection(builder.ConnectionString))
                {
                    await testConn.OpenAsync();
                    return true;
                }
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[CHI TIẾT LỖI DB] {ex.Message}");
                return false;
            }
        }

        public async Task<bool> IsFileImportedAsync(string fileName)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(_config.ConnectionString))
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
            catch (Exception ex)
            { 
                _logger?.Invoke($"Lỗi kết nối DB (Kiểm tra file): {ex.Message}");
                return false; 
            } 
        }

        public async Task MarkFileAsImportedAsync(string fileName, string filePath)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(_config.ConnectionString))
                {
                    await conn.OpenAsync();
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
                _logger?.Invoke($"Lỗi lưu lịch sử (Database): {ex.Message}");
            }
        }

        public async Task<int> ExecuteImportBatchAsync(DataTable dt)
        {
            if (dt == null || dt.Rows.Count == 0) return 0;

            using (SqlConnection conn = new SqlConnection(_config.ConnectionString))
            {
                await conn.OpenAsync();
                using (SqlTransaction trans = conn.BeginTransaction())
                {
                    try
                    {
                        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn, SqlBulkCopyOptions.Default, trans))
                        {
                            bulkCopy.DestinationTableName = "CapacitorLogs";
                            
                            string[] sqlColumns = {
                                "EquipmentNumber", "SorterNum", "StartTime", "WorkflowCode", "LotNo", 
                                "Barcode", "Slot", "Position", "Channel", "Capacity_mAh", 
                                "Capacitance_F", "BeginVoltageSD_mV", "ChargeEndCurrent_mA", "EndVoltage_mV", "EndCurrent_mA", 
                                "DischargeVoltage1_mV", "DischargeVal1_Time", "DischargeVoltage2_mV", "DischargeVal2_Time", "DischargeBeginVoltage_mV", 
                                "DischargeBeginCurrent_mA", "NGInfo", "EndTime"
                            };

                            for (int i = 0; i < 23; i++)
                            {
                                bulkCopy.ColumnMappings.Add(i, sqlColumns[i]);
                            }

                            await bulkCopy.WriteToServerAsync(dt);
                        }

                        trans.Commit();
                        return dt.Rows.Count;
                    }
                    catch (Exception ex)
                    {
                        trans.Rollback();
                        _logger?.Invoke($"Lỗi BulkCopy (Ghi dữ liệu): {ex.Message}");
                        throw;
                    }
                }
            }
        }
    }
}
