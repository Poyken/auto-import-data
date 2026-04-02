using Microsoft.Data.SqlClient;
using System;
using System.IO;
using System.Data;
using System.Threading.Tasks;
using System.Collections.Generic;
using ImportData.Core;

namespace ImportData.Services 
{
    public class DatabaseService
    {
        private const string TableData = "SortingDataImportExcel";
        private const string TableHistory = "ExcelImportHistory";

        private static readonly string[] SqlColumns = {
            "EquipmentNumber", "SorterNum", "StartTime", "WorkflowCode",
            "Barcode", "Slot", "Position", "Channel", "Capacity_mAh", "Capacitance_F", 
            "BeginVoltageSD_mV", "ChargeEndCurrent_mA", "EndVoltage_mV", "EndCurrent_mA", "DischargeVoltage1_mV", 
            "DischargeVoltage1_Time", "DischargeVoltage2_mV", "DischargeVoltage2_Time", "DischargeBeginVoltage_mV", "DischargeBeginCurrent_mA", 
            "NGInfo", "EndTime", "FilePath", "ImportDate"
        };

        private readonly AppConfig _config;
        private readonly Action<string> _logger;
        private string _lastConnectionString;

        public DatabaseService(AppConfig config, Action<string> logger)
        {
            _config = config; 
            _logger = logger; 
            _lastConnectionString = config.ConnectionString; 
        }

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
            catch (Exception ex)
            {
                _logger?.Invoke($"[LỖI-SQL-CONNECT] {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Kiểm tra trùng lặp thông minh: Dựa trên Tên tệp và Dung lượng.
        /// Trả về: 
        /// 0 = Chưa nạp bao giờ.
        /// 1 = Đã nạp rồi và Path giống hệt (Không làm gì).
        /// 2 = Đã nạp rồi nhưng đang ở Path mới (Cần cập nhật Path trong lịch sử).
        /// </summary>
        public async Task<int> CheckImportStatusAsync(string filePath)
        {
            try
            {
                string fileName = Path.GetFileName(filePath);
                long fileSize = new FileInfo(filePath).Length;

                using (var conn = new SqlConnection(_config.ConnectionString)) 
                {
                    await conn.OpenAsync();
                    string sql = $"SELECT FilePath FROM {TableHistory} WHERE (FilePath LIKE '%' + @name) AND FileSize = @size AND Status = 'Success'";
                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@name", fileName);
                        cmd.Parameters.AddWithValue("@size", fileSize);
                        
                        using (var reader = await cmd.ExecuteReaderAsync())
                        {
                            if (await reader.ReadAsync())
                            {
                                string existingPath = reader["FilePath"].ToString();
                                // Nếu Path giống hệt -> Đã nạp hoàn toàn.
                                if (string.Equals(existingPath, filePath, StringComparison.OrdinalIgnoreCase)) return 1;
                                // Nếu Path khác -> Đã nạp nhưng dời chỗ -> Cần cập nhật Path.
                                return 2;
                            }
                        }
                    }
                }
                return 0; // Chưa nạp.
            }
            catch { return 0; }
        }

        public async Task UpdateHistoryPathAsync(string filePath)
        {
            try
            {
                string fileName = Path.GetFileName(filePath);
                long fileSize = new FileInfo(filePath).Length;
                using (var conn = new SqlConnection(_config.ConnectionString))
                {
                    await conn.OpenAsync();
                    string sql = $"UPDATE {TableHistory} SET FilePath = @newPath, ImportedAt = GETDATE() WHERE (FilePath LIKE '%' + @name) AND FileSize = @size";
                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@newPath", filePath);
                        cmd.Parameters.AddWithValue("@name", fileName);
                        cmd.Parameters.AddWithValue("@size", fileSize);
                        await cmd.ExecuteNonQueryAsync();
                    }
                }
                _logger?.Invoke($"[HISTORY-UPDATE] Đã cập nhật vị trí mới cho tệp: {fileName}");
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[LỖI-UPDATE-PATH] {ex.Message}");
            }
        }

        public async Task<int> ExecuteImportBatchAsync(DataTable dt, string fileName, string filePath) 
        {
            if (dt == null || dt.Rows.Count == 0) return 0;

            using (var conn = new SqlConnection(_config.ConnectionString)) 
            {
                await conn.OpenAsync(); 
                using (var trans = conn.BeginTransaction())
                {
                    try
                    {
                        if (!dt.Columns.Contains("FilePath")) dt.Columns.Add("FilePath", typeof(string));
                        if (!dt.Columns.Contains("ImportDate")) dt.Columns.Add("ImportDate", typeof(DateTime));
                        
                        foreach (DataRow row in dt.Rows) 
                        {
                            row["FilePath"] = filePath; 
                            row["ImportDate"] = DateTime.Now; 
                        }

                        using (var bulkCopy = new SqlBulkCopy(conn, SqlBulkCopyOptions.Default, trans))
                        {
                            bulkCopy.DestinationTableName = TableData;
                            bulkCopy.BatchSize = 1000;
                            bulkCopy.BulkCopyTimeout = 120;

                            // Bảng ánh xạ Thủ công chính xác (Ưu tiên số 1)
                            var manualMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
                                { "Equipment Number", "EquipmentNumber" },
                                { "EquipmentNumber", "EquipmentNumber" },
                                { "LotNo", "Barcode" },
                                { "Lot No", "Barcode" },
                                { "Capacity(mAh)", "Capacity_mAh" },
                                { "Capacitance(F)", "Capacitance_F" },
                                { "BeginVoltageSD(mV)", "BeginVoltageSD_mV" },
                                { "Charge EndCurrent(mA)", "ChargeEndCurrent_mA" },
                                { "EndVoltage(mV)", "EndVoltage_mV" },
                                { "EndCurrent(mA)", "EndCurrent_mA" },
                                { "DischargeVoltage1(mV)", "DischargeVoltage1_mV" },
                                { "DischargeVoltage1_Time", "DischargeVoltage1_Time" },
                                { "DischargeVoltage2(mV)", "DischargeVoltage2_mV" },
                                { "DischargeVoltage2_Time", "DischargeVoltage2_Time" },
                                { "DischargeBeginVoltage(mV)", "DischargeBeginVoltage_mV" },
                                { "DischargeBeginCurrent(mA)", "DischargeBeginCurrent_mA" }
                            };

                            foreach (DataColumn dc in dt.Columns)
                            {
                                string colName = dc.ColumnName.Trim();
                                // Xử lý trường hợp xuống dòng trong Header Excel
                                string cleanName = colName.Replace("\r", " ").Replace("\n", " ").Replace("  ", " ").Trim();

                                if (manualMap.ContainsKey(cleanName))
                                {
                                    bulkCopy.ColumnMappings.Add(dc.ColumnName, manualMap[cleanName]);
                                }
                                else
                                {
                                    // Fuzzy Match cho các cột còn lại
                                    string fuzzySearch = cleanName.Replace(" ", "").Replace("_", "").ToLower();
                                    foreach (string sqlCol in SqlColumns)
                                    {
                                        if (sqlCol.Replace("_", "").ToLower().Contains(fuzzySearch) || fuzzySearch.Contains(sqlCol.Replace("_", "").ToLower()))
                                        {
                                            bulkCopy.ColumnMappings.Add(dc.ColumnName, sqlCol);
                                            break;
                                        }
                                    }
                                }
                            }
                            
                            await bulkCopy.WriteToServerAsync(dt);
                        }

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

                        trans.Commit();
                        _logger?.Invoke($"[DB-OK] Đã nạp thành công {dt.Rows.Count} dòng từ tệp {fileName}");
                        return dt.Rows.Count; 
                    }
                    catch (Exception ex)
                    {
                        trans.Rollback();
                        _logger?.Invoke($"[DB-FAIL] Nạp dữ liệu thất bại cho tệp {fileName}: {ex.Message}"); 
                        throw;
                    }
                }
            }
        }
    }
}
