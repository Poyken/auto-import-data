using Microsoft.Data.SqlClient;
using System;
using System.IO;
using System.Data;
using System.Threading.Tasks;
using ImportData.Core;

namespace ImportData.Services 
{
    /// <summary>
    /// Handles all database operations including connection testing, 
    /// duplicate detection, and high-performance bulk data imports.
    /// </summary>
    public class DatabaseService
    {
        private const string TableData = "SortingDataImportExcel_test";
        private const string TableHistory = "ExcelImportHistory_test";

        private static readonly string[] SqlColumns = {
            "EquipmentNumber", "SorterNum", "StartTime", "WorkflowCode", "LotNo",
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

        /// <summary>
        /// Tests if the SQL Server is accessible using the current configuration.
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
                
                var builder = new SqlConnectionStringBuilder(_config.ConnectionString) 
                { 
                    ConnectTimeout = _config.HealthCheckTimeoutSeconds 
                };
                
                using (var testConn = new SqlConnection(builder.ConnectionString)) 
                {
                    await testConn.OpenAsync(); 
                    return true; 
                } 
            }
            catch (Exception)
            {
                // Silently fail, Form1 handles health check status reporting
                return false;
            }
        }

        /// <summary>
        /// Checks if a file has already been successfully imported by looking up the history table.
        /// </summary>
        public async Task<bool> IsFileImportedAsync(string filePath) 
        {
            try
            {
                using (var conn = new SqlConnection(_config.ConnectionString)) 
                {
                    await conn.OpenAsync(); 
                    string sql = $"SELECT COUNT(*) FROM {TableHistory} WHERE FilePath = @path AND Status = 'Success'";
                    
                    using (var cmd = new SqlCommand(sql, conn)) 
                    {
                        cmd.Parameters.AddWithValue("@path", filePath); 
                        int count = (int)await cmd.ExecuteScalarAsync(); 
                        return count > 0; 
                    }
                } 
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[DB-ERROR] Failed to check import history: {ex.Message}");
                return false; 
            }
        }

        /// <summary>
        /// Imports all rows from a DataTable into the SQL Server using high-performance bulk copying.
        /// </summary>
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
                            for (int i = 0; i < SqlColumns.Length; i++) 
                            {
                                bulkCopy.ColumnMappings.Add(i, SqlColumns[i]); 
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
                        _logger?.Invoke($"[DB-OK] Successfully imported {dt.Rows.Count} rows from {fileName}");
                        return dt.Rows.Count; 
                    }
                    catch (Exception ex)
                    {
                        trans.Rollback(); 
                        _logger?.Invoke($"[DB-FAIL] Import failed for {fileName}: {ex.Message}"); 
                        throw; 
                    }
                }
            }
        }
    }
}
