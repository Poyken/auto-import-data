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
        // _config: Nơi lưu trữ các thông số như Chuỗi kết nối SQL.
        private readonly AppConfig _config;
        
        // _logger: Hàm dùng để in các thông báo lỗi hoặc trạng thái ra màn hình ListBox.
        private readonly Action<string> _logger;

        // _lastConnectionString: Ghi nhớ chuỗi kết nối lần trước để phát hiện khi user đổi server trong appsettings.json.
        // Chỉ khi chuỗi này thay đổi ta mới xóa Connection Pool, tránh xóa mù mỗi 10 giây gây lãng phí tài nguyên mạng.
        private string _lastConnectionString;

        // Hàm khởi tạo: Được gọi khi viết "new DatabaseService(...)". 
        // Nó giúp "nhồi" (Inject) cấu hình và trình ghi log vào trong lớp này để sử dụng.
        public DatabaseService(AppConfig config, Action<string> logger)
        {
            _config = config; // Lưu cấu hình vào biến nội bộ.
            _logger = logger; // Lưu trình ghi log vào biến nội bộ.
            _lastConnectionString = config.ConnectionString; // Ghi nhớ chuỗi kết nối ban đầu.
        }

        /// <summary>
        /// Hàm TestConnectionAsync: Kiểm tra xem App có "nói chuyện" được với SQL Server không.
        /// Sử dụng Task<bool> để chạy ngầm, tránh việc App bị treo cứng (Not Responding) khi chờ mạng.
        /// </summary>
        public async Task<bool> TestConnectionAsync()
        {
            try 
            {
                // Chỉ xóa Connection Pool khi user thay đổi chuỗi kết nối trong appsettings.json.
                // Tránh xóa mù mỗi 10 giây làm gián đoạn các kết nối đang hoạt động (ví dụ BulkCopy đang nạp dữ liệu).
                if (_config.ConnectionString != _lastConnectionString)
                {
                    SqlConnection.ClearAllPools();
                    _lastConnectionString = _config.ConnectionString;
                }
                
                // Sử dụng SqlConnectionStringBuilder để tùy chỉnh lại chuỗi kết nối.
                // Timeout lấy từ cấu hình appsettings.json (mặc định 5 giây), đủ cho mạng LAN công nghiệp.
                var builder = new SqlConnectionStringBuilder(_config.ConnectionString) { ConnectTimeout = _config.HealthCheckTimeoutSeconds };
                
                // Khởi tạo một đối tượng kết nối (testConn) với chuỗi cấu hình mới.
                using (var testConn = new SqlConnection(builder.ConnectionString)) 
                {
                    // Thử mở cửa kết nối. Lệnh 'await' giúp CPU rảnh tay làm việc khác trong lúc chờ SQL hồi đáp.
                    await testConn.OpenAsync(); 
                    return true; // Nếu mở thành công (không có Exception), trả về True (Thành công).
                } 
            }
            catch (Exception)
            {
                // Không log ở đây vì health check chạy mỗi 10s, Form1 sẽ tự xử lý thông báo khi trạng thái thay đổi.
                return false;
            }
        }

        /// <summary>
        /// Hàm IsFileImportedAsync: Tra cứu xem cái tên file này đã có trong lịch sử nạp chưa.
        /// Chống việc nạp đi nạp lại cùng một dữ liệu máy đo gây sai số báo cáo.
        /// </summary>
        public async Task<bool> IsFileImportedAsync(string fileName) 
        {
            try
            {
                // Mở một ống dẫn kết nối tới Database.
                using (SqlConnection conn = new SqlConnection(_config.ConnectionString)) 
                {
                    await conn.OpenAsync(); // Mở kết nối.
                    
                // Câu lệnh SQL: Đếm (COUNT) xem có dòng nào trong bảng ExcelHistory_Import có đường dẫn file khớp không.
                // Chỉ tính các file đã nạp thành công (Status = 'Success').
                string sql = "SELECT COUNT(*) FROM ExcelHistory_Import WHERE FilePath = @path AND Status = 'Success'";
                
                using (SqlCommand cmd = new SqlCommand(sql, conn)) // Chuẩn bị gói lệnh SQL gởi đi.
                {
                    // Gán giá trị filePath vào tham số @path một cách an toàn.
                    cmd.Parameters.AddWithValue("@path", filePath); 
                        
                        // Thực thi lệnh và lấy về một con số duy nhất (ExecuteScalar).
                        int count = (int)await cmd.ExecuteScalarAsync(); 
                        
                        // Nếu count > 0 nghĩa là trong quá khứ file này đã nạp rồi.
                        return count > 0; 
                    }
                } 
            }
            catch (Exception)
            {
                // Nếu truy vấn lỗi (có thể do bảng chưa tạo), báo lỗi nhẹ và coi như chưa nạp để App thử lại sau.
                _logger?.Invoke("[LỖI] Không thể kiểm tra lịch sử nạp file."); 
                return false; 
            }
        }

        /// <summary>
        /// Hàm ExecuteImportBatchAsync: Cỗ máy chính để đẩy hàng ngàn dòng dữ liệu vào SQL tốc độ cao.
        /// Sử dụng kỹ thuật SqlBulkCopy giúp nạp 10,000 dòng trong chớp mắt.
        /// </summary>
        public async Task<int> ExecuteImportBatchAsync(DataTable dt, string fileName, string filePath) 
        {
            // Kiểm tra đầu vào: Nếu bảng rỗng hoặc null thì thoát luôn, không làm mất thời gian SQL.
            if (dt == null || dt.Rows.Count == 0) return 0; 

            using (SqlConnection conn = new SqlConnection(_config.ConnectionString)) 
            {
                await conn.OpenAsync(); // Kết nối tới máy chủ SQL.
                
                // Khởi tạo một "Giao dịch" (Transaction). 
                // Ý nghĩa: Đảm bảo tính toàn vẹn. Nếu nạp 100 dòng mà dòng 99 bị lỗi, nó sẽ TỰ HỦY 98 dòng trước đó.
                // Tránh tình trạng dữ liệu bị "nửa ông nửa bà" trong hệ thống.
                using (SqlTransaction trans = conn.BeginTransaction())
                {
                    try
                    {
                        // File Excel máy đo không có cột FilePath và ImportDate.
                        // Ta tự thêm 2 cột này vào bảng dữ liệu ảo (DataTable) trên RAM của C#.
                        dt.Columns.Add("FilePath", typeof(string));
                        dt.Columns.Add("ImportDate", typeof(DateTime));
                        
                        // Lặp qua từng dòng dữ liệu trong bảng để "đóng dấu" thông tin nguồn gốc.
                        foreach (DataRow row in dt.Rows) 
                        {
                            row["FilePath"] = filePath; // Lưu đường dẫn file gốc.
                            row["ImportDate"] = DateTime.Now; // Lưu ngày giờ nạp hiện tại của máy tính.
                        }

                        // SqlBulkCopy: Trình nạp dữ liệu hàng loạt mạnh nhất của C#.
                        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn, SqlBulkCopyOptions.Default, trans))
                        {
                            // Chỉ định tên bảng đích trong Database là "ExcelData_Import".
                            bulkCopy.DestinationTableName = "ExcelData_Import"; 
                            
                            // Định nghĩa danh sách các cột sẽ được map (khớp) từ file Excel sang bảng SQL.
                            // Thứ tự này phải khớp hoàn toàn với cấu trúc bảng trong SQL Server.
                            string[] sqlColumns = {
                                "EquipmentNumber", "SorterNum", "StartTime", "WorkflowCode", "LotNo",
                                "Barcode", "Slot", "Position", "Channel", "Capacity_mAh", "Capacitance_F", 
                                "BeginVoltageSD_mV", "ChargeEndCurrent_mA", "EndVoltage_mV", "EndCurrent_mA", "DischargeVoltage1_mV", 
                                "DischargeVoltage1_Time", "DischargeVoltage2_mV", "DischargeVoltage2_Time", "DischargeBeginVoltage_mV", "DischargeBeginCurrent_mA", 
                                "NGInfo", "EndTime", "FilePath", "ImportDate"
                            };

                            for (int i = 0; i < sqlColumns.Length; i++) 
                            {
                                // Ánh xạ: Cột thứ i của DataTable sẽ chui vào cột mang tên sqlColumns[i] trong DB.
                                bulkCopy.ColumnMappings.Add(i, sqlColumns[i]); 
                            }

                            // Chốt lệnh: Đẩy toàn bộ khối dữ liệu RAM lên Server SQL.
                            await bulkCopy.WriteToServerAsync(dt); 
                        }

                        // Sau khi nạp xong dữ liệu máy đo, ta ghi thêm 1 dòng vào bảng "ExcelHistory_Import".
                        // Để đánh dấu: "Tui đã nạp thành công file này nhé!".
                        long fileSize = new FileInfo(filePath).Length;
                        string sql = "INSERT INTO ExcelHistory_Import (FilePath, FileSize, ImportedAt, RowsInserted, Status) " +
                                     "VALUES (@path, @size, GETDATE(), @rows, 'Success')";
                        
                        using (SqlCommand cmd = new SqlCommand(sql, conn, trans)) 
                        {
                            cmd.Parameters.AddWithValue("@path", filePath);   
                            cmd.Parameters.AddWithValue("@size", fileSize);   
                            cmd.Parameters.AddWithValue("@rows", dt.Rows.Count);   
                            await cmd.ExecuteNonQueryAsync(); // Thực thi lệnh chèn lịch sử.
                        }

                        // Lệnh Commit: Nếu tới đây mọi việc đều mượt mà, ta ra lệnh cho SQL Server "Chốt sổ" lưu vĩnh viễn dữ liệu.
                        trans.Commit(); 
                        _logger?.Invoke($"[OK] Import success {dt.Rows.Count} rows from {fileName}");
                        return dt.Rows.Count; // Trả về số lượng dòng đã nạp thành công để hiện ra màn hình.
                    }
                    catch (Exception)
                    {
                        // Nếu có bất kỳ lỗi gì xảy ra (mất mạng, SQL đầy ổ cứng...), lệnh Rollback sẽ xóa sạch những gì vừa nạp dang dở.
                        trans.Rollback(); 
                        _logger?.Invoke($"[LỖI] Nạp dữ liệu thất bại {dt.Rows.Count} dòng từ {fileName}"); 
                        throw; // Quăng lỗi ra ngoài để hàm cha (Form1) xử lý tiếp.
                    }
                }
            }
        }
    }
}
