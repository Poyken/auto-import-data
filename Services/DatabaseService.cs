using Microsoft.Data.SqlClient;
using System;
using System.Data;
using System.Threading.Tasks;
using ImportData.Core;

namespace ImportData.Services 
{
    public class DatabaseService
    {
        private readonly AppConfig _config;
        private readonly Action<string> _logger;

        // Hàm khởi tạo, bắt buộc phải truyền vào cấu hình và hàm ghi log khi gọi "new DatabaseService()".
        public DatabaseService(AppConfig config, Action<string> logger)
        {
            _config = config; 
            _logger = logger; 
        }

        /// <summary>
        /// Hàm kiểm tra kết nối xem SQL Server có đang hoạt động và đúng chỉ số mảng hay không.
        /// Chữ 'async' giúp hàm chạy ngầm mà không làm treo ứng dụng.
        /// </summary>
        public async Task<bool> TestConnectionAsync()
        {
            try 
            {
                SqlConnection.ClearAllPools(); // Xóa bộ đệm kết nối SQL cũ để kiểm tra trạng thái mạng thực tế nhất.
                
                // Đặt giới hạn thời gian chờ kết nối là 2 giây. Nếu mạng lag hơn 2 giây thì báo lỗi ngay, tránh app bị chờ quá lâu.
                var builder = new SqlConnectionStringBuilder(_config.ConnectionString) { ConnectTimeout = 2 };
                
                using (var testConn = new SqlConnection(builder.ConnectionString)) // Khởi tạo một ống kết nối thử tới SQL Server.
                {
                    await testConn.OpenAsync(); // Ra lệnh xin mở cửa SQL Server. 'await' bắt chương trình chờ cho đến khi mở xong mới chạy tiếp.
                    return true; // Nếu mở cửa thành công không bị báo lỗi, trả về giá trị Đúng (true).
                } 
            }
            catch (Exception)
            {
                _logger?.Invoke("[ERROR] SQL connection failed"); 
                return false;
            }
        }

        /// <summary>
        /// Hàm này kiểm tra xem một file Excel (ví dụ "Result_1.xlsx") đã từng được nạp vào SQL Server chưa.
        /// Giúp ứng dụng không bị nạp trùng lặp một file đã có dữ liệu.
        /// </summary>
        public async Task<bool> IsFileImportedAsync(string fileName) 
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(_config.ConnectionString)) // Mở kết nối.
                {
                    await conn.OpenAsync(); 
                    
                    // Câu lệnh SQL: Đếm xem trong bảng ImportHistory có bao nhiêu dòng chứa tên file này.
                    string sql = "SELECT COUNT(*) FROM ImportHistory WHERE FileName = @name";
                    
                    using (SqlCommand cmd = new SqlCommand(sql, conn)) // Đóng gói câu SQL vào nhóm lệnh để gửi đi.
                    {
                        cmd.Parameters.AddWithValue("@name", fileName); // Gán biến @name bằng tên file thực tế. Cách này giúp chặn rủi ro Hacker chèn mã độc (SQL Injection).
                        
                        // Yêu cầu SQL Server tính toán rồi trả về một con số đếm duy nhất. (ExecuteScalar).
                        int count = (int)await cmd.ExecuteScalarAsync(); 
                        
                        return count > 0; // Nếu số đếm > 0 nghĩa là file đã tồn tại trên DB, trả về true. Ngược lại trả về false.
                    }
                } 
            }
            catch (Exception)
            {
                _logger?.Invoke("[ERROR] History check failed"); 
                return false; 
            }
        }

        /// <summary>
        /// Hàm đẩy hàng ngàn dòng dữ liệu từ màn biến DataTable trên RAM vào bảng SQL Server bằng tính năng siêu tốc (SqlBulkCopy).
        /// Có hỗ trợ Transaction: Một là lưu thành công 100%, hai là sẽ tự hủy lệnh lưu giữa dòng nếu app bị treo màng.
        /// </summary>
        public async Task<int> ExecuteImportBatchAsync(DataTable dt, string fileName, string filePath) 
        {
            if (dt == null || dt.Rows.Count == 0) return 0; // Kiểm tra, nếu bảng trống trơn không có mãng dữ liệu dòng nào thì ngắt luôn, khỏi làm lại.

            using (SqlConnection conn = new SqlConnection(_config.ConnectionString)) 
            {
                await conn.OpenAsync(); 
                
                // Khởi tạo Transaction (Bảo mật giao dịch làm nhiều lệnh). 
                // Nếu bị sập điện nửa chừng, các dữ liệu nạp rở dang sẽ được báo xóa lùi bỏ để vẹn toàn dữ liệu bảng cấu trúc SQL cũ.
                using (SqlTransaction trans = conn.BeginTransaction())
                {
                    try
                    {
                        // File Excel gốc của máy đo không có cột lưu Đường dẫn (FilePath) và ngày giờ phần mềm Nạp (ImportDate).
                        // Ta cần tự tạo thêm 2 cột này vào DataTable trước khi đẩy lên SQL.
                        dt.Columns.Add("FilePath", typeof(string));
                        dt.Columns.Add("ImportDate", typeof(DateTime));
                        
                        foreach (DataRow row in dt.Rows) // L duyệt qua mỗi dữ liệu hàng...
                        {
                            row["FilePath"] = filePath; // Đóng dấu đường dẫn.
                            row["ImportDate"] = DateTime.Now; // Ấn mộc xác thực thời gian hệ thống.
                        }

                        // SqlBulkCopy: Công cụ chuyển dữ liệu nhanh nhất của C#. Thay vì nhồi từng dòng bằng biến INSERT rất nặng CPU, nó sẽ bốc cả khối lớn và ấn vào DB server.
                        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn, SqlBulkCopyOptions.Default, trans))
                        {
                            bulkCopy.DestinationTableName = "CapacitorLogs"; // Chỉ định ghi dữ liệu vào bảng SQL cố định mang tên "CapacitorLogs".
                            
                            // Tạo danh sách 25 cột được mapping từ Excel file vào đúng chỗ trên file SQL Table.
                            string[] sqlColumns = {
                                "EquipmentNumber", "SorterNum", "StartTime", "WorkflowCode", "LotNo",
                                "Barcode", "Slot", "Position", "Channel", "Capacity_mAh", "Capacitance_F", 
                                "BeginVoltageSD_mV", "ChargeEndCurrent_mA", "EndVoltage_mV", "EndCurrent_mA", "DischargeVoltage1_mV", 
                                "DischargeVoltage1_Time", "DischargeVoltage2_mV", "DischargeVoltage2_Time", "DischargeBeginVoltage_mV", "DischargeBeginCurrent_mA", 
                                "NGInfo", "EndTime", "FilePath", "ImportDate"
                            };

                            for (int i = 0; i < sqlColumns.Length; i++) 
                            {
                                // Kết nối cột: Cột lưới thứ 'i' của Data C# phải lưu chính xác vào cột chữ tên sqlColumns[i] trong bảng DB nhé.
                                bulkCopy.ColumnMappings.Add(i, sqlColumns[i]); 
                            }

                            await bulkCopy.WriteToServerAsync(dt); // Lệnh chốt: Xin phép chuyển toàn bộ màng DataTable vào DB.
                        }

                        // Tác vụ Phụ sau cùng: Sau khi nạp dữ liệu Máy đo thì ghi tên Bảng Máy đo (Ví dụ: Result20.xlsx) này luôn qua sổ bảng History (ImportHistory), báo app hoàn tất quá trình này.
                        string sql = "INSERT INTO ImportHistory (FileName, FilePath, ImportTime) VALUES (@name, @path, GETDATE())";
                        
                        using (SqlCommand cmd = new SqlCommand(sql, conn, trans)) // Chú ý cần phải nối trans (transaction) đi kèm khối giao dịch kia để chống lỗi sụp đổ giữa đường.
                        {
                            cmd.Parameters.AddWithValue("@name", fileName);   
                            cmd.Parameters.AddWithValue("@path", filePath);   
                            await cmd.ExecuteNonQueryAsync(); // Bấm cho lệnh Insert chạy đi vào SQL.
                        }

                        // Commit lệnh: Nếu mọi quy trình ở Code chạy mượt vào tới chóp này, có thể ra lệnh chốt SQL hoàn thành vòng chu trình bảo lưu an toàn kết quả vào ổ cứng Cực điểm!.
                        trans.Commit(); 
                        return dt.Rows.Count; // Vẩy tay Báo số liệu dòng đã được chuyển về Form cho biết dòng cuối thực thi.
                    }
                    catch (Exception)
                    {
                        trans.Rollback(); 
                        _logger?.Invoke($"[ERROR] Batch import failed: {fileName}"); 
                        throw; 
                    }
                }
            }
        }
    }
}
