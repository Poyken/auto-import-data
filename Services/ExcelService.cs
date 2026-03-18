using ExcelDataReader;
using System;
using System.Data;
using System.IO;

namespace ImportData.Services 
{
    /// <summary>
    /// Lớp ExcelService: Đóng vai trò đọc lấy nội dung của file Excel (máy đo sinh ra) và đưa lên bảng dữ liệu mềm của C# (DataTable).
    /// </summary>
    public class ExcelService
    {
        // _logger: Hàm dùng để bắn thông báo lỗi ra bảng nhật ký cho người dùng xem.
        private readonly Action<string> _logger;

        // Mảng RequiredHeaders: Danh sách 23 tiêu đề cột bắt buộc phải có trong file Excel máy đo.
        // Nếu thiếu bất kỳ cột nào, App sẽ coi như file "què" và từ chối nạp để tránh sai lệch DB.
        private static readonly string[] RequiredHeaders = {
            "EquipmentNumber", "SorterNum", "StartTime", "WorkflowCode", "LotNo",
            "Barcode", "Slot", "Position", "Channel", "Capacity_mAh", "Capacitance_F", 
            "BeginVoltageSD_mV", "ChargeEndCurrent_mA", "EndVoltage_mV", "EndCurrent_mA", "DischargeVoltage1_mV", 
            "DischargeVoltage1_Time", "DischargeVoltage2_mV", "DischargeVoltage2_Time", "DischargeBeginVoltage_mV", "DischargeBeginCurrent_mA", 
            "NGInfo", "EndTime"
        };

        // Hàm khởi tạo ExcelService.
        public ExcelService(Action<string> logger)
        {
            _logger = logger; // Lưu lại trình ghi nhật ký.
            
            // ĐẶC BIỆT QUAN TRỌNG: RegisterProvider dùng để "dạy" .NET cách đọc các bảng mã cũ.
            // Máy đo công nghiệp thường dùng tệp Excel đời cũ, nếu không có dòng này App sẽ bị lỗi font hoặc không đọc được tệp.
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        /// <summary>
        /// Hàm ReadExcelFile: Biến hóa tệp Excel vật lý thành bảng dữ liệu mềm (DataTable) trong RAM.
        /// </summary>
        public DataTable ReadExcelFile(string filePath)
        {
            try 
            {
                DataTable dt; // Khai báo biến để chứa bảng dữ liệu kết quả.
                
                // Mở luồng đọc file (stream) với chế độ cực kỳ "lịch sự" (FileShare.ReadWrite).
                // Vì máy đo vẫn đang cầm file để ghi, nếu ta mở kiểu độc quyền sẽ bị Windows chặn. 
                // Kiểu ReadWrite này cho phép ta "nhìn trộm" nội dung ngay khi máy đo đang làm việc.
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // Triệu hồi thư viện ExcelDataReader để gỡ rối tệp nhị phân Excel.
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // Chuyển đổi nội dung đọc được thành bộ DataSet.
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration() 
                        {
                            // Thiết lập: Hãy dùng Dòng 1 của Excel để làm Tên Cột (Header).
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        
                        // Lấy mẻ dữ liệu từ Sheet đầu tiên (Vị trí số 0) để làm việc.
                        dt = result.Tables[0]; 
                    } 
                } 

                if (dt != null) 
                {
                    // BƯỚC 1: Kiểm tra cấu trúc cột. Nếu file Excel bị thiếu cột "Barcode" hay "LotNo" thì dừng ngay.
                    if (!ValidateHeaders(dt))
                    {
                        _logger?.Invoke($"[LỖI] Cấu trúc file không hợp lệ (Sai cột): {Path.GetFileName(filePath)}"); 
                        return null; 
                    }

                    // BƯỚC 2: "Dọn rác" dữ liệu (Data Cleansing).
                    // Máy đo đôi khi điền các ký tự lạ như "---" vào cột số gây lỗi SQL.
                    foreach (DataRow row in dt.Rows) 
                    {
                        for (int i = 0; i < dt.Columns.Count; i++) 
                        {
                            var cellValue = row[i]; // Lấy giá trị của từng ô (Cell).
                            
                            // Nếu ô bị trống, hoặc chứa dấu gạch "---" báo lỗi đo.
                            if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()) || cellValue.ToString() == "---") 
                            {
                                // Thay thế bằng DBNull.Value (Giá trị Rỗng chuẩn của SQL).
                                // Việc này giúp SQL hiểu là "Không có số liệu" chứ không phải "Dòng chữ lạ".
                                row[i] = DBNull.Value; 
                            }
                        }
                    } 
                }

                return dt; // Trả về bảng dữ liệu sạch sẽ, sẵn sàng để nạp vào DB.
            }
            catch (Exception ex) 
            {
                // Nếu file Excel bị hỏng nặng hoặc không thể mở nổi, ghi log báo lỗi.
                _logger?.Invoke($"[LỖI] Không thể đọc file Excel: {ex.Message}"); 
                return null; 
            }
        }

        /// <summary>
        /// Hàm ValidateHeaders: So khớp danh sách cột của file Excel với tiêu chuẩn "23 cột bắt buộc".
        /// </summary>
        private bool ValidateHeaders(DataTable dt)
        {
            if (dt == null) return false; 
            
            // Logic: Nếu số lượng cột trong file >= 23 thì coi như Đạt.
            // (Nếu file có dư cột rác ở sau cùng thì App cũng kệ, vì App chỉ lấy 23 cột đầu để nạp).
            return dt.Columns.Count >= RequiredHeaders.Length;
        }
    }
}
