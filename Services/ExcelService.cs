using ExcelDataReader;
using System;
using System.Data;
using System.IO;

namespace ImportData.Services 
{
    /// <summary>
    /// Lớp ExcelService: Chuyên đọc và phân giải nội dung tệp Excel từ máy đo.
    /// Chuyển đổi dữ liệu từ tệp vật lý sang bảng DataTable trong bộ nhớ RAM để SQL có thể nuốt được.
    /// </summary>
    public class ExcelService
    {
        private readonly Action<string> _logger; // Hàm log để bắn lỗi ra màn hình chính.

        // Mảng RequiredHeaders: Danh sách 23 tiêu đề cột bắt buộc phải có trong file Excel máy đo.
        // Nếu thiếu bất kỳ cột nào, App sẽ coi tệp bị lỗi định dạng và bỏ qua.
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
            _logger = logger; 
            // RegisterProvider: Quan trọng để Windows biết cách đọc bảng mã cũ của tệp Excel máy đo.
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        /// <summary>
        /// Đọc nội dung tệp Excel vật lý và chuyển thành DataTable.
        /// Sử dụng FileShare.ReadWrite để không làm gián đoạn máy đo.
        /// </summary>
        public DataTable ReadExcelFile(string filePath)
        {
            try 
            {
                DataTable dt; 
                
                // Mở File theo chế độ 'Lịch sử' (ReadWrite): Cho phép App đọc file ngay cả khi máy đo đang bắt đầu ghi file mới.
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // Triệu hồi thư viện ExcelDataReader để giải mã các mảng nhị phân.
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration() 
                        {
                            // Thiết lập quan trọng: Hãy coi dòng đầu tiên của Excel là Tên Cột (Header).
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        dt = result.Tables[0]; // Lấy dữ liệu từ Sheet đầu tiên của file Excel.
                    } 
                } 

                if (dt != null) 
                {
                    // BƯỚC 1: Kiểm thử cấu trúc. File đo xịn phải đủ 23 cột tiêu chuẩn.
                    if (!ValidateHeaders(dt))
                    {
                        _logger?.Invoke($"[LỖI-EXCEL] Cấu trúc tệp không đúng tiêu chuẩn: {Path.GetFileName(filePath)}"); 
                        return null; 
                    }

                    // BƯỚC 2: Dọn dẹp dữ liệu rác (Data Cleansing).
                    foreach (DataRow row in dt.Rows) 
                    {
                        for (int i = 0; i < dt.Columns.Count; i++) 
                        {
                            var val = row[i]; 
                            // Nếu ô đó bị trống hoặc chứa dấu gạch "---" (Lỗi đo), thay bằng giá trị Rỗng chuẩn của SQL (DBNull).
                            if (val == null || string.IsNullOrWhiteSpace(val.ToString()) || val.ToString() == "---") 
                            {
                                row[i] = DBNull.Value; 
                            }
                        }
                    } 
                }
                return dt; // Trả về bảng dữ liệu sạch sẽ, sẵn sàng nạp.
            }
            catch (Exception ex) 
            {
                // Nếu file bị hỏng nặng hoặc đang bị khóa cứng không cho đọc thì báo lỗi.
                _logger?.Invoke($"[LỖI-EXCEL] Không xử lý được tệp Excel: {ex.Message}"); 
                return null; 
            }
        }

        /// <summary>
        /// Kiểm thử tiêu đề cột: Đảm bảo số lượng cột trong Excel >= 23 cột bắt buộc.
        /// </summary>
        private bool ValidateHeaders(DataTable dt)
        {
            if (dt == null) return false; 
            return dt.Columns.Count >= RequiredHeaders.Length;
        }
    }
}
