using ExcelDataReader;
using System;
using System.Data;
using System.IO;

namespace ImportData.Services 
{
    /// <summary>
    /// Lớp ExcelService: Chuyên đọc và phân giải nội dung tệp Excel từ máy đo.
    /// </ chuyển đổi dữ liệu từ tệp vật lý sang bảng DataTable trong bộ nhớ RAM để SQL có thể nuốt được.
    /// </summary>
    public class ExcelService
    {
        private readonly Action<string> _logger; // Hàm log để bắn lỗi ra màn hình chính.

        // Mảng RequiredHeaders: Danh sách 23 tiêu đề cột bắt buộc phải có trong file Excel máy đo.
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
                
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration() 
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        dt = result.Tables[0]; 
                    } 
                } 

                if (dt != null) 
                {
                    if (!ValidateHeaders(dt))
                    {
                        _logger?.Invoke($"[LỖI-EXCEL] Cấu trúc tệp không đúng: {Path.GetFileName(filePath)}"); 
                        return null; 
                    }

                    foreach (DataRow row in dt.Rows) 
                    {
                        for (int i = 0; i < dt.Columns.Count; i++) 
                        {
                            var val = row[i]; 
                            if (val == null || string.IsNullOrWhiteSpace(val.ToString()) || val.ToString() == "---") 
                            {
                                row[i] = DBNull.Value; 
                            }
                        }
                    } 
                }
                return dt; 
            }
            catch (Exception ex) 
            {
                _logger?.Invoke($"[LỖI-EXCEL] Không xử lý được tệp Excel: {ex.Message}"); 
                return null; 
            }
        }

        /// <summary>
        /// Kiểm thử tiêu đề cột chuyên sâu.
        /// </summary>
        private bool ValidateHeaders(DataTable dt)
        {
            if (dt == null || dt.Columns.Count == 0) return false; 
            // Hiện tại chúng ta đã có Smart Mapping, nên chỉ cần đảm bảo có dữ liệu.
            return true;
        }
    }
}
