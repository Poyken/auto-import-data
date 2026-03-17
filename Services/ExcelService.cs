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
        private readonly Action<string> _logger;

        // Mảng Tên 23 cột bắt buộc phải tồn tại trong file excel máy xuất (Lưu ý: Mới bổ sung thêm cột "LotNo").
        private static readonly string[] RequiredHeaders = {
            "EquipmentNumber", "SorterNum", "StartTime", "WorkflowCode", "LotNo",
            "Barcode", "Slot", "Position", "Channel", "Capacity_mAh", "Capacitance_F", 
            "BeginVoltageSD_mV", "ChargeEndCurrent_mA", "EndVoltage_mV", "EndCurrent_mA", "DischargeVoltage1_mV", 
            "DischargeVoltage1_Time", "DischargeVoltage2_mV", "DischargeVoltage2_Time", "DischargeBeginVoltage_mV", "DischargeBeginCurrent_mA", 
            "NGInfo", "EndTime"
        };

        // Hàm khởi tạo, khi hệ thống bật lên, ta cấu hình lại Encoding.
        public ExcelService(Action<string> logger)
        {
            _logger = logger; 
            
            // Một số bản .NET hiện hành đã chặn hỗ trợ bảng mã ASCII hoặc ISO quá cũ của những tệp .xls ngày xưa. 
            // RegisterProvider khắc phục lệnh khuyết này cho hệ thống app C# chúng ta quét được luôn những tệp xuất từ cỗ máy đo dòng đời rớt lâu.
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        /// <summary>
        /// Hàm tiếp nhận một đường truyền File (filePath) để xử lý giải mã nội dung file Excel thành dạng DataTable lưới trên RAM.
        /// </summary>
        public DataTable ReadExcelFile(string filePath)
        {
            try 
            {
                DataTable dt; // Biến nhận bảng dữ liệu lấy được.
                
                // Mở luồng đọc file bằng quyền chia sẻ (FileShare.ReadWrite): 
                // Quan trọng: Máy đo liên tục ghi dữ liệu vài lần nên file Excel thường bị hệ điều hành 'Khóa lại'. 
                // Bằng việc gởi cờ chia sẻ 'ReadWrite' ở đây, Windows vẫn cấp cho tệp chúng ta thẩm quyền được xem, được đọc dù máy đo vẫn đang mở sửa bên dưới nền.  
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // Chuyển dòng dữ liệu file cho công cụ ExcelDataReader đọc (nó rất thông minh sẽ hiểu cả .xls hay .xlsx).
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration() 
                        {
                            // Đánh dấu thiết lập bắt thư viện chuyển dòng 1 của File làm Tên các Cột Thay vì nhận dạng là một dòng Data chứa biến đếm.
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        dt = result.Tables[0]; // Chỉ thu lại duy nhất Sheet đầu tiên (Vị trí 0) làm khung làm việc chính thức cho mã. Lướt qua sheet thừa (vd sheet2, sheet3 nếu có).
                    } 
                } 

                if (dt != null) 
                {
                    // 1. Kiểm duyệt thông số cột (Validate Headers) trước khi đổ số vào cơ sở dữ liệu.
                    // Loại trừ việc bản nâng cấp máy đo vô tình xuất rớt 1 cột đi (Tạo lỗi cột lệch vào SQL).  
                    if (!ValidateHeaders(dt))
                    {
                        _logger?.Invoke($"[ERROR] Invalid layout: {Path.GetFileName(filePath)}"); 
                        return null; 
                    }

                    // 2. Tinh Xử Lý Dữ Liệu Chống Lỗi (Data Cleansing) - Bẩn giá trị
                    foreach (DataRow row in dt.Rows) // Truy vấn trên mỗi dòng dữ liệu...
                    {
                        for (int i = 0; i < dt.Columns.Count; i++) // Và di qua các cột đo bên trong hàng...
                        {
                            var cellValue = row[i]; // Lấy giá trị từng điểm đo lưu vào cellValue.
                            
                            // Một số máy đo xả một giá trị khuyết đặc biệt báo lỗi ví dụ chuỗi có ký hiệu "---" (Trống tính điện) hoặc bị mất giá trị nhị phân null.
                            if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()) || cellValue.ToString() == "---") 
                            {
                                // Sẽ bị gán cho biến giá trị 'Khuyết Hệ Thống' DBNull.Value.
                                // SQL rất cần một Null gốc làm chuẩn này thay vì phải đọc được cụm chữ "---" trong cột Number gây phát sập lỗi.
                                row[i] = DBNull.Value; 
                            }
                        }
                    } 
                }

                return dt; // Trả về dạng khung dữ liệu lưới DataTable hoàn thiện không còn chứa các biến khuyết hư cho phép để gới lên SQL Nạp Data..
            }
            catch (Exception ex) 
            {
                _logger?.Invoke($"[ERROR] Excel read: {ex.Message}"); 
                return null; 
            }
        }

        // Kiểm định và bảo mật tính toàn vẹn Số Cột của Bảng Lưới Excel Đo lấy ra:
        private bool ValidateHeaders(DataTable dt)
        {
            if (dt == null) return false; 
            
            // Logic >= có nghĩa Cột sinh ra nếu chẳng may có chứa 1 cột nhảm ở phía ngoài rìa màn > 23 cột. 
            // C# không xem là Lỗi vì nó lúc sau sẽ cắt và map vừa khít lấy 23 cột này đẩy lên SQL.  
            // Chỉ khi Cột máy Xuất đưa ra dưới 23 (Tức là đang đánh mất 1 cột thiết yếu) code mới thả False chặn!
            return dt.Columns.Count >= RequiredHeaders.Length;
        }
    }
}
