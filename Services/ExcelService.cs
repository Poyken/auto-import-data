using ExcelDataReader;
using System;
using System.Data;
using System.IO;

namespace ImportData.Services
{
    /// <summary>
    /// Service chuyên xử lý logic liên quan đến file Excel
    /// </summary>
    public class ExcelService
    {
        private readonly Action<string> _logger;

        // Cấu trúc tiêu đề cột mong đợi từ máy đo (Đã thêm LotNo theo yêu cầu người dùng)
        private static readonly string[] RequiredHeaders = {
            "EquipmentNumber", "SorterNum", "StartTime", "WorkflowCode", "LotNo",
            "Barcode", "Slot", "Position", "Channel", "Capacity_mAh", "Capacitance_F", 
            "BeginVoltageSD_mV", "ChargeEndCurrent_mA", "EndVoltage_mV", "EndCurrent_mA", "DischargeVoltage1_mV", 
            "DischargeVoltage1_Time", "DischargeVoltage2_mV", "DischargeVoltage2_Time", "DischargeBeginVoltage_mV", "DischargeBeginCurrent_mA", 
            "NGInfo", "EndTime"
        };

        public ExcelService(Action<string> logger)
        {
            _logger = logger;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

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
                    // 1. KIỂM TRA CẤU TRÚC CỘT
                    if (!ValidateHeaders(dt))
                    {
                        _logger?.Invoke($"[LỖI CẤU TRÚC] File {Path.GetFileName(filePath)} không đúng định dạng cột yêu cầu.");
                        return null;
                    }

                    // 2. TIỀN XỬ LÝ (DATA CLEANSING)
                    foreach (DataRow row in dt.Rows)
                    {
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            var cellValue = row[i];
                            if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()) || cellValue.ToString() == "---")
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
                _logger?.Invoke($"Lỗi khi đọc file Excel: {ex.Message}");
                return null;
            }
        }

        private bool ValidateHeaders(DataTable dt)
        {
            if (dt == null) return false;
            // Kiểm tra số lượng cột tối thiểu (22 cột)
            return dt.Columns.Count >= RequiredHeaders.Length;
        }
    }
}
