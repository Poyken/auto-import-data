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

        public ExcelService(Action<string> logger)
        {
            _logger = logger;
            // Đăng ký bộ Encoding để đọc được file Excel chứa mã ký tự đặc biệt/tiếng Việt
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
                        dt = result.Tables[0]; // Chỉ lấy Sheet đầu tiên của máy đo sinh ra
                    }
                }

                // TIỀN XỬ LÝ (DATA CLEANSING): Giúp SqlBulkCopy không bị crash vì sai Type
                if (dt != null)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            var cellValue = row[i];
                            
                            // Nếu ô trống (null hoặc thẻ string rỗng "") -> Ép thành DBNull để Database chèn giá trị NULL thay vì báo lỗi
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
    }
}
