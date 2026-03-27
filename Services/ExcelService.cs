using ExcelDataReader;
using System;
using System.Data;
using System.IO;

namespace ImportData.Services 
{
    /// <summary>
    /// Handles Excel file reading and basic data validation.
    /// Provides methods to convert raw Excel files into structured DataTables.
    /// </summary>
    public class ExcelService
    {
        private readonly Action<string> _logger;

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

        /// <summary>
        /// Reads an Excel file and returns a cleaned DataTable.
        /// Returns null if file is invalid or structural validation fails.
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
                        _logger?.Invoke($"[EXCEL-FAIL] Invalid header structure: {Path.GetFileName(filePath)}"); 
                        return null; 
                    }

                    // Basic data cleansing (e.g., replacing '---' with DBNull)
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
                _logger?.Invoke($"[EXCEL-ERROR] Failed to read Excel file: {ex.Message}"); 
                return null; 
            }
        }

        /// <summary>
        /// Validates that the DataTable contains at least the required number of columns.
        /// </summary>
        private bool ValidateHeaders(DataTable dt)
        {
            if (dt == null) return false; 
            return dt.Columns.Count >= RequiredHeaders.Length;
        }
    }
}
