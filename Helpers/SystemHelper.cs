using Microsoft.Win32;
using System;
using System.Windows.Forms;

namespace ImportData.Helpers
{
    /// <summary>
    /// Các Helper tương tác với System
    /// </summary>
    public static class SystemHelper
    {
        public static void SetStartup(Action<string> logger)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    key.SetValue("AutoImportData", Application.ExecutablePath);
                }
            }
            catch (Exception ex)
            {
                logger?.Invoke($"Lỗi thiết lập khởi động cùng Windows: {ex.Message}");
            }
        }
    }
}
