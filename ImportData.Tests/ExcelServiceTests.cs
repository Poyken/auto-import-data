using System;
using System.IO;
using Xunit;
using ImportData.Services;

namespace ImportData.Tests
{
    /// <summary>
    /// Lớp kiểm thử (Unit Test) cho ExcelService.
    /// Kiểm tra tính bền bỉ của trình đọc Excel khi gặp file lỗi hoặc mất file.
    /// </summary>
    public class ExcelServiceTests
    {
        /// <summary>
        /// Test: Nếu đường dẫn file không tồn tại, hàm phải trả về null và ghi log lỗi.
        /// </summary>
        [Fact]
        public void ReadExcelFile_ShouldReturnNull_AndLog_WhenFileNotFound()
        {
            // Arrange
            var service = new ExcelService(msg => {
                // Kiểm tra xem log có chứa nội dung báo lỗi không
                Assert.Contains("Lỗi đọc file Excel", msg);
            });
            // nonExistentPath ví dụ: "C:\Windows\System32\khong_co_file.xlsx"
            string nonExistentPath = "C:\\InvalidPath\\file_not_exist_xyz.xlsx";

            // Act
            var dt = service.ReadExcelFile(nonExistentPath);

            // Assert
            Assert.Null(dt);
        }

        /// <summary>
        /// Test: Nếu file không phải định dạng Excel (.xlsx), hệ thống phải bắt được lỗi và trả về null.
        /// </summary>
        [Fact]
        public void ReadExcelFile_ShouldCatchExceptionAndReturnNull_WhenFileIsInvalid()
        {
            // Arrange
            // tempFile ví dụ: "C:\Users\...\AppData\Local\Temp\tmpA1B2.tmp"
            string tempFile = Path.GetTempFileName();
            File.WriteAllText(tempFile, "Nội dung văn bản thông thường, không phải mã nhị phân Excel");

            bool errorLogged = false;
            var service = new ExcelService(msg => {
                if (msg.Contains("Lỗi đọc file Excel")) errorLogged = true;
            });

            // Act
            var exception = Record.Exception(() => 
            {
                var dt = service.ReadExcelFile(tempFile);
                Assert.Null(dt);
            });

            // Assert
            Assert.Null(exception); // Đảm bảo App không bị vỡ (Crash)
            Assert.True(errorLogged);

            // Xóa file tạm sau khi test xong
            File.Delete(tempFile);
        }
    }
}

