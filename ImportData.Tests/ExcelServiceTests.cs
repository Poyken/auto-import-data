using System;
using System.IO;
using Xunit;
using ImportData.Services;

namespace ImportData.Tests
{
    public class ExcelServiceTests
    {
        [Fact]
        public void ReadExcelFile_ShouldReturnNull_AndLog_WhenFileNotFound()
        {
            // Arrange
            var service = new ExcelService(msg => {
                Assert.Contains("Lỗi khi đọc file Excel", msg);
            });
            string nonExistentHtml = "C:\\InvalidPath\\file_not_exist_xyz.xlsx";

            // Act
            var dt = service.ReadExcelFile(nonExistentHtml);

            // Assert
            Assert.Null(dt);
        }

        // Test này sẽ mock tạo ra file .xlsx tạm, tuy nhiên thư viện ExcelDataReader
        // cần nội dung chuẩn, thay vì vậy chúng ta có thể dừng ở mức return null khi lỗi là đã an toàn.
        
        [Fact]
        public void ReadExcelFile_ShouldCatchExceptionAndReturnNull_WhenFileIsInvalid()
        {
            // Arrange
            string tempFile = Path.GetTempFileName();
            File.WriteAllText(tempFile, "Nội dung vớ vẩn không phải .xlsx"); // Cố tình gây lỗi

            bool errorLogged = false;
            var service = new ExcelService(msg => {
                if (msg.Contains("Lỗi khi đọc file Excel"))
                {
                    errorLogged = true;
                }
            });

            // Act
            var exception = Record.Exception(() => 
            {
                var dt = service.ReadExcelFile(tempFile);
                Assert.Null(dt);
            });

            // Assert
            Assert.Null(exception); // App không được văng Exception
            Assert.True(errorLogged, "Phải ghi nhận lỗi qua Logger");

            // Xóa file tạm
            File.Delete(tempFile);
        }
    }
}
