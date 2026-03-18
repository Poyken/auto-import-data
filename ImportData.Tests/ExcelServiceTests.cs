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
        /// TEST 1: Nếu đường dẫn file không tồn tại, hàm phải trả về null và ghi log báo lỗi.
        /// </summary>
        [Fact]
        public void ReadExcelFile_ShouldReturnNull_AndLog_WhenFileNotFound()
        {
            // 1. Chuẩn bị (Arrange): Tạo dịch vụ Excel và một đường dẫn "ma" không hề tồn tại.
            var service = new ExcelService(msg => {
                // Kiểm tra xem trong chuỗi Log có chứa từ khóa báo lỗi đọc file không.
                Assert.Contains("[ERROR] Excel read", msg);
            });
            
            string nonExistentPath = "C:\\Duong_Dan_Gia\\file_khong_co_that.xlsx";

            // 2. Hành động (Act): Thử đọc file không tồn tại đó.
            var dt = service.ReadExcelFile(nonExistentPath);

            // 3. Kiểm tra (Assert):
            // Kết quả trả về phải là Null (Không có bảng dữ liệu nào được sinh ra).
            Assert.Null(dt);
        }

        /// <summary>
        /// TEST 2: Nếu file không phải định dạng Excel hợp lệ, hệ thống phải bắt được lỗi và trả về null thay vì bị sập.
        /// </summary>
        [Fact]
        public void ReadExcelFile_ShouldCatchExceptionAndReturnNull_WhenFileIsInvalid()
        {
            // 1. Chuẩn bị (Arrange): 
            // Tạo một file tạm .tmp nhưng ghi nội dung là văn bản bình thường (không phải cấu trúc nhị phân của Excel).
            string tempFile = Path.GetTempFileName();
            File.WriteAllText(tempFile, "Day la noi dung van ban, be ngoai la file nhung ben trong khong phai Excel.");

            bool errorLogged = false;
            var service = new ExcelService(msg => {
                // Đánh dấu nếu App có gọi hàm ghi log báo lỗi đọc Excel.
                if (msg.Contains("[ERROR] Excel read")) errorLogged = true;
            });

            // 2. Hành động (Act):
            // Thử đọc file rác đó và dùng Record.Exception để xem có lỗi Exception nào văng ra làm sập App không.
            var exception = Record.Exception(() => 
            {
                var dt = service.ReadExcelFile(tempFile);
                Assert.Null(dt); // Kết quả đọc file rác phải là Null.
            });

            // 3. Kiểm tra (Assert):
            Assert.Null(exception); // App phải chạy xuyên suốt, không được nổ lỗi Exception.
            Assert.True(errorLogged); // Phải có dòng Log báo lỗi xuất hiện cho người dùng biết.

            // Dọn dẹp: Xóa file tạm sau khi đã Test xong để đỡ rác ổ cứng.
            if (File.Exists(tempFile)) File.Delete(tempFile);
        }
    }
}

