using System;
using System.IO;
using Xunit;
using ImportData.Core;

namespace ImportData.Tests
{
    /// <summary>
    /// Lớp kiểm thử (Unit Test) cho AppConfig.
    /// Đảm bảo việc nạp cấu hình luôn chính xác và không gây sập App.
    /// </summary>
    public class AppConfigTests : IDisposable
    {
        private readonly string _testSettingsPath;

        public AppConfigTests()
        {
            // _testSettingsPath ví dụ: "C:\...\ImportData.Tests\bin\Debug\appsettings.json"
            _testSettingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
        }

        public void Dispose()
        {
            // Dọn dẹp: Xóa file cấu hình giả lập sau mỗi bài test để không ảnh hưởng bài sau
            if (File.Exists(_testSettingsPath))
            {
                File.Delete(_testSettingsPath);
            }
        }

        /// <summary>
        /// TEST 1: Nếu không có file cấu hình appsettings.json, App phải tự dùng thông số mặc định.
        /// </summary>
        [Fact]
        public void AppConfig_Load_ShouldFallbackToDefault_IfNoFileExists()
        {
            // 1. Chuẩn bị (Arrange): Xóa file appsettings.json nếu nó đang tồn tại lù lù ở đó.
            if (File.Exists(_testSettingsPath)) File.Delete(_testSettingsPath);
            
            var config = new AppConfig(); // Tạo mới đối tượng cấu hình.
            
            // Lưu lại chuỗi kết nối mặc định mà App vừa tự sinh ra trong hàm dựng (Constructor).
            string defaultConn = config.ConnectionString;

            // 2. Hành động (Act): Gọi lệnh nạp (Load).
            bool logCalled = false;
            config.Load(msg => {
                // Kiểm tra xem App có báo là "Cấu hình đã nạp" hay không (Hàm Load in log khi tìm thấy file).
                // Ở đây ta xóa file rồi, nên log sẽ không được gọi (hoặc gọi lỗi).
                logCalled = true; 
            });

            // 3. Kiểm tra (Assert):
            // Vì file không có, giá trị ConnectionString phải Y HỆT giá trị mặc định ban đầu.
            Assert.Equal(defaultConn, config.ConnectionString);
        }

        /// <summary>
        /// TEST 2: Nếu file JSON hợp lệ, App phải nạp đúng các giá trị mới thay thế cho mặc định.
        /// </summary>
        [Fact]
        public void AppConfig_Load_ShouldOverrideValues_IfFileExistsAndValid()
        {
            // 1. Chuẩn bị (Arrange): Tạo nội dung file JSON giả lập với các thông số Test.
            string validJson = @"{
                ""ConnectionStrings"": {
                    ""DefaultConnection"": ""Server=SERVER_KIEM_THU;Database=TestDB;Integrated Security=True;TrustServerCertificate=True;""
                },
                ""FolderSettings"": {
                    ""BaseFolder"": ""C:\\Thu_Muc_Test""
                }
            }";
            
            // Ghi nội dung này xuống file vật lý appsettings.json trong thư mục chạy Test.
            File.WriteAllText(_testSettingsPath, validJson);
            
            var config = new AppConfig(); // Khởi tạo AppConfig.

            // 2. Hành động (Act): Nạp file.
            config.Load(msg => { }); // Truyền hàm log trống để không làm phiền màn hình.

            // 3. Kiểm tra (Assert):
            // So sánh xem ConnectionString có đúng là "Server=SERVER_KIEM_THU..." như ta vừa ghi không.
            Assert.Equal("Server=SERVER_KIEM_THU;Database=TestDB;Integrated Security=True;TrustServerCertificate=True;", config.ConnectionString);
            
            // Kiểm tra thư mục quét có đúng tên "Thu_Muc_Test" không.
            Assert.Equal("C:\\Thu_Muc_Test", config.BaseFolder);
        }

        /// <summary>
        /// TEST 3: Nếu file JSON bị sai cú pháp (ví dụ thiếu dấu ngoặc), App không được bị sập (Crash).
        /// </summary>
        [Fact]
        public void AppConfig_Load_ShouldNotCrash_IfJsonIsInvalid()
        {
            // 1. Chuẩn bị (Arrange): Ghi một chuỗi JSON "què cụt", thiếu dấu đóng ngoặc }.
            string invalidJson = @"{ ""ConnectionStrings"": { ""DefaultConnection"": ""LOI_CU_PHAP"" "; 
            
            File.WriteAllText(_testSettingsPath, invalidJson);
            var config = new AppConfig();
            
            // Ghi nhớ giá trị cũ trước khi nạp file lỗi.
            string originalConn = config.ConnectionString;

            // 2. Hành động (Act): Chạy lệnh Load và dùng Record.Exception để bắt lấy lỗi nếu App bị văng Exception.
            var exception = Record.Exception(() => config.Load(msg => { }));

            // 3. Kiểm tra (Assert):
            Assert.Null(exception); // Kết quả phải là Null (Nghĩa là App chạy êm, không bị nổ lỗi ra ngoài).
            
            // Dữ liệu ConnectionString phải giữ nguyên giá trị cũ, không được bị thay đổi bởi file lỗi.
            Assert.Equal(originalConn, config.ConnectionString); 
        }
    }
}

