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
        /// Test: Nếu không có file JSON, App phải dùng giá trị mặc định.
        /// </summary>
        [Fact]
        public void AppConfig_Load_ShouldFallbackToDefault_IfNoFileExists()
        {
            // Arrange (Chuẩn bị)
            if (File.Exists(_testSettingsPath)) File.Delete(_testSettingsPath);
            var config = new AppConfig();
            
            // defaultConn ví dụ: "Server=.;Database=CapacitorDB;..."
            string defaultConn = config.ConnectionString;

            // Act (Hành động)
            bool logCalled = false;
            config.Load(msg => {
                if (msg.Contains("dùng cấu hình mặc định")) logCalled = true;
            });

            // Assert (Kiểm tra kết quả)
            Assert.True(logCalled);
            Assert.Equal(defaultConn, config.ConnectionString);
        }

        /// <summary>
        /// Test: Nếu file JSON hợp lệ, App phải nạp đúng giá trị trong file đó.
        /// </summary>
        [Fact]
        public void AppConfig_Load_ShouldOverrideValues_IfFileExistsAndValid()
        {
            // Arrange
            // validJson ví dụ: '{"ConnectionStrings": {"DefaultConnection": "Server=TEST_SERVER;..."}}'
            string validJson = @"{
                ""ConnectionStrings"": {
                    ""DefaultConnection"": ""Server=TEST_SERVER;Database=TestDB;Integrated Security=True;TrustServerCertificate=True;""
                },
                ""FolderSettings"": {
                    ""BaseFolder"": ""C:\\TestFolder""
                }
            }";
            
            File.WriteAllText(_testSettingsPath, validJson);
            var config = new AppConfig();

            // Act
            config.Load(msg => { });

            // Assert
            // Kiểm tra ConnectionString có đúng là "Server=TEST_SERVER;..." không
            Assert.Equal("Server=TEST_SERVER;Database=TestDB;Integrated Security=True;TrustServerCertificate=True;", config.ConnectionString);
            Assert.Equal("C:\\TestFolder", config.BaseFolder);
        }

        /// <summary>
        /// Test: Nếu file JSON bị lỗi cú pháp, App không được sập (Crash).
        /// </summary>
        [Fact]
        public void AppConfig_Load_ShouldNotCrash_IfJsonIsInvalid()
        {
            // Arrange
            // invalidJson bị thiếu dấu ngoặc đóng }
            string invalidJson = @"{ ""ConnectionStrings"": { ""DefaultConnection"": ""ABC"" "; 
            
            File.WriteAllText(_testSettingsPath, invalidJson);
            var config = new AppConfig();
            string originalConn = config.ConnectionString;

            // Act
            var exception = Record.Exception(() => config.Load(msg => { }));

            // Assert
            Assert.Null(exception); // Không được có lỗi văng ra
            Assert.Equal(originalConn, config.ConnectionString); // Dữ liệu cũ phải được giữ nguyên
        }
    }
}

