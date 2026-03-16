using System;
using System.IO;
using Xunit;
using ImportData.Core;

namespace ImportData.Tests
{
    public class AppConfigTests : IDisposable
    {
        private readonly string _testSettingsPath;

        public AppConfigTests()
        {
            _testSettingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
        }

        public void Dispose()
        {
            // Dọn dẹp file test sau khi chạy xong
            if (File.Exists(_testSettingsPath))
            {
                File.Delete(_testSettingsPath);
            }
        }

        [Fact]
        public void AppConfig_Load_ShouldFallbackToDefault_IfNoFileExists()
        {
            // Arrange
            if (File.Exists(_testSettingsPath)) File.Delete(_testSettingsPath);
            var config = new AppConfig();
            
            // Lấy giá trị ban đầu để so sánh
            string defaultConn = config.ConnectionString;
            string defaultFolder = config.BaseFolder;

            // Act
            bool logCalled = false;
            config.Load(msg => {
                if (msg.Contains("dùng cấu hình mặc định")) logCalled = true;
            });

            // Assert
            Assert.True(logCalled, "Phải gọi log báo hiệu không tìm thấy file và dùng mặc định.");
            Assert.Equal(defaultConn, config.ConnectionString);
            Assert.Equal(defaultFolder, config.BaseFolder);
        }

        [Fact]
        public void AppConfig_Load_ShouldOverrideValues_IfFileExistsAndValid()
        {
            // Arrange
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
            bool successLog = false;
            config.Load(msg => {
                if (msg.Contains("Đã tải cấu hình")) successLog = true;
            });

            // Assert
            Assert.True(successLog, "Phải có log báo tải thành công.");
            Assert.Equal("Server=TEST_SERVER;Database=TestDB;Integrated Security=True;TrustServerCertificate=True;", config.ConnectionString);
            Assert.Equal("C:\\TestFolder", config.BaseFolder);
        }

        [Fact]
        public void AppConfig_Load_ShouldNotCrash_IfJsonIsInvalid()
        {
            // Arrange
            string invalidJson = @"{
                ""ConnectionStrings"": {
                    ""DefaultConnection"": ""Server=TEST_SERVER;
                }
            }"; // Lỗi cú pháp thiếu dấu ngoặc kép hoặc đóng ngoặc
            
            File.WriteAllText(_testSettingsPath, invalidJson);
            var config = new AppConfig();
            
            string originalConn = config.ConnectionString;

            // Act
            var exception = Record.Exception(() => config.Load(msg => { }));

            // Assert
            Assert.Null(exception); // Phải không được văng lỗi làm sập ứng dụng (Crash)
            Assert.Equal(originalConn, config.ConnectionString); // Phải giữ nguyên kết nối cấu hình gốc
        }
    }
}
