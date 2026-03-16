using Microsoft.Extensions.Configuration; // Gọi thư viện của Microsoft để có công cụ đọc file cấu hình json.
using System; // Gọi bộ thư viện gốc của C# để dùng các kiểu cơ bản (String, Exception).
using System.IO; // Gọi thư viện hệ thống để tương tác với File và Ổ cứng (đọc/ghi file).

namespace ImportData.Core // Khai báo "địa chỉ nhà" (namespace) cho class này nằm ở khu vực Core của dự án.
{
    /// <summary>
    /// Lớp AppConfig: Quản lý cấu hình hệ thống.
    /// Giúp App biết "Kết nối DB nào?" và "Quét thư mục nào?".
    /// </summary>
    public class AppConfig // Tạo một bản thiết kế (class) mang tên AppConfig. Chữ public để các file khác lấy ra dùng được.
    {
        // Tạo một hằng số (const - không thể sửa đổi khi chạy) chứa đường dẫn dự phòng tới cơ sở dữ liệu nếu file appsettings.json bị xóa mất.
        private const string DEFAULT_DB_CONN = @"Server=.;Database=CapacitorDB;Integrated Security=True;TrustServerCertificate=True;";
        
        // Tạo hằng số chứa tên thư mục quét dự phòng là "task".
        private const string DEFAULT_FOLDER = @"task";

        public string ConnectionString { get; set; } // Tạo một cái hộp (biến) tên là ConnectionString để lưu địa chỉ SQL thực tế rốt cuộc app sẽ xài.
        public string BaseFolder { get; set; }      // Tạo một cái hộp (biến) tên là BaseFolder để lưu đường dẫn thư mục sẽ giám sát.

        /// <summary>
        /// Khởi tạo mặc định ban đầu. Hàm này sẽ tự chạy đầu tiên khi gõ "new AppConfig()".
        /// </summary>
        public AppConfig() 
        {
            ConnectionString = DEFAULT_DB_CONN; // Gán địa chỉ SQL dự phòng vào biến chính, phòng hờ xíu nữa đọc file bị lỗi thì vẫn có cái xài tạm.
            
            // Hỏi hệ điều hành Windows xem Desktop nằm ở đâu, rồi lấy nó ghép với chữ "task" để làm thư mục giám sát dự phòng.
            BaseFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), DEFAULT_FOLDER); 
        }

        /// <summary>
        /// Hàm này dùng để nạp file cấu hình appsettings.json lên.
        /// Cho phép nạp lại khi App đang chạy (nhờ quét timer ngầm bên ngoài).
        /// </summary>
        public void Load(Action<string> logger) // Nhận 1 cái phễu gửi tin (logger) để xíu nữa có gì nó nhờ Form1 in chữ lên màn hình.
        {
            try // Báo cho C# biết: "Hãy thử chạy đoạn code này đi. Lỡ như đọc file rách thì đừng sụp phần mềm nhé".
            {
                // Tìm xem cái thư mục gốc chứa file đuôi .exe đang chạy này thực tế đang nằm phân vùng nào.
                string runDir = AppDomain.CurrentDomain.BaseDirectory; 
                
                // Lùi ra ngoài 3 thư mục mảng (quét ra khỏi bin/Debug/net...). Lý do: khi lập trình hay bị tách ra thư mục Debug và thư mục gốc mã nguồn.
                string sourceDir = Path.GetFullPath(Path.Combine(runDir, @"..\..\..\")); 

                // Lấy đường dẫn thư mục hiện tại ghép với "appsettings.json" để trỏ tới file chạy debug bên trong.
                string settingsPath = Path.Combine(runDir, "appsettings.json"); 
                
                // Trỏ tới file cấu hình nằm ở thư mục gốc (áp dụng khi đang code trong Visual Studio sửa cho tiện).
                string sourceSettingsPath = Path.Combine(sourceDir, "appsettings.json"); 

                // Hỏi Windows: "Cái file ở thư mục gốc kia có còn sống (tồn tại) ko?". Có thì xài nó, không thì lấy file ở thư mục đang chạy (settingsPath).
                string targetFile = File.Exists(sourceSettingsPath) ? sourceSettingsPath : settingsPath; 

                if (File.Exists(targetFile)) // Lại hỏi nhẹ Windows: Rốt cuộc ông CÓ Thực Sự nhìn thấy file cấu hình trên ổ cứng không? CÓ mới đi làm tiếp.
                {
                    string json; // Tạo sẵn một cái hộp (biến) tên json để xíu nữa gắp đống chữ đọc từ file thả vào đây.
                    
                    // Sinh ra một cái vòi hút File (FileStream). Dùng chế độ OPEN (chỉ mở để ngó), READ (chỉ đọc) và READWRITE (Tuyệt kỹ cho OS biết: Đứa nào thích mở Notepad file này sửa tôi cho sửa luôn, khỏi khóa).
                    using (var fs = new FileStream(targetFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) 
                    
                    // Gắn vào vòi hút một cục giải mã (StreamReader) để nó dịch số nhị phân (010101) ổ cứng thành kí tự văn bản con người (utf-8).
                    using (var sr = new StreamReader(fs)) 
                    {
                        json = sr.ReadToEnd(); // Hút 1 làn khí thật sâu: Đọc cạn toàn bộ các dòng chữ cái trong file từ đầu tới cuối trút vào biến json.
                    } // Gặp dấu ngoặc đóng } này lệnh using tự động gọi hàm vứt bỏ (Dispose), bảo HĐH Windows: "Tôi đọc xong rồi hãy rút ống ra đi, đóng file lại".

                    if (!string.IsNullOrWhiteSpace(json)) // Dò xem cái đoạn chữ json vừa móc dưới dĩa lên có trắng bóc (rỗng) hay không? Có chữ thì mới chạy.
                    {
                        // Máy tính hãy dịch đống chữ json bùi nhùi đó thành cây dữ liệu phân cấp (Document) có cành lá dễ lấy.
                        using (var doc = System.Text.Json.JsonDocument.Parse(json)) 
                        {
                            var root = doc.RootElement; // Cầm vào cái Gốc (Root) của cây Json đó.
                            
                            // Hỏi xem trong Gốc có cái rễ nào tên "ConnectionStrings" không? (TryGetProperty) Nếu có, chọc tay vô lấy tiếp rễ nhỏ mang tên "DefaultConnection".
                            if (root.TryGetProperty("ConnectionStrings", out var connStrings) && 
                                connStrings.TryGetProperty("DefaultConnection", out var defaultConn)) // Đoạn nối dòng trên (Lấy giá trị của nhánh con DefaultConnection).
                            {
                                // Lấy cái cục chữ bên trong thả đổ vào thẻ ConnectionString. Dấu "?? ConnectionString" nghĩa là lỡ vế trước bị Null (ko có gì) thì đổ lại đồ cũ đang có để giữ mạng.
                                ConnectionString = defaultConn.GetString() ?? ConnectionString; 
                            }

                            // Tương tự, lục lõi trong Gốc coi có nhánh mang tên "FolderSettings" không?
                            if (root.TryGetProperty("FolderSettings", out var folderSettings) && 
                                folderSettings.TryGetProperty("BaseFolder", out var baseFld)) // Có thì vạch tiếp lá "BaseFolder" ra lục...
                            {
                                // Kéo chữ đường dẫn folder chứa trong lá đổ vào biến BaseFolder của nhà mình. (?? phòng hờ hỏng file).
                                BaseFolder = baseFld.GetString() ?? BaseFolder; 
                            }
                        } // Đóng chặt cây Json, cho ông kẹ (Garbage Collector) dọn cây rác đó phân giải đi khỏi RAM.
                    }
                    
                    if (logger != null) logger.Invoke($"Đã tải cấu hình từ file {Path.GetFileName(targetFile)}"); // Nếu cục truyền tin (logger) còn xài được thì hô Form1 in lên dòng xanh: Đã tải...
                }
            }
            catch (Exception) // Hứng trọn 100% mọi Lỗi ngớ ngẩn sinh ra. (Như việc người ta gõ sai cú pháp ngoặc kép của file Json).
            {
                // Tỏ Vẻ Ngó Lơ, Lập lờ cho qua. App sẽ Không Bị Sập, nó ôm cấu hình cũ đi tiếp. Sẽ đợi tới nhịp canh 10s sau lúc họ điền json đúng Hợp lệ thì tự nạp tự sống.
            }
        }
    }
}
