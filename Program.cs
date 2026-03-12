namespace ImportData
{
    // Đây là file cốt lõi nhất của toàn bộ chương trình, nơi phần mềm quyết định chạy theo trình tự nào.
    internal static class Program
    {
        /// <summary>
        ///  Điểm vào (Entry point) chính của Ứng dụng. 
        ///  Khi bạn click đúp vào file .exe, nó sẽ tìm đến hàm Main() này để chạy đầu tiên.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Cấu hình nâng cao cho app (ví dụ làm sắc nét font chữ, DPI, ... ở .NET mới)
            ApplicationConfiguration.Initialize();

            // Lệnh quan trọng nhất: Ra lệnh cho Windows "Hãy khởi tạo và bật cái màn hình Form1 lên, giữ cho app chạy liên tục"
            Application.Run(new Form1());
        }
    }
}