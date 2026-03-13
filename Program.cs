namespace ImportData
{
    /// <summary>
    /// Lớp Program là điểm khởi đầu của ứng dụng.
    /// Đây là nơi đầu tiên được thực thi khi bạn chạy tệp .exe.
    /// </summary>
    internal static class Program
    {
        /// <summary>
        ///  Điểm vào (Entry point) chính của Ứng dụng. 
        ///  [STAThread] đánh dấu luồng thực thi chính là Single-Threaded Apartment, 
        ///  yêu cầu bắt buộc đối với các ứng dụng giao diện Windows Forms.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Tải cấu hình khởi tạo cho ứng dụng (Font, DPI, Style...)
            ApplicationConfiguration.Initialize();

            // Khởi chạy vòng lặp thông điệp của Windows (Message Loop)
            // Và hiển thị giao diện chính là Form1.
            // Ứng dụng sẽ tiếp tục chạy cho đến khi Form1 được đóng hoàn toàn.
            Application.Run(new Form1());
        }
    }
}