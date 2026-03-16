/* 
   HƯỚNG DẪN: SCRIPT CẬP NHẬT DATABASE CHO HỆ THỐNG AUTO IMPORT
   ----------------------------------------------------------
   Mục tiêu: Đảm bảo bảng CapacitorLogs có đủ các cột cần thiết (LotNo, FilePath, ImportDate)
   và bảng ImportHistory được tạo mới để theo dõi lịch sử nạp file.
*/

-- Chọn Database mục tiêu để thực thi lệnh
USE CapacitorDB; 
GO

-- 1. BỔ SUNG CỘT [LotNo]
-- sys.columns: Bảng hệ thống chứa thông tin toàn bộ cột trong Database.
-- Nếu trong bảng 'CapacitorLogs' chưa có cột tên là 'LotNo' thì mới tiến hành thêm.
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('CapacitorLogs') AND name = 'LotNo')
BEGIN
    ALTER TABLE [dbo].[CapacitorLogs] ADD [LotNo] [nvarchar](100) NULL;
    PRINT 'Đã thêm thành công cột LotNo vào bảng CapacitorLogs.';
END
GO

-- 2. BỔ SUNG CÁC CỘT QUẢN LÝ (FilePath, ImportDate)
-- FilePath: Lưu đường dẫn đầy đủ của file nạp, giúp truy vết sau này.
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('CapacitorLogs') AND name = 'FilePath')
BEGIN
    ALTER TABLE [dbo].[CapacitorLogs] ADD [FilePath] [nvarchar](500) NULL;
    PRINT 'Đã thêm cột FilePath.';
END

-- ImportDate: Lưu ngày giờ thực hiện nạp dữ liệu (Mặc định lấy giờ hệ thống SQL).
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('CapacitorLogs') AND name = 'ImportDate')
BEGIN
    ALTER TABLE [dbo].[CapacitorLogs] ADD [ImportDate] [datetime] DEFAULT GETDATE();
    PRINT 'Đã thêm cột ImportDate.';
END
GO

-- 3. KHỞI TẠO BẢNG [ImportHistory]
-- Bảng này cực kỳ quan trọng: App sẽ dựa vào đây để biết file nào đã nạp rồi, 
-- tránh việc nạp trùng lặp dữ liệu khi App quét lại folder.
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ImportHistory]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[ImportHistory](
        [Id] [int] IDENTITY(1,1) PRIMARY KEY,     -- Mã định danh tự tăng
        [FileName] [nvarchar](500) UNIQUE,        -- Tên file (Duy nhất, không cho phép trùng)
        [FilePath] [nvarchar](max) NULL,          -- Đường dẫn file
        [ImportTime] [datetime] DEFAULT GETDATE() -- Thời gian nạp
    );
    PRINT 'Đã tạo bảng ImportHistory thành công.';
END
GO

PRINT '------------------------------------------------------------';
PRINT '>>> [HOÀN TẤT] Database đã sẵn sàng chạy cùng ứng dụng mới! <<<';
PRINT '------------------------------------------------------------';

