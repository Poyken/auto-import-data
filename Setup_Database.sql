/* 
   SCRIPT CẬP NHẬT DATABASE CHO HỆ THỐNG AUTO IMPORT (TIỀN XỬ LÝ CHO MÁY THỰC TẾ)
   -------------------------------------------------------------------------
   Nhiệm vụ: 
   1. Kiểm tra và bổ sung cột LotNo vào bảng CapacitorLogs đã có sẵn.
   2. Kiểm tra và bổ sung các cột quản lý (FilePath, ImportDate) nếu thiếu.
   3. Tạo bảng ImportHistory để App nhận diện file đã nạp, tránh trùng dữ liệu.
*/

-- Chọn Database của bạn
USE CapacitorDB; 
GO

-- 1. THÊM CỘT [LotNo] (Dành cho dữ liệu lô sản xuất mới)
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('CapacitorLogs') AND name = 'LotNo')
BEGIN
    ALTER TABLE [dbo].[CapacitorLogs] ADD [LotNo] [nvarchar](100) NULL;
    PRINT 'Da them cot LotNo vao bang CapacitorLogs.';
END
GO

-- 2. ĐẢM BẢO CÓ CỘT [FilePath] VÀ [ImportDate] (Để lưu vết nạp dữ liệu)
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('CapacitorLogs') AND name = 'FilePath')
BEGIN
    ALTER TABLE [dbo].[CapacitorLogs] ADD [FilePath] [nvarchar](500) NULL;
END

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('CapacitorLogs') AND name = 'ImportDate')
BEGIN
    ALTER TABLE [dbo].[CapacitorLogs] ADD [ImportDate] [datetime] DEFAULT GETDATE();
END
GO

-- 3. TẠO BẢNG [ImportHistory] (Nhật ký file đã nạp)
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ImportHistory]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[ImportHistory](
        [Id] [int] IDENTITY(1,1) PRIMARY KEY,
        [FileName] [nvarchar](500) UNIQUE,
        [FilePath] [nvarchar](max) NULL,
        [ImportTime] [datetime] DEFAULT GETDATE()
    );
    PRINT 'Da tao bang ImportHistory thanh cong.';
END
GO

PRINT '>>> [OK] Database da san sang de chay cung App moi! <<<';
