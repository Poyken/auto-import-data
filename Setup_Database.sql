/* 
   SCRIPT KHỞI TẠO CƠ SỞ DỮ LIỆU VÀ CÁC BẢNG CHO HỆ THỐNG AUTO IMPORT
   -------------------------------------------------------------------
   1. Hướng dẫn: Chạy script này trên SQL Server Management Studio (SSMS).
   2. Mục tiêu:
      - Tạo Database CapacitorDB để quản lý dữ liệu.
      - Tạo bảng CapacitorLogs: Lưu trữ các thông số kỹ thuật đọc được từ Excel.
      - Tạo bảng ImportHistory: Lưu vết lịch sử các file đã nhập để tránh trùng lặp.
*/

-- Kiểm tra và tạo Database nếu chưa tồn tại
IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = 'CapacitorDB')
BEGIN
    CREATE DATABASE CapacitorDB;
END
GO

USE CapacitorDB;
GO

-- 1. BẢNG CapacitorLogs: Lưu dữ liệu chi tiết từ các file Excel kết quả kiểm tra tụ điện
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CapacitorLogs]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[CapacitorLogs](
        [Id] [int] IDENTITY(1,1) PRIMARY KEY,       -- Khóa chính tự tăng
        [EquipmentNumber] [nvarchar](100) NULL,      -- Mã số thiết bị
        [SorterNum] [nvarchar](50) NULL,             -- Số máy phân loại
        [StartTime] [datetime] NULL,                 -- Thời gian bắt đầu test
        [WorkflowCode] [nvarchar](100) NULL,         -- Mã quy trình (Workflow)
        [LotNo] [nvarchar](100) NULL,                -- Số lô sản xuất (Lot)
        [Barcode] [nvarchar](100) NULL,              -- Mã vạch sản phẩm
        [Slot] [nvarchar](50) NULL,                  -- Vị trí khay (Slot)
        [Position] [nvarchar](50) NULL,              -- Vị trí trong khay (ví dụ: 01-04)
        [Channel] [nvarchar](50) NULL,               -- Kênh kiểm tra (ví dụ: 01-04)
        [Capacity_mAh] [float] NULL,                 -- Dung lượng (mAh)
        [Capacitance_F] [float] NULL,                -- Điện dung (F)
        [BeginVoltageSD_mV] [float] NULL,            -- Điện áp bắt đầu SD (mV)
        [ChargeEndCurrent_mA] [float] NULL,          -- Dòng điện kết thúc sạc (mA)
        [EndVoltage_mV] [float] NULL,                -- Điện áp kết thúc (mV)
        [EndCurrent_mA] [float] NULL,                -- Dòng điện kết thúc (mA)
        [DischargeVoltage1_mV] [float] NULL,         -- Điện áp xả lần 1 (mV)
        [DischargeVal1_Time] [nvarchar](50) NULL,    -- Thời gian xả lần 1
        [DischargeVoltage2_mV] [float] NULL,         -- Điện áp xả lần 2 (mV)
        [DischargeVal2_Time] [nvarchar](50) NULL,    -- Thời gian xả lần 2
        [DischargeBeginVoltage_mV] [float] NULL,     -- Điện áp bắt đầu xả
        [DischargeBeginCurrent_mA] [float] NULL,     -- Dòng điện bắt đầu xả
        [NGInfo] [nvarchar](max) NULL,               -- Thông tin lỗi (nếu có)
        [EndTime] [datetime] NULL,                   -- Thời gian kết thúc test
        [ImportDate] [datetime] DEFAULT GETDATE()    -- Ngày hệ thống tự động nhập vào DB
    );
END
GO

-- 2. BẢNG ImportHistory: Quản lý danh sách các file đã được nhập thành công
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ImportHistory]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[ImportHistory](
        [Id] [int] IDENTITY(1,1) PRIMARY KEY,       -- Khóa chính
        [FileName] [nvarchar](500) UNIQUE,           -- Tên file (Dùng để kiểm tra trùng lặp)
        [FilePath] [nvarchar](max) NULL,             -- Đường dẫn đầy đủ đến file tại thời điểm nhập
        [ImportTime] [datetime] DEFAULT GETDATE()    -- Thời điểm thực hiện nhập file
    );
END
ELSE
BEGIN
    -- Kiểm tra và tự động nâng cấp bảng nếu thiếu cột FilePath (Hỗ trợ tương thích ngược)
    IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('ImportHistory') AND name = 'FilePath')
    BEGIN
        ALTER TABLE ImportHistory ADD FilePath NVARCHAR(MAX);
    END
END
GO

PRINT 'Hoan thanh khoi tao Database va Table voi day du comment!';
