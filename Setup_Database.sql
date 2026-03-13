/* 
   SCRIPT KHỞI TẠO DATABASE VÀ CÁC BẢNG CHO HỆ THỐNG AUTO IMPORT
   -----------------------------------------------------------
   1. Chạy script này trên SQL Server Management Studio (SSMS).
   2. Script sẽ tự động tạo Database CapacitorDB nếu chưa có.
   3. Tạo bảng CapacitorLogs (Lưu dữ liệu Excel).
   4. Tạo bảng ImportHistory (Lưu vết các file đã xử lý).
*/

IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = 'CapacitorDB')
BEGIN
    CREATE DATABASE CapacitorDB;
END
GO

USE CapacitorDB;
GO

-- 1. Bảng lưu trữ dữ liệu chi tiết từ Excel
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CapacitorLogs]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[CapacitorLogs](
        [Id] [int] IDENTITY(1,1) PRIMARY KEY,
        [EquipmentNumber] [nvarchar](100) NULL,
        [SorterNum] [nvarchar](50) NULL,
        [StartTime] [datetime] NULL,
        [WorkflowCode] [nvarchar](100) NULL,
        [LotNo] [nvarchar](100) NULL,
        [Barcode] [nvarchar](100) NULL,
        [Slot] [nvarchar](50) NULL,
        [Position] [int] NULL,
        [Channel] [int] NULL,
        [Capacity_mAh] [float] NULL,
        [Capacitance_F] [float] NULL,
        [BeginVoltageSD_mV] [float] NULL,
        [ChargeEndCurrent_mA] [float] NULL,
        [EndVoltage_mV] [float] NULL,
        [EndCurrent_mA] [float] NULL,
        [DischargeVoltage1_mV] [float] NULL,
        [DischargeVal1_Time] [nvarchar](20) NULL,
        [DischargeVoltage2_mV] [float] NULL,
        [DischargeVal2_Time] [nvarchar](20) NULL,
        [DischargeBeginVoltage_mV] [float] NULL,
        [DischargeBeginCurrent_mA] [float] NULL,
        [NGInfo] [nvarchar](max) NULL,
        [EndTime] [datetime] NULL,
        [ImportDate] [datetime] DEFAULT GETDATE()
    );
END
GO

-- 2. Bảng lưu lịch sử các file đã Import thành công
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ImportHistory]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[ImportHistory](
        [Id] [int] IDENTITY(1,1) PRIMARY KEY,
        [FileName] [nvarchar](500) UNIQUE,
        [FilePath] [nvarchar](max) NULL,
        [ImportTime] [datetime] DEFAULT GETDATE()
    );
END
ELSE
BEGIN
    -- Kiểm tra nếu bảng cũ thiếu cột FilePath thì tự động bổ sung
    IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('ImportHistory') AND name = 'FilePath')
    BEGIN
        ALTER TABLE ImportHistory ADD FilePath NVARCHAR(MAX);
    END
END
GO

PRINT 'Khoi tao Database va Table thanh cong!';
