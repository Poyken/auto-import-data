/* 
   SCRIPT KHỞI TẠO LẠI DATABASE GIỐNG HỆT MÁY THỰC TẾ
   Dùng để giả lập môi trường thực tế trên máy Local
*/

USE master;
GO

-- Xóa và tạo mới Database
IF EXISTS (SELECT * FROM sys.databases WHERE name = 'CapacitorDB')
BEGIN
    ALTER DATABASE CapacitorDB SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE CapacitorDB;
END
CREATE DATABASE CapacitorDB;
GO

USE CapacitorDB;
GO

-- 1. TẠO BẢNG CapacitorLogs GIỐNG HỆT HÌNH ẢNH MÁY THỰC TẾ (TURN 11:02)
CREATE TABLE [dbo].[CapacitorLogs](
    [Id] [int] IDENTITY(1,1) PRIMARY KEY,
    [EquipmentNumber] [nvarchar](100) NULL,
    [SorterNum] [nvarchar](100) NULL,
    [StartTime] [datetime] NULL,
    [WorkflowCode] [nvarchar](100) NULL,
    [LotNo] [nvarchar](100) NULL,             -- Cột LotNo bạn tự thêm
    [Barcode] [nvarchar](100) NULL,
    [Slot] [nvarchar](10) NULL,
    [Position] [nvarchar](50) NULL,
    [Channel] [int] NULL,                    -- Hình thực tế để kiểu int
    [Capacity_mAh] [nvarchar](50) NULL,      -- Hình thực tế để nvarchar
    [Capacitance_F] [nvarchar](50) NULL,
    [BeginVoltageSD_mV] [nvarchar](50) NULL,
    [ChargeEndCurrent_mA] [nvarchar](50) NULL,
    [EndVoltage_mV] [nvarchar](50) NULL,
    [EndCurrent_mA] [nvarchar](50) NULL,
    [DischargeVoltage1_mV] [nvarchar](50) NULL,
    [DischargeVoltage1_Time] [nvarchar](50) NULL,
    [DischargeVoltage2_mV] [nvarchar](50) NULL,
    [DischargeVoltage2_Time] [nvarchar](50) NULL,
    [DischargeBeginVoltage_mV] [nvarchar](50) NULL,
    [DischargeBeginCurrent_mA] [nvarchar](50) NULL,
    [NGInfo] [nvarchar](max) NULL,
    [EndTime] [datetime] NULL,
    [FilePath] [nvarchar](500) NULL,         -- Đã có sẵn trong máy thực
    [ImportDate] [datetime] NULL             -- Đã có sẵn trong máy thực
);
GO

-- 2. TẠO BẢNG ImportHistory (Để App chạy được)
CREATE TABLE [dbo].[ImportHistory](
    [Id] [int] IDENTITY(1,1) PRIMARY KEY,
    [FileName] [nvarchar](500) UNIQUE,
    [FilePath] [nvarchar](max) NULL,
    [ImportTime] [datetime] DEFAULT GETDATE()
);
GO

PRINT '>>> [XONG] Da tao lai Database giong het moi truong thuc te! <<<';
