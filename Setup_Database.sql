/* 
   HƯỚNG DẪN: SCRIPT TẠO BẢNG DỮ LIỆU CHO HỆ THỐNG AUTO IMPORT (BẢN FINAL)
   ----------------------------------------------------------
   Mục tiêu: Tạo 2 bảng mới với tên độc nhất để tránh trùng lặp dữ liệu cũ.
*/

-- Chọn Database mục tiêu
USE SmartFactoryV2; 
GO

-- 1. TẠO BẢNG DỮ LIỆU MÁY ĐO [ExcelData_Import] 
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ExcelData_Import]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[ExcelData_Import](
        [Id] [int] IDENTITY(1,1) PRIMARY KEY,
        [EquipmentNumber] [nvarchar](100) NULL,
        [SorterNum] [nvarchar](100) NULL,
        [StartTime] [datetime] NULL,
        [WorkflowCode] [nvarchar](100) NULL,
        [LotNo] [nvarchar](100) NULL,
        [Barcode] [nvarchar](100) NULL,
        [Slot] [nvarchar](10) NULL,
        [Position] [nvarchar](50) NULL,
        [Channel] [int] NULL,
        [Capacity_mAh] [nvarchar](50) NULL,
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
        [NGInfo] [nvarchar](MAX) NULL,
        [EndTime] [datetime] NULL,
        [FilePath] [nvarchar](500) NULL,
        [ImportDate] [datetime] DEFAULT GETDATE()
    );
    PRINT 'Đã tạo bảng ExcelData_Import thành công.';
END
GO

-- 2. TẠO BẢNG THEO DÕI LỊCH SỬ [ExcelHistory_Import]
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ExcelHistory_Import]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[ExcelHistory_Import](
        [Id] [int] IDENTITY(1,1) PRIMARY KEY,
        [FilePath] [nvarchar](500) UNIQUE,
        [FileSize] [bigint] NULL,
        [ImportedAt] [datetime2](7) DEFAULT GETDATE(),
        [RowsInserted] [int] NULL,
        [Status] [nvarchar](30) NULL,
        [ErrorMessage] [nvarchar](MAX) NULL
    );
    PRINT 'Đã tạo bảng ExcelHistory_Import thành công.';
END
GO

PRINT '------------------------------------------------------------';
PRINT '>>> [HOÀN TẤT] Database đã sẵn sàng chạy cùng ứng dụng mới! <<<';
PRINT '------------------------------------------------------------';

