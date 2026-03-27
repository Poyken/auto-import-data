/* 
   HƯỚNG DẪN: SCRIPT TẠO BẢNG DỮ LIỆU CHO HỆ THỐNG AUTO IMPORT
   ----------------------------------------------------------
   Mục tiêu: Tạo bảng SortingDataImportExcel (theo cấu trúc máy đo)
   và bảng ImportHistory để theo dõi lịch sử nạp file.
*/

-- Chọn Database mục tiêu (Thay đổi nếu dùng tên khác)
USE SmartFactoryV2; 
GO

-- 1. TẠO BẢNG CHỨA DỮ LIỆU MÁY ĐO [SortingDataImportExcel]
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SortingDataImportExcel]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[SortingDataImportExcel](
        [Id] [int] IDENTITY(1,1) PRIMARY KEY,
        [EquipmentNumber] [nvarchar](100) NULL,
        [SorterNum] [nvarchar](100) NULL,
        [StartTime] [datetime] NULL,
        [WorkflowCode] [nvarchar](100) NULL,
        [LotNo] [nvarchar](100) NULL, -- Cột bổ sung theo yêu cầu
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
        [FilePath] [nvarchar](500) NULL, -- Cột quản lý (đường dẫn file)
        [ImportDate] [datetime] DEFAULT GETDATE() -- Cột quản lý (ngày nạp)
    );
    PRINT 'Đã tạo bảng SortingDataImportExcel thành công.';
END
GO

-- 2. TẠO BẢNG THEO DÕI LỊCH SỬ [ImportHistory]
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ImportHistory]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[ImportHistory](
        [Id] [int] IDENTITY(1,1) PRIMARY KEY,
        [FileName] [nvarchar](500) UNIQUE,
        [FilePath] [nvarchar](max) NULL,
        [ImportTime] [datetime] DEFAULT GETDATE()
    );
    PRINT 'Đã tạo bảng ImportHistory thành công.';
END
GO

PRINT '------------------------------------------------------------';
PRINT '>>> [HOÀN TẤT] Database đã sẵn sàng chạy cùng ứng dụng mới! <<<';
PRINT '------------------------------------------------------------';

