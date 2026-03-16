/* 
   HƯỚNG DẪN: SCRIPT GIẢ LẬP MÔI TRƯỜNG MÁY THỰC TẾ TRÊN MÁY LOCAL
   -------------------------------------------------------------
   Công dụng: Tạo lại từ đầu Database 'CapacitorDB' với cấu trúc bảng 'CapacitorLogs' 
   giống hệt như máy đo thực tế, phục vụ việc lập trình và kiểm thử (Testing).
*/

USE master; -- Chuyển về database hệ thống để thực hiện xóa/tạo mới
GO

-- KIỂM TRA VÀ XÓA DATABASE CŨ (Nếu đã tồn tại)
IF EXISTS (SELECT * FROM sys.databases WHERE name = 'CapacitorDB')
BEGIN
    -- SINGLE_USER: Ngắt kết nối các ứng dụng khác đang dùng Database này để có thể xóa được.
    ALTER DATABASE CapacitorDB SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE CapacitorDB;
    PRINT 'Đã xóa Database cũ.';
END

-- TẠO MỚI DATABASE 'CapacitorDB'
CREATE DATABASE CapacitorDB;
GO

USE CapacitorDB;
GO

-- 1. TẠO BẢNG 'CapacitorLogs'
-- Cấu trúc này được xây dựng dựa trên hình ảnh thực tế từ máy đo.
CREATE TABLE [dbo].[CapacitorLogs](
    [Id] [int] IDENTITY(1,1) PRIMARY KEY,
    [EquipmentNumber] [nvarchar](100) NULL, -- Số hiệu thiết bị
    [SorterNum] [nvarchar](100) NULL,       -- Số hiệu máy phân loại
    [StartTime] [datetime] NULL,           -- Thời gian bắt đầu đo
    [WorkflowCode] [nvarchar](100) NULL,    -- Mã quy trình làm việc
    [LotNo] [nvarchar](100) NULL,           -- Số Lô sản xuất (Cột mới yêu cầu thêm)
    [Barcode] [nvarchar](100) NULL,         -- Mã vạch sản phẩm
    [Slot] [nvarchar](10) NULL,             -- Vị trí khay
    [Position] [nvarchar](50) NULL,         -- Vị trí đo
    [Channel] [int] NULL,                   -- Kênh đo
    [Capacity_mAh] [nvarchar](50) NULL,     -- Dung lượng (mAh)
    [Capacitance_F] [nvarchar](50) NULL,    -- Điện dung (F)
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
    [NGInfo] [nvarchar](max) NULL,          -- Thông tin lỗi (nếu có)
    [EndTime] [datetime] NULL,              -- Thời gian kết thúc đo
    [FilePath] [nvarchar](500) NULL,        -- Lưu vết nguồn gốc File Excel nạp vào
    [ImportDate] [datetime] NULL            -- Thời điểm dữ liệu được nạp vào
);
GO

-- 2. TẠO BẢNG 'ImportHistory'
-- Dùng để quản lý danh sách các tệp Excel đã xử lý thành công.
CREATE TABLE [dbo].[ImportHistory](
    [Id] [int] IDENTITY(1,1) PRIMARY KEY,
    [FileName] [nvarchar](500) UNIQUE,
    [FilePath] [nvarchar](max) NULL,
    [ImportTime] [datetime] DEFAULT GETDATE()
);
GO

PRINT '-------------------------------------------------------------------';
PRINT '>>> [XONG] Đã khởi tạo lại môi trường giả lập SQL Server thành công! <<<';
PRINT '-------------------------------------------------------------------';

