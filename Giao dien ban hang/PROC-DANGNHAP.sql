USE QUANLYSIEUTHI
GO 
-- Query cho đăng nhập:
-- Insert dữ liệu:
INSERT INTO dbo.nguoidung
(
    Tennguoidung,
    gmail,
    Matkhau,
    Quyenhan,
    Diachi,
    SDT,
    Luong,
    Sogiolam,
    Manhanvien
)
VALUES
(   N'Nguyễn Văn Thương',  -- Tennguoidung - nvarchar(50)
    N'nguyenvanthuong@gmail.com',  -- gmail - nvarchar(50)
    '1',   -- Matkhau - varchar(50)
    N'admin',  -- Quyenhan - nvarchar(50)
    N'Thanh Hóa',  -- Diachi - nvarchar(200)
    123456,    -- SDT - int
    123456, -- Luong - money
    6969,    -- Sogiolam - int
    )
INSERT INTO dbo.nguoidung
(
    Tennguoidung,
    gmail,
    Matkhau,
    Quyenhan,
    Diachi,
    SDT,
    Luong,
    Sogiolam,
    Manhanvien
)
VALUES
(   N'Mai Xuân Thưởng',  -- Tennguoidung - nvarchar(50)
    N'maixuanthuong@gmail.com',  -- gmail - nvarchar(50)
    '2',   -- Matkhau - varchar(50)
    N'Staff',  -- Quyenhan - nvarchar(50)
    N'LongAn',  -- Diachi - nvarchar(200)
    123456,    -- SDT - int
    123456, -- Luong - money
    6969,    -- Sogiolam - int
  
    )
	INSERT INTO dbo.nguoidung
(
    Tennguoidung,
    gmail,
    Matkhau,
    Quyenhan,
    Diachi,
    SDT,
    Luong,
    Sogiolam,
    Manhanvien
)
VALUES
(   N'Nguyễn Ngọc Sơn',  -- Tennguoidung - nvarchar(50)
    N'nguyenngocson@gmail.com',  -- gmail - nvarchar(50)
    '3',   -- Matkhau - varchar(50)
    N'staff',  -- Quyenhan - nvarchar(50)
    N'Gia lai',  -- Diachi - nvarchar(200)
    123456,    -- SDT - int
    123456, -- Luong - money
    6969,    -- Sogiolam - int
  
    )
	INSERT INTO dbo.nguoidung
(
    Tennguoidung,
    gmail,
    Matkhau,
    Quyenhan,
    Diachi,
    SDT,
    Luong,
    Sogiolam,
    Manhanvien
)
VALUES
(   N'Nguyễn Khánh Duy',  -- Tennguoidung - nvarchar(50)
    N'nguyenkhanhduy@gmail.com',  -- gmail - nvarchar(50)
    '4',   -- Matkhau - varchar(50)
    N'staff',  -- Quyenhan - nvarchar(50)
    N'Hà Tĩnh',  -- Diachi - nvarchar(200)
    123456,    -- SDT - int
    123456, -- Luong - money
    6969,    -- Sogiolam - int
  
    )
	INSERT INTO dbo.nguoidung
(
    Tennguoidung,
    gmail,
    Matkhau,
    Quyenhan,
    Diachi,
    SDT,
    Luong,
    Sogiolam,
    Manhanvien
)
VALUES
(   N'Huỳnh Công Duy',  -- Tennguoidung - nvarchar(50)
    N'huynhcongduy@gmail.com',  -- gmail - nvarchar(50)
    '5',   -- Matkhau - varchar(50)
    N'staff',  -- Quyenhan - nvarchar(50)
    N'Long An',  -- Diachi - nvarchar(200)
    123456,    -- SDT - int
    123456, -- Luong - money
    6969,    -- Sogiolam - int
 
    )

SELECT * FROM dbo.nguoidung
-- PROC:
GO 
CREATE PROC proc_dangnhap 
@userGmail varchar(50), @passWork varchar(50)
AS 
BEGIN
	SELECT * FROM dbo.nguoidung WHERE @userGmail=gmail AND @passWork= Matkhau
END
GO
EXEC dbo.proc_dangnhap @userGmail = '', -- varchar(50)
                       @passWork = ''   -- varchar(50)
