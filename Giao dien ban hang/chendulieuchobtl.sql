SELECT* FROM dbo.Sanpham

	DELETE dbo.Sanpham WHERE IDhang='H0001'
	INSERT INTO [dbo].[Sanpham]
	(
	    IDhang,
    TenHang,
    Loaihang,
    SoLuongtrongkho,
    Sothutu,
    TheTich,
    PICTURE
	)
SELECT 'H0001',N'Bánh London', N'1',20,'1', N'550g',* FROM OPENROWSET(BULK N'C:\Users\Thuong Nhu\Desktop\Giao diện app\Giao dien ban hang\ANHSP\6178_490055.JPG',SINGLE_BLOB)AS IMG
INSERT INTO [dbo].[Sanpham]
	(
	    IDhang,
    TenHang,
    Loaihang,
    SoLuongtrongkho,
    Sothutu,
    TheTich,
    PICTURE
	)
SELECT 'H0002',N'Bánh London2', N'1',30,'1', N'450g',* FROM OPENROWSET(BULK N'C:\Users\Thuong Nhu\Desktop\Giao diện app\Giao dien ban hang\ANHSP\BANH LONDON.JPG',SINGLE_BLOB)AS IMG
INSERT INTO [dbo].[Sanpham]
	(
	    IDhang,
    TenHang,
    Loaihang,
    SoLuongtrongkho,
    Sothutu,
    TheTich,
    PICTURE
	)
SELECT 'H0003',N'Nước SPRIZE', N'2',15,'2', N'450ml',* FROM OPENROWSET(BULK N'C:\Users\Thuong Nhu\Desktop\Giao diện app\Giao dien ban hang\ANHSP\7up-sleek-330ml-2-700x467-1.JPG',SINGLE_BLOB)AS IMG

INSERT INTO [dbo].[Sanpham]
	(
	    IDhang,
    TenHang,
    Loaihang,
    SoLuongtrongkho,
    Sothutu,
    TheTich,
    PICTURE
	)
	SELECT 'H0004',N'Nước SPRIZE', N'2',20,'2', N'1.5L',* FROM OPENROWSET(BULK N'C:\Users\Thuong Nhu\Desktop\Giao diện app\Giao dien ban hang\ANHSP\nuoc-ngot-co-gas-7up-pet-15l-2-org.JPG',SINGLE_BLOB)AS IMG

INSERT INTO [dbo].[Sanpham]
	(
	    IDhang,
    TenHang,
    Loaihang,
    SoLuongtrongkho,
    Sothutu,
    TheTich,
    PICTURE
	)
	SELECT 'H0005',N'Nước COCA', N'2',30,'3', N'450g',* FROM OPENROWSET(BULK N'C:\Users\Thuong Nhu\Desktop\Giao diện app\Giao dien ban hang\ANHSP\nuoc-ngot-coke-sleek-330ml-1-700x467-1.JPG',SINGLE_BLOB)AS IMG

INSERT INTO [dbo].[Sanpham]
	(
	    IDhang,
    TenHang,
    Loaihang,
    SoLuongtrongkho,
    Sothutu,
    TheTich,
    PICTURE
	)SELECT 'H0006',N'Nước COCA', N'2',15,'2', N'450ml',* FROM OPENROWSET(BULK N'C:\Users\Thuong Nhu\Desktop\Giao diện app\Giao dien ban hang\ANHSP\z7M-zs3QaPiv.JPG',SINGLE_BLOB)AS IMG

INSERT INTO [dbo].[Sanpham]
	(
	    IDhang,
    TenHang,
    Loaihang,
    SoLuongtrongkho,
    Sothutu,
    TheTich,
    PICTURE
	)SELECT 'H0007',N'Mỳ Cay Hàn Quốc', N'3',20,'1', N'300g',* FROM OPENROWSET(BULK N'C:\Users\Thuong Nhu\Desktop\Giao diện app\Giao dien ban hang\ANHSP\muamy-mi-han-quoc-nongshim-mi-goi-shin-ramyun-noodle-soup-1.JPG',SINGLE_BLOB)AS IMG

INSERT INTO [dbo].[Sanpham]
	(
	    IDhang,
    TenHang,
    Loaihang,
    SoLuongtrongkho,
    Sothutu,
    TheTich,
    PICTURE
	)SELECT 'H0008',N'Mỳ Cay', N'3',0,'1', N'200g',* FROM OPENROWSET(BULK N'C:\Users\Thuong Nhu\Desktop\Giao diện app\Giao dien ban hang\ANHSP\Dich-vu-cong-bo-chat-luong-mi-goi-mi-an-lien-everestlaw.JPG',SINGLE_BLOB)AS IMG

INSERT INTO [dbo].[Sanpham]
	(
	    IDhang,
    TenHang,
    Loaihang,
    SoLuongtrongkho,
    Sothutu,
    TheTich,
    PICTURE
	)SELECT 'H0009',N'Nước Hoa', N'3',15,'4', N'450ml',* FROM OPENROWSET(BULK N'C:\Users\Thuong Nhu\Desktop\Giao diện app\Giao dien ban hang\ANHSP\images (1).JPG',SINGLE_BLOB)AS IMG

INSERT INTO [dbo].[Sanpham]
	(
	    IDhang,
    TenHang,
    Loaihang,
    SoLuongtrongkho,
    Sothutu,
    TheTich,
    PICTURE
	)SELECT 'H0010',N'kem Dưỡng Trắng da', N'3',20,'3', N'350g',* FROM OPENROWSET(BULK N'C:\Users\Thuong Nhu\Desktop\Giao diện app\Giao dien ban hang\ANHSP\nuoc-ngot-co-gas-7up-pet-15l-2-org.JPG',SINGLE_BLOB)AS IMG

INSERT INTO [dbo].[Sanpham]
	(
	    IDhang,
    TenHang,
    Loaihang,
    SoLuongtrongkho,
    Sothutu,
    TheTich,
    PICTURE
	)SELECT 'H0011',N'Nước STING', N'2',5,'3', N'450ml',* FROM OPENROWSET(BULK N'C:\Users\Thuong Nhu\Desktop\Giao diện app\Giao dien ban hang\ANHSP\nuoc-ngot-coke-sleek-330ml-1-700x467-1.JPG',SINGLE_BLOB)AS IMG

INSERT INTO [dbo].[Sanpham]
	(
	    IDhang,
    TenHang,
    Loaihang,
    SoLuongtrongkho,
    Sothutu,
    TheTich,
    PICTURE
	)SELECT 'H0012',N'Snack Bắp', N'4',15,'3', N'250g',* FROM OPENROWSET(BULK N'C:\Users\Thuong Nhu\Desktop\Giao diện app\Giao dien ban hang\ANHSP\images (3).JPG',SINGLE_BLOB)AS IMG

INSERT INTO [dbo].[Sanpham]
	(
	    IDhang,
    TenHang,
    Loaihang,
    SoLuongtrongkho,
    Sothutu,
    TheTich,
    PICTURE
	)SELECT 'H0013',N'Snack Bí Đỏ', N'4',15,'3', N'250g',* FROM OPENROWSET(BULK N'C:\Users\Thuong Nhu\Desktop\Giao diện app\Giao dien ban hang\ANHSP\images (2).JPG',SINGLE_BLOB)AS IMG

INSERT INTO [dbo].[Sanpham]
	(
	    IDhang,
    TenHang,
    Loaihang,
    SoLuongtrongkho,
    Sothutu,
    TheTich,
    PICTURE
	)SELECT 'H0014',N'Snack Hành', N'4',15,'3', N'250g',* FROM OPENROWSET(BULK N'C:\Users\Thuong Nhu\Desktop\Giao diện app\Giao dien ban hang\ANHSP\images (4).JPG',SINGLE_BLOB)AS IMG
ALTER TABLE dbo.HangHoaTheonam DROP PK__HangHoaT__E653D28D37BEC353 
ALTER TABLE dbo.HangHoaTheonam ALTER COLUMN Nam DATETIME

INSERT INTO dbo.HangHoaTheonam
(
    IDhang,
    Nam,
    Gia
)
VALUES
(   'H0001',  GETDATE(), 50 ),(   'H0002',  GETDATE(), 45 ),(   'H0003',  GETDATE(), 10  ),(   'H0004',  GETDATE(), 18 ),(   'H0005',  GETDATE(), 10 ),
(   'H0006',  GETDATE(), 18 ),(   'H0007',  GETDATE(), 20 ),(   'H0008',  GETDATE(), 5 ),(   'H0009',  GETDATE(), 150 ),(   'H0010',  GETDATE(), 200 ),
(   'H0011',  GETDATE(), 8 ),(   'H0012',  GETDATE(), 5 ),(   'H0013',  GETDATE(), 6 ),(   'H0014' , GETDATE(), 6 )


SELECT   Sanpham.IDhang,TenHang ,Nam,Gia FROM  dbo.Sanpham INNER JOIN dbo.HangHoaTheonam ON HangHoaTheonam.IDhang = Sanpham.IDhang
DELETE FROM dbo.CTPN
INSERT INTO dbo.CTPN
(
    IDhang,
    IDphieunhap,
    SoLuong,
    Gia
)
VALUES
(   'H0002', 'PN001',5,45 ),(   'H0005', 'PN001',2,10 ),(   'H0001', 'PN002',10,50 ),(   'H0013', 'PN002',10,6 )

INSERT INTO dbo.PhieuXuat
(
    IDphieuxuat,
    Tongtien,
    IDKhach
)
VALUES
(   'PX001',  NULL, 'K0001' ),(   'PX002',  NULL, 'K0002' ),(   'PX003',  NULL, 'K0001' ),(   'PX004',  NULL, 'K0004' )

INSERT INTO dbo.CTPX
(
    IDhang,
    IDphieuxuat,
    SoLuong,
    Gia,
    Donvitinh,
    Tenhang
)
VALUES
(   'H0001',   'PX001',   3, 50,  N'Thùng',  N'Bánh London'  ),(   'H0002',   'PX001',   1, 50,  N'Thùng',  N'Bánh London2'  ),
(   'H0004',   'PX002',   3, 50,  N'Thùng',  N'Nước SPRIZE'  ) 


SELECT * FROM dbo.Sanpham