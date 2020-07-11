ALTER PROC USP_NhapHang 
@IDhang CHAR(5), @IDphieunhap CHAR(5),
@Ngaylap CHAR(20),@thanhtien MONEY,@IDnhacc CHAR (5),@soluong int ,@Gia money ,
@tongtien MONEY 

AS BEGIN
		IF NOT  EXISTS(SELECT IDphieunhap  FROM dbo.PhieuNhap WHERE IDphieunhap=@IDphieunhap )
			BEGIN
			  /* SET  @IDphieunhap= (SELECT Max( IDphieunhap ) FROM dbo.PhieuNhap) 
     WHILE @IDphieunhap>=0 AND @IDphieunhap<9
	 SET @IDphieunhap= 'PN00'+CONVERT(CHAR, CONVERT(INT , @IDphieunhap) + 1) 
	 WHILE @IDphieunhap>9
	 SET @IDphieunhap= 'PN0'+CONVERT(CHAR, CONVERT(INT , @IDphieunhap) + 1) */   
	   
      INSERT INTO  dbo.PhieuNhap
      (
          IDphieunhap,
          IDnhacc,
          Ngaylap,
          Tongtien
      )
      VALUES
      (   @IDphieunhap,        -- IDphieunhap - char(5)
          @IDnhacc,        -- IDnhacc - char(5)
          @Ngaylap, -- Ngaylap - date
          @tongtien      -- Tongtien - money
          ) 
		/*IF  Exists(SELECT  IDhang FROM  dbo.CTPN WHERE  IDhang=@IDhang)
        BEGIN 
           PRINT N'Hàng Đã Tồn Tại trên PN đó '
		  
           ROLLBACK TRANSACTION 
        END */ 
		
      
           INSERT INTO  dbo.CTPN
           (
               IDhang,
               IDphieunhap,
               SoLuong,
               Gia,
               ThanhTien
           )
           VALUES
           (   @IDhang,   -- IDhang - char(5)
                @IDphieunhap,   -- IDphieunhap - char(5)
              @soluong, -- SoLuong - numeric(18, 0)
               @Gia, -- Gia - money
              @thanhtien  -- ThanhTien - money
               )
			   SET @thanhtien=(SELECT  SoLuong*Gia FROM dbo.CTPN)

				
		  END 
		  IF EXISTS (SELECT  IDphieunhap FROM dbo.CTPN) 

		  
           INSERT INTO  dbo.CTPN
           (
               IDhang,
               IDphieunhap,
               SoLuong,
               Gia,
               ThanhTien
           )
           VALUES
           (   @IDhang,   -- IDhang - char(5)
                @IDphieunhap,   -- IDphieunhap - char(5)
              @soluong, -- SoLuong - numeric(18, 0)
               @Gia, -- Gia - money
              @thanhtien  -- ThanhTien - money
               )


END 
EXECUTE USP_NhapHang  'H0010','PN009',NULL,NULL,'NCC02',6,NULL,NULL
DELETE  dbo.CTPN WHERE IDhang='H0009'

UPDATE dbo.CTPN SET ThanhTien=(SoLuong*Gia)*0.9 WHERE IDhang=''AND IDphieunhap='';

SELECT * FROM  dbo.KhachHang
SELECT * FROM dbo.PhieuNhap
SELECT IDhang,CTPN.IDphieunhap,SoLuong,Gia,ThanhTien FROM dbo.CTPN,dbo.PhieuNhap


SELECT * FROM  dbo.PhieuXuat INNER JOIN dbo.CTPX ON CTPX.IDphieuxuat = PhieuXuat.IDphieuxuat 



CREATE PROC USp_xuathang 
@MaPhieuxuat CHAR(5),@mahang CHAR(5),@soluong INT , @gia MONEY,@dvt NVARCHAR(20),@tenhang NVARCHAR(50),@togtien MONEY, @Makhach CHAR(5)
AS BEGIN
       IF  NOT EXISTS(SELECT IDphieuxuat FROM dbo.PhieuXuat WHERE IDphieuxuat= @MaPhieuxuat) 
		BEGIN
					   INSERT INTO dbo.PhieuXuat
				(
					IDphieuxuat,
					Tongtien,
					IDKhach
				)
				VALUES
				(   @MaPhieuxuat,  NULL, @Makhach )

				IF  NOT EXISTS(SELECT  IDhang,IDphieuxuat FROM dbo.CTPX WHERE IDhang=@mahang OR   IDphieuxuat= @MaPhieuxuat )
					BEGIN
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
						(  @mahang,  @MaPhieuxuat ,   @soluong, @gia, @dvt,  @tenhang  )
					END 
					ELSE BEGIN
							UPDATE dbo.CTPX SET IDhang= @mahang, IDphieuxuat=@MaPhieuxuat,SoLuong=@soluong,Gia=@gia,Donvitinh=@dvt ,Tenhang= @tenhang  WHERE IDhang=@mahang AND  IDphieuxuat= @MaPhieuxuat
				
					END 
				END 

			
   END
   EXECUTE USp_xuathang 'PX005','H0006',3 ,18 ,N'Thùng',N'Nước COCA', NULL,'K0004'



CREATE PROC usp_chonsp
AS BEGIN
       SELECT*FROM dbo.Sanpham
   END
   EXECUTE  usp_chonsp
  
  CREATE PROC USP_noiSanpham
   AS BEGIN
          SELECT CTPX.IDhang ,CTPX.Tenhang,SoLuongtrongkho,TheTich,PICTURE,CTPX.IDphieuxuat,Donvitinh,HangHoaTheonam.Gia,Tongtien FROM dbo.Sanpham INNER JOIN dbo.CTPX ON CTPX.IDhang = Sanpham.IDhang JOIN dbo.PhieuXuat ON PhieuXuat.IDphieuxuat = CTPX.IDphieuxuat JOIN dbo.HangHoaTheonam ON HangHoaTheonam.IDhang = Sanpham.IDhang
      END
	  EXECUTE USP_noiSanpham
	 ALTER TABLE   dbo.HoaDon ADD Tinhtrang NVARCHAR(50)
	 SELECT*FROM dbo.PhieuXuat 