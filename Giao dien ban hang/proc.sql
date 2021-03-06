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
EXECUTE USP_NhapHang  'H0009','PN008',NULL,NULL,'NCC01',6,45,NULL
SELECT * FROM  dbo.CTPX
SELECT * FROM dbo.PhieuNhap