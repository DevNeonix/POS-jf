-- =====================================================
use VENTAS
-- =====================================================
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =====================================================
-- Author:		MABEL MOLINA
-- Create date: 10-MAU-2013
-- Description:	DEVUELVE CANTIDADES MOVIDAS EN UN
--				RANGO DE FECHAS, CODIGO Y/O TIENDA
-- =====================================================
alter PROCEDURE SP_ENTRA_SALE_FECHA 
	-- Add the parameters for the stored procedure here
	@INI	char(10),
	@FIN	char(10),
	@TDA	char(2),
	@COD	char(25),
	@TIP	CHAR (2)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
	/*	TIP --> PP = POR PRODUCTO	/	PG = POR GRUPO		*/	

DECLARE @strsql    nvarchar(4000)

SET @strsql= 'SET DATEFORMAT DMY; SELECT DISTINCT '
IF ltrim(rtrim(@TIP))<>'PG' 	
	SET @strsql= @strsql + ' CODART AS CODIGO,  '
ELSE 
	SET @strsql= @strsql + 'GRUPO, '
SET @strsql= @strsql + ' descri, sum(entra) as ingresos, sum(sale) as salidas	FROM VIEW_VENTAS_ARTICULO '
SET @strsql= @strsql + ' WHERE  FECHA between '''+@INI+''''
SET @strsql= @strsql + ' AND DateAdd(day,1,'''+@FIN+''')'
SET @strsql= @strsql + ' and tipdoc in (''BL'',''FC'',''NC'') '
IF ltrim(rtrim(@TDA))<> 'TT' 
	SET @strsql=@strsql+ ' and TIENDA ='''+ltrim(rtrim(@TDA))+''''
	
IF ltrim(rtrim(@TIP))<> 'PG' 
	begin
		SET @strsql=@strsql+ ' and CODART ='''+ltrim(rtrim(@cod))+''''
		SET @strsql=@strsql+ ' group by codart, descri ORDER BY 1'
	end	
ELSE
	begin
		SET @strsql=@strsql+ ' and GRUPO ='''+ltrim(rtrim(@cod))+''''
		SET @strsql=@strsql+ ' group by grupo, descri ORDER BY 1'
	end
	


exec sp_executesql @strsql
END
GO
