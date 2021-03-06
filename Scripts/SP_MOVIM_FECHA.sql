USE [VENTAS]
GO
/****** Object:  StoredProcedure [dbo].[SP_ENTRA_SALE_FECHA]    Script Date: 05/17/2013 12:09:42 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =====================================================
-- Author:		MABEL MOLINA
-- Create date: 20-MAY-2013
-- Description:	DEVUELVE CABECERAS DE MOVIMEINTOS EN 
--				UN RANGO DE FECHAS, CODIGO Y/O TIENDA
-- =====================================================
alter PROCEDURE SP_MOVIM_FECHA 
	-- Add the parameters for the stored procedure here
	@INI	char(10),
	@FIN	char(10),
	@TDA	char(2),
	@DOC	nvarchar(200)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
	/*	DOC --> TIPO(S) DE DOCUMENTOS SOLICITADOS		*/	

DECLARE @strsql    nvarchar(4000)

SET @strsql= 'SET DATEFORMAT DMY; SELECT DISTINCT '
SET @strsql= @strsql + ' TIENDA AS TDA, OPERACION AS OPE, CODDOC,  SERIE, NUMDOC, DOCORI, numori, FECDOC, '
SET @strsql= @strsql + ' PVP, DESCUENTO AS DCT, SUBTOT AS STOT, IGV, TOTAL AS TOT, M1.CLIENTE, NOMBRE '
SET @strsql= @strsql + ' FROM MOVIMCAB AS M1 LEFT OUTER JOIN CLIENTES AS C1 ON M1.CLIENTE = C1.CLIENTE '
SET @strsql= @strsql + ' WHERE  FECHA between '''+@INI+''''
SET @strsql= @strsql + ' AND DateAdd(day,1,'''+@FIN+''')'

IF ltrim(rtrim(@TDA))<> 'TT' 
	SET @strsql=@strsql+ ' and TIENDA ='''+ltrim(rtrim(@TDA))+''''
	
IF ltrim(rtrim(@DOC))<> '' 
	begin
		SET @strsql= @strsql + ' and CODDOC in ('+@DOC+ ') '
	end	
		
SET @strsql=@strsql+ ' ORDER BY TIENDA, OPERACION'
	


exec sp_executesql @strsql
END
