-- ================================================
-- Template generated from Template Explorer using:
-- Create Procedure (New Menu).SQL
--
-- Use the Specify Values for Template Parameters 
-- command (Ctrl-Shift-M) to fill in the parameter 
-- values below.
--
-- This block of comments will not be included in
-- the definition of the procedure.
-- ================================================
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- ========================================================
-- Author:		MABEL MOLINA
-- Create date: 18-07-2012
-- Description:	LLENA EL MAESTRO DE ARTICULOS POR TIENDA
-- PASOS:
-- 1.- COLOCA ESTADO ELIMINADO A TODOS LOS ARTICULOS QUE 
--     NO EXISTEN EN EL REAL COMO VIGENTES
-- 2.- INSERTA TODOS LOS ARTICULOS NUEVOS DEL REAL
-- 3.- 
-- ========================================================
ALTER PROCEDURE DBO.sp_LLENA_TIENDA2
	-- Add the parameters for the stored procedure here
	@TDA CHAR(2)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	-- INSERTA TODOS LOS ARTICULOS NUEVOS DEL REAL
	INSERT INTO articulos (TIENDA, CODIGO, STOCK, MINIMO, LISTA1, LISTA2, LISTA3, PLANILLA, USUARIO, FECHA, ESTADO)
	(SELECT DISTINCT @tda AS TIENDA, A2.AR_CCODIGO, 0 AS STOCK, 2 AS MINIMO, 
		A2.AR_NPRECI1 AS LISTA1, A2.AR_NPRECI2 AS LISTA2, A2.AR_NPRECI3 AS LISTA3, 0 AS PLANILLA,
	   'SISTEMAS' AS USUARIO, GETDATE() AS fecha, 
                      'A' AS estado
FROM         RSFACCAR.dbo.AL0012ARTI AS A2 LEFT OUTER JOIN
                      dbo.ARTICULOS AS AAA ON AAA.CODIGO = A2.AR_CCODIGO COLLATE Modern_Spanish_CI_AI
WHERE     (A2.AR_CESTADO = 'v') AND (A2.AR_CFSTOCK = 's') AND (A2.AR_CFAMILI IN ('JF', 'MI')) 
AND ((@TDA + A2.AR_CCODIGO COLLATE Modern_Spanish_CI_AI) NOT IN (SELECT     xx.TIENDA + xx.CODIGO AS codig
                            FROM dbo.ARTICULOS as xx))
	)

END
GO
