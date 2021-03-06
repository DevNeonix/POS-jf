USE [VENTAS]
GO
/****** Object:  StoredProcedure [dbo].[SP_CAJA_DIA_todas]    Script Date: 12/01/2017 11:36:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- ===============================================================
-- Author:		MABEL MOLINA
-- Create date: 25-03-2016
-- Description:	DEVUELVE LOS DATOS PARA CUADRE DE CAJA DIARIO
-- ===============================================================
-- FAC 20130121 dias atras adicione una segunda fecha
-- donde utilice la f(x) DateAdd, adiciono inner con la cabecera
-- cambio c1.fecha por c2.fecdoc a fecha
-- dar prefencia a la fecha documento
-- FAC 20130305 Se adiciona c3.cliente, c3nombre inner clientes
-- MMB - para aque acumule todas las tiendas
ALTER PROCEDURE [dbo].[SP_CAJA_DIA_todas] 
	-- Add the parameters for the stored procedure here
	@fec1 VARCHAR(10),
	@fec2 varchar(10)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
	set dateformat DMY;
    -- Insert statements for procedure here
	select c1.operacion, c2.fecdoc as fecha, c1.moneda,c1.tipo, SOL= case c1.moneda  when 'US' THEN (c1.monto*c1.TCAMBIO)
		ELSE c1.MONTO END, c1.monto AS DOLAR, c2.coddoc, c2.serie+'-'+c2.numdoc as numdoc, c2.total, c3.Cliente,
		 c3.nombre, c1.nota, convert(char(5),c2.fecdoc, 108) [hora]
	from caja c1
	inner join movimcab c2 on c1.operacion= c2.operacion
	inner join clientes c3 on c2.cliente= c3.cliente
	where c2.fecdoc between @fec1 and @fec2+' 23:59:59.999'  and c1.estado='A'
	order by  c2.serie+'-'+c2.numdoc
END
