-----------------------------------------------------------------------------------
USE JACINTA
GO

CREATE TABLE	[dbo].[HISTORIA](
	[ANO]		[char](4)  ,
	[MES]		[CHAR](2)  ,
	[TIENDA]	[char](2)  ,
	[VENTA]		[NUMERIC] (18,2),
PRIMARY KEY (ANO, MES, TIENDA)
) ON [PRIMARY]

	
-- INSERTA LOS VALORES HISTÓRICOS DE VENTAS DE LAS TIENDAS DESDE EL INICIO DE LOS TIEMPOS
/* delete historia
INSERT INTO HISTORIA
select YEAR(FECDOC) AS ANO, RIGHT('00'+  CONVERT(varchar(2), MONTH(FECDOC)),2) AS MES, TIENDAs.VENDE ,TOT =  sum(monto) 
from VENTAS..caja 
inner join VENTAS..movimcab c2 on caja.operacion= c2.operacion
inner join tiendas on caja.tienda collate Modern_Spanish_CI_AS = tiendas.codigo
where caja.estado='A' and tiendas.estado = 'A'
GROUP BY CAJA.TIENDA, YEAR(FECDOC), MONTH(FECDOC), TIENDAS.VENDE
ORDER BY  YEAR(FECDOC), MONTH(FECDOC) , CAJA.TIENDA 
*/
select * from tiendas	
SELECT * FROM HISTORIA