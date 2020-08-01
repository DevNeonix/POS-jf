
ALTER procedure SP_COMPARA
AS

DECLARE @columns varchar(MAX);
DECLARE @sql nvarchar(max)
CREATE TABLE #Table
(
 TIENDA char(2), 
ANO char(4), 
VENTA numeric(18,7),
MES char(2) 
);

INSERT INTO #Table
SELECT tienda, ano, venta, mes from historia

 SET @columns = STUFF( ( SELECT   ',' + QUOTENAME(LTRIM(mes))
 FROM
   (SELECT DISTINCT mes     FROM #Table   ) AS T ORDER BY mes FOR XML PATH('') ), 1, 1, '');

 SET @sql = N'
 SELECT   *   FROM  (    SELECT  TIENDA, ANO AS [AÑO], VENTA, MES
  FROM #Table  ) AS T
  PIVOT     (  SUM(Venta)  FOR mes IN (' + @columns + N')
  ) AS P order by  tienda, ano desc;'; 

EXEC sp_executesql @sql;
--DROP TABLE #Table;