
select USUARIO as seller_id, nombres as seller_name, 'P101' AS mall_id, '101' AS store_number INTO #VEN from JACINTA.dbo.usuarios
go

DECLARE @OutputFilePath nvarchar(max); 
SET @OutputFilePath = 'C:\temp'

DECLARE @ExportSQL nvarchar(max); 
SET @ExportSQL = N'EXEC master.dbo.xp_cmdshell ''bcp "SELECT * FROM jacinta.dbo.#ven ORDER BY RowNumber" queryout "' + @OutputFilePath + '\OutputData.csv" -T -c -t -S WIN-SIITTJOB7OV'''
EXEC(@ExportSQL)

drop table #ven