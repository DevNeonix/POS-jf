


select USUARIO as seller_id, nombres as seller_name, 'P101' AS mall_id, '101' AS store_number INTO VEN from JACINTA.dbo.usuarios
go

exec master..xp_cmdshell 'bcp "SELECT seller_id, seller_name, mall_id,store_number FROM jacinta.dbo.VEN" queryout "C:\temp\VENDEDORES-1_P101_101_20190815.csv" -c -t ";" -S  -T'
go
DROP TABLE VEN