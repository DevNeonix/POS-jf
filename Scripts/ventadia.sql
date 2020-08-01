
select operacion as sale_id, 'T'+ltrim(rtrim(tienda)) as store_id, 'V'+ltrim(rtrim(vendedor)) as seller_id,
convert(smalldatetime,fecdoc,113) as date_sale, 
(select sum(sale) from  movimdet as m2 where m1.operacion= m2.operacion) as qty,
m1.total as total_amount, 'p101' as mall_id, '101' as store_number
from movimcab as m1
where  fecdoc   between '16-04-2019' and '16-04-2019'+' 23:59:59.999'
and tienda = '08' and coddoc in ('BL', 'FT')
UNION
select operacion as sale_id, 'T'+ltrim(rtrim(tienda)) as store_id, 'V'+ltrim(rtrim(vendedor)) as seller_id,
convert(smalldatetime,fecdoc,113) as date_sale, 
(select sum(ENTRA) from  movimdet as m2 where m1.operacion= m2.operacion)*-1 as qty,
m1.total*-1 as total_amount, 'p101' as mall_id, '101' as store_number
from movimcab as m1
where  fecdoc   between '16-04-2019' and '16-04-2019'+' 23:59:59.999'
and tienda = '08' and coddoc = 'NC'

--VENDEDORES
select USUARIO as seller_id, nombres as seller_name, 'P101' AS mall_id, '101' AS store_number from usuarios




--SELECT * FROM MOVIMDET WHERE OPERACION ='0000033891'

--exec master..xp_cmdshell 'bcp "SELECT cod_personal,  fotocheck, ide FROM planillas.DBO.chichoobr" queryout "C:\temp\chicho2.csv" -q';' -c -T'