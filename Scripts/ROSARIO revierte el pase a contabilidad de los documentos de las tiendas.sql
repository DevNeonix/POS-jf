-- Borra el temporal del real
delete RSFACCAR..CF0012FACT where CF_CNUMSER = '114'

-- elimina la marca de pase a Contabilidad en MOVIMCAB
update movimcab
set CONTA = '' where TIENDA = '10' and YEAR(fecdoc) = 2017
and MONTH(fecdoc) = 8 and (DAY(fecdoc)> 19 and DAY(fecdoc) <= 31)