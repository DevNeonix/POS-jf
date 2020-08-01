select * from movimcab where  operacion ='0000000001'
union
select * from movimcab where  operacion ='0000005651' order by 1

insert into movimcab (operacion, tienda, coddoc, tipmov, serie, numdoc,
fecdoc, docori, numori,usuario, fecha, estado)
values ('0000005651', '09', 'NI','E','000','EXCEL', '2013-01-03', 'TR', '0000000000', 'SIST', '2013-01-03','A')
select * from movimdet where  operacion ='0000000001'
UNION
select * from movimdet where  operacion ='0000005651' order by 1
SELECT * INTO BKP_MOVIMDET FROM MOVIMDET
UPDATE TBL_CHACA SET TIENDA = '09' , OPERACION = '0000005651', ITEM = '01'
SELECT * FROM TBL_CHACA ORDER BY 1
INSERT INTO MOVIMDET 

SELECT OPERACION,TT.TIENDA, ITEM, TT.CODIGO,STK,0,0,0,0,0,NULL,''
FROM TBL_CHACA TT  INNER JOIN ARTICULOS AA ON TT.CODIGO = AA.CODIGO  
WHERE NOT TT.CODIGO IS NULL AND AA.TIENDA = '09'