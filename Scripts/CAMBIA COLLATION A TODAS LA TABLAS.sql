/**********************************************************************
PEDIR SALIDA A TEXTO Y COPIAR EN OTRO QUERY LAS SENTENCIAS A EJECUTAR
***********************************************************************/
SELECT 'ALTER TABLE ' + o.name + ' ALTER COLUMN ' + C.name + ' ' + 
 	T.name +'(' + CAST(c.length as varchar(4)) + ') COLLATE ' + 
 ' SQL_Latin1_General_CP1_CI_AS;'
 from syscolumns C, sysobjects O, systypes T 
 where (C.collation is not null) and C.id = O.id and c.TYPE = T.TYPE AND
 C.XUSERTYPE = T.XUSERTYPE AND o.type = 'U'  
 ORDER BY O.NAME;