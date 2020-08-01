<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
<!-- eliminacion de duplicados-->
<!-- update movimcab set estado='E' where operacion ='0000000550' -->
<!-- update movimdet set entra=0,sale=0,vale=0 where operacion='0000000550' -->
<!-- update caja set estado='E' where operacion='0000000550' -->
<!-- update articulos set stock=(select sum(m.entra)-sum(m.sale) as stock from movimdet m 
		where m.tienda='07' and articulos.codigo=m.codart group by tienda,codart) 
        where articulos.estado='A' and articulos.tienda='07' and not codigo like 'RAA%'-->
<!-- update articulos set stock=(select sum(m.entra)-sum(m.sale) as stock from movimdet m where m.tienda = '10' and articulos.codigo=m.codart group by tienda,codart) where articulos.tienda='10'-->
<p> Pronto </p>
</body>

</html>
