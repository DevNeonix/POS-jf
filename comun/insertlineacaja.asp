<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = true%>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos")%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<link REL="stylesheet" TYPE="text/css" HREF="../ventas.CSS" >
<!--#include file="funcionescomunes.asp"-->
<!--#include file="../includes/funcionesVBscript.asp"-->
<!--#include file="../includes/cnn.inc"-->

<%
	operacion = request.QueryString("ope")
	tienda = request.QueryString("tda")
    mon = request.QueryString("mon")
    tipo = request.QueryString("tipo")
    monto = request.QueryString("monto")
    obs = request.QueryString("obs")

cad =   "insert into jacinta..caja(tienda,operacion,lin,MONEDA,TIPO,MONTO,TCAMBIO,USUARIO,FECHA,ESTADO,NOTA) SELECT top(1) '"+tienda+"','"+operacion+"',LIN+1,'"+mon+"','"+tipo+"',"+monto+",TCAMBIO,'csaba',getdate(),'A','"+obs+"' FROM jacinta..caja WHERE tienda = '"+tienda+"'  AND OPERACION = '"+operacion+"' order by lin desc"
response.write(cad)
CNN.EXECUTE CAD    
    
%>

<body></body>
<script language="jscript" type="text/jscript">
    alert("<%=cad%>")
	window.close();
</script>
</html>
