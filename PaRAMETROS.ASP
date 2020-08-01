<%@LANGUAGE="VBSCRIPT"  CODEPAGE="65001"%>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<%	tienda = Request.Cookies("tienda")("pos") 	%>
<%	USUARIO = Request.Cookies("tienda")("USR") 	%>



<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<!--#include file="includes/Cnn.inc"-->
<!--#include file="COMUN/FUNCIONESCOMUNES.ASP"-->
<!--#include file="COMUN/COMUNtbl.ASP"-->
<body>


<meta http-equiv="Content-Type" content="text/html; charset=utf-8" lang="es-pe" />
<title>PARAMETROS</title>
</head>

<% CAD = "EXEC SP_TIPODECAMBIO"
RS.Open CAD, Cnn%>
<%=rs.recordcount %>
<table align="center" cellspacing="5">
    <tr><td height="30" valign="middle" colspan="3" align="center" class="Estilo6">TIPO DE CAMBIO DEL DIA</td></tr>
	<tr><td  align="center" class="Estilo9">Venta</td>
    <td  align="center" class="Estilo7">&nbsp;</td>
    <td  align="center" class="Estilo9">Compra</td>
</tr>
<tr><td  align="center" class="Estilo1"><%=RS("ALTO") %></td>
    <td  align="center" class="Estilo7">&nbsp;</td>
    <td  align="center" class="Estilo1"><%=RS("bajo") %></td>
</tr>
</table>
<%rs.close 
rs.open "select igv from parametros", cnn 
rs.movefirst%>
<table align="center" cellspacing="5">
    <tr><td height="30" valign="middle" colspan="3" align="center" class="Estilo6">I.G.V.</td></tr>

<tr>
    <td  align="center" class="Estilo1"><%=RS("IGV") %> &nbsp;%</td>
</tr>
</table>
</body>
</html>
