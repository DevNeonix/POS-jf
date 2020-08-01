<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<%Usuario = Request.Cookies("tienda")("usr")%>
<%tda     = Request.Cookies("tienda")("pos")%>
<%descri  = Request.Cookies("tienda")("tda")%>
<%Response.Buffer = TRUE %>
<!--#include file="../includes/Cnn.inc"-->
<!--#include file="../comun/funcionescomunes.asp"-->
<!--#include file="../includes/funcionesVBscript.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<body>
<% TEM= REQUEST.QueryString("TEM")
correos = "csaba@elmodelador.com.pe"
cAD =   " exec sp_reponer   '"&TDA&"', '"&descri&"' , '"&correos&"', '"&tem&"' "
' Inicia transacción , para que los datos no queden a medias
'RESPONSE.WRITE("<br>")
RESPONSE.WRITE(CAD)
'RESPONSE.END
Cnn.BeginTrans	
Cnn.Execute(CAD)
if  err.number <> 0 then
	Response.Write ("No se han podido actualizar los datos soliciatados,  Reintente en unos minutos")
	Cnn.RollbackTrans
	Cnn.Abort
else
	Cnn.CommitTrans	
end if

Cnn.Close	
set Cnn = Nothing
SET RS = Nothing	
'regresa a la página de donde fué llamado, para que vea que agregó el registro
%>
<script language="javascript" type="text/jscript">

cad = '../reportes/reposiciones.ASP'
window.location.replace = cad
</script>
</BODY>
</HTML>
