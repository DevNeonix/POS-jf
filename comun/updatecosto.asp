<%@LANGUAGE="VBSCRIPT"  CODEPAGE="65001"%>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<%Usuario = Request.Cookies("tienda")("usr")%>
<%tda     = Request.Cookies("tienda")("pos")%>
<!--#include file="../includes/Cnn.inc"-->
<!--#include file="../comun/funcionescomunes.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<body>
<% 
'*********************** OJO **********************
'REVISAR EL LARGO DE LAS COLUMNAS de la tabla
'*********************** OJO **********************
' codigo / grupo
cod = Request.QueryString("cod")
pre = Request.QueryString("pre")
lin = Request.QueryString("lin")
cla = Request.QueryString("cla")


'**************************************************************************************
' ver si COD tiene 5 o 10 caracteres para determinar si es por artículo o por grupo
' ver si TDA dice TODAS... o es por una tienda específica
'**************************************************************************************


CAD =   " UPDATE costos SET          " & _
        " costo = "&PRE&",           " & _
        " linea = '"&lin&"',         " & _ 
        " clase = '"&cla&"'          " & _
        " WHERE CODIGO = '"&COD&"' ; "
' Inicia transacción , para que los datos no queden a medias
RESPONSE.WRITE("<br>")
RESPONSE.WRITE(CAD)

Cnn.BeginTrans	
Cnn.Execute(CAD)
if  err.number <> 0 then
	Response.Write ("No se han podido actualizar los datos soliciatados,  Reintente en unos minutos")
	Cnn.RollbackTrans
	Cnn.Abort
else
	Cnn.CommitTrans
end if

precio = formatnumber(pre,2,,,true)
response.write("<br>")
response.write(precio)
Cnn.Close	
set Cnn = Nothing
SET RS = Nothing	
'regresa a la página de donde fué llamado, para que vea que agregó el registro
%>

<script type="text/jscript" language="jscript">
    op = parseInt("<%=Request.QueryString("opc") %>",10)
    precio = "<%=formatnumber(pre,2,,true)%>"   
    eval("parent.window.document.all.LS"+op+".value = precio") 
</script>

</BODY>
</HTML>

