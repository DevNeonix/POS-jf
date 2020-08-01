<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%session.LCID = 2057 %>
<%Usuario = Request.Cookies("tienda")("usr")%>
<%tienda  = Request.Cookies("tienda")("pos")%>
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

cod = Request.QueryString("cod")
lin = Request.QueryString("lin")
pla = Request.QueryString("pla")

 CAD = "UPDATE ARTICULOS SET planilla = "&pla&" WHERE CODIGO = '"&COD&"'; "

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

Cnn.Close	
set Cnn = Nothing
SET RS = Nothing	
'regresa a la página de donde fué llamado, para que vea que agregó el registro
%>

<script type="text/jscript" language="jscript">
    lin = parseInt("<%=Request.QueryString("lin") %>",10)
    pla = parseInt("<%=Request.QueryString("pla") %>",10)

    //alert("pla")
    eval("parent.window.document.all.pl"+lin+".selectedIndex = pla") 
</script>

</BODY>
</HTML>

