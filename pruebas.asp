<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<% tienda = Request.Cookies("tienda")("pos") %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html>

<head>
<!--#include file="COMUN/FUNCIONESCOMUNES.ASP"-->
<!--#include file="includes/Cnn.inc"-->
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>
<body onload="document.all.soles.focus()">
<table border="0" align="left" cellpadding="0" cellspacing="10" width="100%">
    <tr>
        <td><input id="soles" name="soles" type="text" class="Estilo2"/></td>
        <td><input type="button" onclick="manda()" id="mm" name="mm" value="OK" /></td>
    </tr>
    <tr>
        <td colspan="3" align="center" ><input type ="text" id="sonn" name="sonn" class="Estilo0" /></td>
    </tr>
</table>
       

<script language="jscript" type="text/jscript">

    function manda() { document.getElementById('sonn').value = 'Son : ' + FComson(document.all.soles.value) }

</script> 

</body>
</html>
