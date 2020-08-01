<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<% tienda = Request.Cookies("tienda")("pos") %>
<% BANCO = Request.Cookies("tienda")("bcp") %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK REL="Shortcut Icon" HREF="imagenes/jf_icon.ico"> 
<title>Jacinta Fernandez - POS</title>

<!--#include file="includes/Cnn.inc"-->
<!--#include file="comun/funcionescomunes.asp"-->
<script type="text/jscript" language="jscript">

function calcHeight() {

//find the height of the internal page
var the_height =  document.getElementById('body0').contentWindow.       document.body.scrollHeight;

    //change the height of the iframe
    document.getElementById('body0').height = the_height + 20;

}
</script>

</head>
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<body align="center">

<center>
<table align="center" border="0" cellpadding="0" cellspacing="5" width="100%">
<tr>
   	<td align="left" class="Estilo7" width="25%" valign="bottom" rowspan="1" colspan="2"><%=DATE()%></td>
	<td align="center" width="50%" rowspan="2">
		<img src="images/logo_tit.jpg" width="212" height="70" align="center" />
	</td>
	<td align="right" class="Estilo7" width="10%" valign="bottom">Tienda : </td>
	<td align="right" class="Estilo7" width="15%" id="tiend" valign="bottom"></td>
</tr>
<tr>
	<td align="left" class="Estilo7" width="5%" valign="top" rowspan="1">&nbsp;</Td>
	<td valign="top"  align="left">&nbsp;</td>
	
</tr>
<tr>
<td colspan="5" style="height:20px; background-color:#c82f8a; color:White" id="titu" name="titu"><font color:#000>&nbsp;</font></td>
</tr>
</table>



<iframe width="100%" onLoad="calcHeight();" src="inicioBanco.asp" id="body0" name="body0"  scrolling="yes" frameborder="3" height="1" align="middle">
</iframe>
</center>
</body>


</html>
