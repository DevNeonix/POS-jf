<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = true %>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos") %>
<%'***********************************************************************************
' BORRA OPERACION SIN ACTUALIZAR STOCKS 
'************************************************************************************* %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<LINK REL="stylesheet" TYPE="text/css" HREF="VENTAS.CSS">
<!--#include file="comun/funcionescomunes.asp"-->
<!--#include file="includes/funcionesVBscript.asp"-->
<!--#include file="includes/cnn.inc"-->
<!--#include file="comun/comunQRY.asp"-->
<script language="jscript" type="text/jscript">
function AGRANDA() {
    top.parent.window.document.getElementById('body0').height = 480
}
var oldrow = 1

function dd2(ff) {	// LLENA TEXTBOX ADICIONALES AL COMUN
    // LOS DEL COMUN SON CODIGO Y DESCRIPCION
    var t = document.all.TABLA;
    var pos = parseInt(ff);
    dd(ff, 0);
}
</script>
<%
TDA = request.QueryString("TDA")
ope = request.QueryString("oper")
if trim(tda) = "TT" then
    cad =   " select * from movimcab where operacion =  '"&ope&"' "
else
    cad =   " select * from movimcab where tienda = '"&TDA&"' and operacion =  '"&ope&"' "
end if
'RESPONSE.WRITE(cAD)
'response.end
%><center>
<span class="Estilo18">
<%
rs.open cad,cnn
if rs.recordcount <=0 then 
    response.Write("<br><br><br>")
    response.write("Operacion no Existe en la tienda indicada ")
    RESPONSE.End

end if

cad =   " delete caja where operacion = '"&ope&"' ;     " & _
        " delete movimdet where operacion = '"&ope&"' ; " & _
        " DELETE MOVIMCAB WHERE OPERACION = '"&OPE&"'   "
CNN.EXECUTE CAD

%>
</span></center>
<body onload="AGRANDA()">
<center>

<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="TABLA" name="TABLA"  >
    <tr> 
        <td>SE HA ELIMINADO LA OPERACION : <%=OPE%></td>    
    </tr>
</table>

<iframe  width="100%" src="" id="body0" name="body0" scrolling="yes" frameborder="1" height="40" align="middle" style="display:none" ></iframe>


<script language="jscript" type="text/jscript">

</script>
</center>
</body>
</html>
