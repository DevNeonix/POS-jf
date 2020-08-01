<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = true %>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos") %>
<%'**************************************************************************
' BORRA OPERACION Y ACTUALIZA STOCKS (PARA EL CASO DEL DOBLE INGRESO DE GUIAS)
' OPERACIONES DIFERENTES CON EL MISMO NUMERO DE DOCUMENTO 
'**************************************************************************** %>
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
rs.open cad,cnn
%><center>
<span class="Estilo18">
<%
if rs.recordcount <=0 then 
   response.Write("<br><br><br>")
    response.write("Operacion no Existe de acuerdo a lo solicitado ")
    RESPONSE.End
       
end if
rs.movefirst
tip = rs("tipmov")
rs.close
msg = ""
CAD = "SELECT *  FROM MOVIMDET  where operacion = '"&ope&"' ; "
RS.OPEN CAD, CNN
if rs.recordcount <=0 then 
    msg = "Operacion no tiene lineas de detalle, se ha eliminado Cabecera y Caja"
end if
rs.movefirst
cad= ""
do while not rs.eof
    cod = rs("codart")
    if tip = "E" then can = cdbl(rs("entra"))*-1 else can = cdbl(rs("sale"))
    tda = rs("tienda")
    cad = cad + " update articulos set stock = stock + "&can&" where tienda = '"&tda&"' and codigo = '"&cod&"' ; "
    rs.movenext
loop


cad =  cad +    " delete caja where operacion = '"&ope&"' ;     " & _
                " delete movimdet where operacion = '"&ope&"' ; " & _
                " DELETE MOVIMCAB WHERE OPERACION = '"&OPE&"'   "

'response.write(cad)
'response.end
CNN.EXECUTE CAD

msg= "Stock actualizado, y Movimientos Borrados"
%>
</span></center>
<body onload="AGRANDA()">
<center>

<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="TABLA" name="TABLA"  >
    <tr class="Estilo10"> 
        <td colspan="2">SE HA ELIMINADO LA OPERACION : <%=OPE%></td>    
    </tr>
    <tr class="Estilo10"> 
        <td colspan="2"><%=ucase(msg)%></td>    
    </tr>
    <%if tip="E" then 
    cad= "select * from articulos where tienda = '"&tda&"' and stock <0"
    rs.close
    rs.open cad,cnn
        if rs.recordcount > 0 then
            rs.movefirst
        %>
        <tr class="Estilo6"> 
            <td colspan="2">Revise los stocks negativos que se han generado</td>    
        </tr>
        <%do while not rs.eof %>
            <tr class="Estilo10"> 
                <td><%=rs("codigo")%></td>    
                <td><%=rs("stock")%></td>    
            </tr>

            <%rs.movenext%>
        <%loop%>
        <%end if %>
    <%end if %>
</table>

<iframe  width="100%" src="" id="body0" name="body0" scrolling="yes" frameborder="1" height="40" align="middle" style="display:none" ></iframe>


<script language="jscript" type="text/jscript">

</script>
</center>
</body>
</html>
