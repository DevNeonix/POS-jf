<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = true %>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos") %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<LINK REL="stylesheet" TYPE="text/css" HREF="../VENTAS.CSS">
<!--#include file="../comun/funcionescomunes.asp"-->
<!--#include file="../includes/funcionesVBscript.asp"-->
<!--#include file="../includes/cnn.inc"-->

<script language="jscript" type="text/jscript">
function AGRANDA() {
    top.parent.window.document.getElementById('body0').height = 480
}


</script>
<%
TDA = request.QueryString("TDA")
tip = request.QueryString("tipo")
ini = request.QueryString("ini")
fin = request.QueryString("fin")
cod = request.QueryString("cod")
'*********************************************************************************************
cad =   " EXEC SP_ENTRA_SALE_FECHA  '"&INI&"' , '"&FIN&"', '"&TDA&"', '"&COD&"', '"&TIP&"' "
'*********************************************************************************************
'RESPONSE.WRITE(CAD)
'response.end
rs.open cad,cnn
if rs.recordcount <=0 then RESPONSE.End
%>

<body onload="AGRANDA()">
<center>

<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="listado" name="listado"  >
	<tr> 
    <%FOR I=0 TO RS.FIELDS.COUNT-1 %>
        <td align="center" class="Estilo8"><%=RS.FIELDS(I).NAME %></td>
    <%NEXT %>
	</tr >
    <%ing= 0 
    sal = 0%>
<%do while not rs.eof %>
	<tr bgcolor='<%=application("color2") %>'  class="EstiloT" align="left">
		<td align="center">&nbsp;<%=ucase(RS.FIELDS.ITEM(0))%>&nbsp;</td>
		<td>&nbsp;<%=ucase(RS.FIELDS.ITEM(1))%>&nbsp;</td>
		<td align="right"  ondblclick="show('E','<%=trim(ucase(RS.FIELDS.ITEM(0)))%>')" style="cursor:pointer"><%=formatnumber(rs("INGRESOS"),0,,true)%>&nbsp;</td>
        <td align="right" ondblclick="show('S','<%=ucase(RS.FIELDS.ITEM(0))%>')" style="cursor:pointer"><%=formatnumber(rs("SALIDAS"),0,,true)%>&nbsp;</td>
        <%ing =  ing + cdbl(rs("ingresos")) %>
        <%sal =  sal + cdbl(rs("salidas")) %>
	</tr> 
    <%rs.movenext%>
<%loop %>

<tr   class="Estilo3" align="left">
			<td align="center">&nbsp;&nbsp;</td>
			<td>&nbsp;</td>
			<td align="right" >&nbsp;<%=formatnumber(ing,0,,true)	%>&nbsp;</td>
            <td align="right">&nbsp;<%=formatnumber(SAL,0,,true)		%>&nbsp;</td>
		</tr> 
</table>

<iframe  width="100%" src="" id="body0" name="body0" scrolling="yes" frameborder="1" height="40" align="middle" style="display:none" ></iframe>

<%RS.CLOSE %>
<script language="jscript" type="text/jscript">
function show(op,cod) {
    tipo = '<%=TIP%>'
    if (tipo == 'PG') {
        alert("El detalle solo se muestra por producto")
        return true

    }
        cad = 'showOPERA.asp?tda=' + '<%=tda%>' + '&cod=' +cod + '&ini=' + '<%=ini%>' + '&fin=' + '<%=fin%>' + '&op=' + op + '&TIP=' + '<%=TIP%>'

        document.all.body0.style.display = 'block'
        document.all.body0.height = "250"
        document.all.body0.width = "90%"
        document.all.body0.src = cad
        return true
    }
</script>
</center>
</body>
</html>
