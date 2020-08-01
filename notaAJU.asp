<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = true %>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos") %>
<% tda = Request.Cookies("tienda")("tda") %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<%tit = request.QueryString("tit")
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=tit%></title>
</head>

<LINK REL="stylesheet" TYPE="text/css" HREF="VENTAS.CSS">
<!--#include file="comun/funcionescomunes.asp"-->
<!--#include file="includes/funcionesVBscript.asp"-->
<!--#include file="includes/cnn.inc"-->

<script language="jscript" type="text/jscript">
function AGRANDA() {
    top.parent.window.document.getElementById('body0').height = 480
}
</script>

<%ope = request.QueryString("ope") 
'*************************************************************
cad =   "set dateformat dmy; SELECT M1.OPERACION, M1.NUMDOC, M1.TIENDA, M1.CODDOC,          " & _
        " M1.SERIE, M1.FECDOC, V1.DESCRIPCION AS OBS, M2.DESCRIPCION AS CODMOT, V1.MOTI     " & _
        " FROM movimcab as m1 inner join varios as v1 on m1.operacion = v1.operacion        " & _
        " INNER JOIN MOTITRANS AS M2 ON V1.MOTI = M2.CODIGO WHERE m1.OPERACION = '"&OPE&"'; " 
'*************************************************************
'RESPONSE.WRITE(CAD)
'response.end
rs.open cad,cnn
if rs.recordcount <=0 then RESPONSE.End
%>

<body>
<center>
<table align="center" cellpadding="2" cellspacing="1" bordercolor='<%=application("color1") %>' border="0" >
	<tr  class="Estilo17"> 
    <td>NOTA DE INGRESO --> <%=rs("coddoc")%> &nbsp;<%=trim(rs("serie"))%>-<%=trim(rs("numdoc"))%></td>
    </tr>
</table>



<table align="center" cellpadding="2" cellspacing="2" bordercolor='<%=application("color1") %>' border="0" >
	<tr><td class="EstiloT" align="left">Fecha : </td>
    <td colspan="4" class="EstiloT" align="left"><%=formatdatetime(rs("fecdoc"),2) %></td></tr>
    <tr> 
        <td class="EstiloT" align="left">Tienda : </td>
        <td class="EstiloT" align="left" colspan="2"><%=rs("tienda") %>&nbsp;<%=tda %></td>
        <td class="EstiloT" align="right">Operacion : </td>
        <td class="EstiloT" align="right"><%=rs("operacion") %></td>
    </tr>
    <tr> 
        <td class="EstiloT" align="left">Motivo: </td>
        <td class="EstiloT" align="left"><%=rs("CODMOT")%>-<%=RS("MOTI") %></td>
        <td width="40%">&nbsp;</td>
        <td class="EstiloT" align="right">Destino : </td>
        <td class="EstiloT" align="right"><%=rs("OBS")%></td>
    </tr>
</table>

<%cad =   " set dateformat dmy;         " & _
        " SELECT * FROM VIEW_DETALLE    " & _
        " WHERE  OPERACION = '"&OPE&"'  " & _
        " order by item                 "
'RESPONE.WRITE (CAD)
rs.close
rs.open cad,cnn        
if rs.recordcount <=0 then RESPONSE.End         %>
<table align="center" cellpadding="2" cellspacing="1" bordercolor='<%=application("color1") %>' border="0" >
	<tr  class="Estilo17"> 
        <td align="center">IT</td>
        <td align="center">CODIGO</td>
        <td align="center">DESCRIPCION</td>
        <td align="center">CANT</td>
       
	</tr >
    <%can= 0 
     PVP = 0
     DCT = 0
     TOT = 0%>
<%do while not rs.eof %>
	<tr bgcolor='<%=application("color2") %>'  class="EstiloT" align="left">
			<td>&nbsp;<%=RS("item")%>&nbsp;</td>
            <td>&nbsp;<%=RS("codart")%>&nbsp;</td>
            <td>&nbsp;<%=RS("descri")%>&nbsp;</td>
        
            <td align="right">&nbsp;<%=FORMATNUMBER(RS("ENTRA"),0,,,TRUE)%>&nbsp;</td>
            <% can = can +  cdbl(RS("ENTRA"))%>
	</tr> 
    
    <%rs.movenext%>
<%loop %>
<tr class="Estilo17"  align="right">
 <td colspan="3" style="text-align:right">Total&nbsp;</td>
 <td align="right" ><%=can%></td>
</tr>

</table>

</center>
</body>
<%RS.CLOSE %>

</html>
