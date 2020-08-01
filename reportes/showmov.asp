<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = true %>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos") %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<%tit = request.QueryString("tit")
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=tit%></title>
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

<%ope = request.QueryString("ope") 
'*************************************************************
cad =   " set dateformat dmy;           " & _
        " SELECT *, '' as destino, '' as mot, '' as motiv FROM VIEW_CABECERAS  " & _
        " WHERE  OPERACION = '"&OPE&"'; " 


'*************************************************************
'RESPONSE.WRITE(CAD)
'response.end
rs.open cad,cnn
if rs.recordcount <=0 then RESPONSE.End
%>

<body>
<center>

<table align="center" cellpadding="0" cellspacing="0" bordercolor='<%=application("color1") %>' border="0" >
	<tr><td class="EstiloT" align="right" colspan="4">Fecha : </td>
    <td  class="EstiloT" align="right"><%=formatdatetime(rs("fecdoc"),2) %></td></tr>
    <tr> 
        <td class="EstiloT" align="left">Tienda : </td>
        <td class="EstiloT" align="left"><%=rs("tienda") %></td>
        <td width="40%">&nbsp;</td>
        <td class="EstiloT" align="right">Operacion : </td>
        <td class="EstiloT" align="right"><%=rs("operacion") %></td>
    </tr>
    <tr> 
        <td class="EstiloT" align="left">Documento : </td>
        <td class="EstiloT" align="left"><%=rs("coddoc")%>&nbsp;<%=trim(rs("serie"))%>-<%=trim(rs("numdoc"))%></td>
        <td>&nbsp;</td>
        <td class="EstiloT" align="right">Doc. Relacionado : </td>
        <td class="EstiloT" align="right"><%=rs("docori")%>&nbsp;<%=trim(rs("serori"))%>-<%=trim(rs("numori"))%></td>
    </tr>
    <%IF LEN(LTRIM(RTRIM(RS("DESTINO")))) > 0 THEN %>
    <tr> 
        <td class="EstiloT" align="left">Destino : </td>
        <td class="EstiloT" align="left"><%=rs("destino")%></td>
        <td>&nbsp;</td>
        <td class="EstiloT" align="right">Motivo: </td>
        <td class="EstiloT" align="right"><%=rs("mot")%>&nbsp;<%=trim(rs("motiv"))%></td>
    </tr>
    <%END IF %>
</table>

<%cad =   " set dateformat dmy;         " & _
        " SELECT * FROM VIEW_DETALLE    " & _
        " WHERE  OPERACION = '"&OPE&"'  " & _
        " order by item                 "
TIPMOV = RS("TIPMOV")
rs.close
rs.open cad,cnn        
'response.write(cad)
'response.end
if rs.recordcount <=0 then RESPONSE.End         %>
<table align="center" cellpadding="0" cellspacing="1" bordercolor='<%=application("color1") %>' border="0">
	<tr> 
        <td align="center" class="Estilo8">IT</td>
        <td align="center" class="Estilo8">CODIGO</td>
        <td align="center" class="Estilo8">DESCRIPCION</td>
        <td align="center" class="Estilo8">CANT</td>
        <%IF TIPMOV = "S" THEN %>
          <td align="center" class="Estilo8">PVP</td>
          <td align="center" class="Estilo8">DCT</td>
          <td align="center" class="Estilo8">%</td>
          <td align="center" class="Estilo8">IGV</td>
          <td align="center" class="Estilo8">TOT</td>
        <%END IF %>
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
        <%IF TIPMOV = "S" THEN %>
            <td align="right">&nbsp;<%=FORMATNUMBER(RS("sale"),2,,,TRUE)%>&nbsp;</td>
    		<td align="right">&nbsp;<%=FORMATNUMBER(RS("pvp"),2,,,TRUE)%>&nbsp;</td>
            <td align="right">&nbsp;<%=FORMATNUMBER(RS("descuento"),2,,,TRUE)%>&nbsp;</td>
            <td align="right">&nbsp;<%=FORMATNUMBER(RS("Pordes"),2,,,TRUE)%>&nbsp;</td>
            <td align="right">&nbsp;<%=FORMATNUMBER(RS("igv"),2,,,TRUE)%>&nbsp;</td>
    		<td align="right">&nbsp;<%=FORMATNUMBER(RS("total"),2,,,TRUE)%>&nbsp;</td>
            <%  can = can + cdbl(RS("sale"))
                pvp = pvp + cdbl(RS("pvp"))
                dct = dct + cdbl(RS("descuento"))
                igv = igv + cdbl(RS("igv"))
                tot = tot + cdbl(RS("total"))
            %>
        <%else%>
            <td>&nbsp;<%=FORMATNUMBER(RS("entra"),2,,,TRUE)%>&nbsp;</td>
            <% can = can +  cdbl(RS("entra"))%>
        <%END IF %>
	</tr> 
    
    <%rs.movenext%>
<%loop %>
<tr class="Estilo8"  align="right">
 <td colspan="3" style="text-align:right">Total&nbsp;</td>
<%if tipmov= "S" then %>
    <td align="center" ><%=can%></td>
    <td>&nbsp;<%=FORMATNUMBER(pvp,2,,,TRUE)%>&nbsp;</td>
    <td>&nbsp;<%=FORMATNUMBER(dct,2,,,TRUE)%>&nbsp;</td>
    <td></td>
    <td>&nbsp;<%=FORMATNUMBER(igv,2,,,TRUE)%>&nbsp;</td>
    <td>&nbsp;<%=FORMATNUMBER(tot,2,,,TRUE)%>&nbsp;</td>
<%else%> 
    <td align="center" ><%=can%></td>

<%end if%>
</tr>

</table>

</center>
</body>
<%RS.CLOSE %>

</html>
