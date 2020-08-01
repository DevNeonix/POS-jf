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
fec = request.QueryString("fec")
fec2= request.QueryString("fec2")
cad = "exec SP_CAJA_DIA '"&tienda&"', '"&fec&"', '"&fec2&"'"
'RESPONSE.WRITE(CAD)
rs.open cad,cnn
if rs.recordcount <=0 then RESPONSE.End
%>

<body onload="AGRANDA()">
<center>
<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="listado" name="listado"  >
	<tr> 
   		<td align="center" class="Estilo8">Operacion</td>
        <td align="center" class="Estilo8">Documento</td>
        <td align="center" class="Estilo8">Total</td>
        <td align="center" class="Estilo8">&nbsp;Pago&nbsp;</td>
        <td align="center" class="Estilo8">Fecha</td>
	    <td align="center" class="Estilo8">Soles</td>
        <td align="center" class="Estilo8">US. $</td>
        <td align="center" class="Estilo8">ST S/.</td>
	</tr>
		<%I=0 
      	tota = 0
		RS.MOVEFIRST
      	F=0     
      	MM=0    %>
    	<%do while not rs.eof %>
    	<%OPER = RS("OPERACION") %>
    	<%F=0     
      	MM=0 %>
		<%do while not rs.eof AND RTRIM(LTRIM(RS("OPERACION"))) = RTRIM(LTRIM(OPER))%>
            
            <tr bgcolor='<%=application("color2") %>'  class="EstiloT" align="right">
                <td>&nbsp;<%IF f=0 THEN RESPONSE.WRITE(rs("operacion"))%>&nbsp;</td>
                <td>&nbsp;<%IF f=0 THEN response.write(rs("coddoc")&"-"&rs("numdoc"))%>&nbsp;</td>
                <td>&nbsp;<%IF f=0 THEN response.write(formatnumber(rs("total"),2,,true))%>&nbsp;</td>
                <td>&nbsp;<%=rs("tipo")%>&nbsp;</td>
                <td>&nbsp;<%=formatdatetime(rs("fecha"),2)%>&nbsp;</td>
                <td>&nbsp;<%=FORMATNUMBER(RS("SOL"),2,,TRUE)%>&nbsp;</td>
                <td>&nbsp;<%if CDBL(rs("DOLAR"))<>CDBL(RS("SOL")) then response.write(formatnumber(rs("DOLAR"),2,,true))%>&nbsp;</td>
                <td><% MM = MM + CDBL(RS("SOL"))
                       tota = tota + CDBL(RS("SOL"))%>&nbsp;
                </td>
           </tr> 
           <%I= I + 1%>
           <%F= F + 1%>
           <%rs.movenext%>
           <%IF RS.EOF THEN EXIT DO %>
        <%LOOP%>
        <tr bgcolor= "<%=application("color1")%>" class="Estilo0" align="right">
        <td colspan="5">&nbsp;Total Documento&nbsp;</td>
        <td >&nbsp;<%=formatnumber(mm   ,2,,true) %>&nbsp;</td>
        </tr>
        <%IF RS.EOF THEN EXIT DO %>
   <%loop %>
    <tr bgcolor= "<%=application("barra")%>" class="Estilo3" align="right">
        <td colspan="5">&nbsp;Total Ingresos: S/: <%=fec%>  &nbsp;</td>
        <td >&nbsp;<%=formatnumber(tota   ,2,,true) %>&nbsp;</td>
        </tr>

</table>
<p>&nbsp;</p>
<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="Table1" name="listado"  >
   <tr> <td align="center" class="Estilo8" colspan="3">RESUMEN:</td>   </tr>
   <tr> <td align="center" class="Estilo8">Tipo</td>  
    <td align="center" class="Estilo8">Soles</td>  
    <td align="center" class="Estilo8">Dolares</td>   </tr>

<%RS.CLOSE
cad = "exec SP_RESUMEN_CAJA_DIA '"&tienda&"', '"&fec&"' "
'RESPONSE.WRITE (CAD)
rs.open cad,cnn 
RS.MOVEFIRST%>
<%do while not rs.eof %>
    <tr bgcolor='<%=application("color2") %>'  class="EstiloT" align="right">
         <td>&nbsp;<%=rs("tipo")%>&nbsp;</td>
         <td>&nbsp;<%=formatnumber(rs("soles"),2,,true)%>&nbsp;</td>
         <td>&nbsp;<%if not isnull(rs("dolares")) then response.write(formatnumber(rs("dolares"),2,,true))%>&nbsp;</td>
    </tr>
    <%rs.movenext%>  
<%LOOP%>
 <tr> 
    <td align="center" class="Estilo8">Total S/.</td>  
    <td align="center" class="Estilo8">&nbsp;<%=formatnumber(tota   ,2,,true) %>&nbsp;</td>  
    <td align="center" class="Estilo8">&nbsp;</td>  
     </tr>

</table>

</center>
</body>

</html>
