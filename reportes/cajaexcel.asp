<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = true %>
<%Session.LCID=2058%>

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

<%tienda = request.QueryString("tda")
fec = request.QueryString("fec")
fec2 = request.QueryString("fec2")
cad = "exec SP_CAJA_DIA '"&tienda&"', '"&fec&"', '"&fec2&"' "
'RESPONSE.WRITE (CAD)
rs.open cad,cnn
if rs.recordcount <=0 then RESPONSE.End

  archivo = "c:\temp\cajaexcel.xls"
    Response.Charset = "UTF-8"
    Response.ContentType = "application/vnd.ms-excel" 
    Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo 

    %>

<body onload="AGRANDA()">
<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="Table2" name="listado"  >
<tr>
<td><input type="button" value="Excel " onclick="REPORTE(1)" /></td>


</tr>
</table>
<center>
<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="listado" name="listado"  >
	<tr> 
   		<td align="center" class="Estilo8">Operación</td>
        <td align="center" class="Estilo8">RUC</td>
        <td align="center" class="Estilo8">Cliente</td>
        <td align="center" class="Estilo8">Fecha</td>
        <td align="center" class="Estilo8">Documento</td>
        <td align="center" class="Estilo8">Total<br />Docto.</td>
        <td align="center" class="Estilo8">S/.</td>
        <td align="center" class="Estilo8">Visa</td>
        <td align="center" class="Estilo8">Mcard</td>
        <td align="center" class="Estilo8">N/C</td>
        <td align="center" class="Estilo8">Total<br>Dia</td>
	</tr>
	<%I=0 
    tota = 0
    tlnsol=0
	tlnvis=0
	tlnmas=0
	tlnncr=0
	RS.MOVEFIRST
    F=0     
    MM=0    %>
	<%do while not rs.eof
    	OPER = RS("OPERACION")
    	F=0     
      	MM=0 
		lnsol=0
		lnvis=0
		lnmas=0
		lnncr=0
		'ldfec1=formatdatetime(rs("fecha"),2)%>
		<tr bgcolor='<%=application("color2") %>'  class="EstiloT" align="right">
        	<td><%IF f=0 THEN RESPONSE.WRITE(rs("operacion"))%></td>
            <td><%=rs("Cliente")%>&nbsp;</td>
            <td align="left"><%=rs("Nombre")%></td>
            <td><%IF f=0 THEN response.write(formatdatetime(rs("fecha"),2))%></td>
            <td><%IF f=0 THEN response.write(rs("coddoc")&"-"&rs("numdoc"))%></td>
            <td><%IF f=0 THEN response.write(formatnumber(rs("total"),2,,true))%></td>
            <%do while not rs.eof AND RTRIM(LTRIM(RS("OPERACION"))) = RTRIM(LTRIM(OPER))
				select case rs("tipo")
				case "SOL"
					lnsol = lnsol + CDBL(rs("SOL"))
				case "VIS"
					lnvis = lnvis + CDBL(rs("SOL"))
				case "MAS"
					lnmas = lnmas + CDBL(rs("SOL"))
				case "NCR"
					lnncr = lnncr + CDBL(rs("SOL"))
				end select
				MM = MM + CDBL(RS("SOL"))
                tota = tota + CDBL(RS("SOL"))
                tlnsol=tlnsol+lnsol
	            tlnvis=tlnvis+lnvis
	            tlnmas=tlnmas+lnmas
	            tlnncr=tlnncr+lnncr
				rs.movenext
				IF RS.EOF THEN EXIT DO
			loop%>
			<td><%=FORMATNUMBER(lnsol,2,,,TRUE)%></td>
			<td><%=FORMATNUMBER(lnvis,2,,,TRUE)%></td>
			<td><%=FORMATNUMBER(lnmas,2,,,TRUE)%></td>
			<td><%=FORMATNUMBER(lnncr,2,,,TRUE)%></td>
			<td><%=formatnumber(MM,2,,,true)%></td>
		</tr> 
        <%I= I + 1%>
        <%F= F + 1%>
		<!--tr bgcolor= "<'%=application("color1")%>" class="Estilo0" align="right">
			<td colspan="10">&nbsp;Total Dia&nbsp;</td>
			<td >&nbsp;<%'=formatnumber(mm   ,2,,true) %>&nbsp;</td>
        </tr !-->
        <%IF RS.EOF THEN EXIT DO %>
	<%loop %>
    <tr bgcolor= "<%=application("barra")%>" class="Estilo3" align="right">
        <td colspan="5">Total Ingresos: S/: <%=fec%>  </td>
        <td ><%=formatnumber(tota   ,2,,true) %></td>
        <%RS.CLOSE
    CAD = "exec SP_CAJA_DIA_TOT '"&tienda&"', '"&fec&"', '"&fec2&"', 'SOL' " 
    RS.OPEN CAD,CNN
    IF RS.RECORDCOUNT > 0 THEN 
        RS.MOVEFIRST  
        tlnsol = RS("TOT") 
    ELSE 
        tlnsol= 0 
    END IF%>

        <td >&nbsp;<%=formatnumber(tlnsol   ,2,,true) %>&nbsp;</td>
         <%RS.CLOSE
    CAD = "exec SP_CAJA_DIA_TOT '"&tienda&"', '"&fec&"', '"&fec2&"', 'VIS' " 
    RS.OPEN CAD,CNN
    IF RS.RECORDCOUNT > 0 THEN 
        RS.MOVEFIRST  
        tlnvis = RS("TOT") 
    ELSE 
        tlnvis= 0 
    END IF%>
        <td >&nbsp;<%=formatnumber(tlnvis   ,2,,true) %>&nbsp;</td>
         <%RS.CLOSE
    CAD = "exec SP_CAJA_DIA_TOT '"&tienda&"', '"&fec&"', '"&fec2&"', 'MAS' " 
    RS.OPEN CAD,CNN
    IF RS.RECORDCOUNT > 0 THEN 
        RS.MOVEFIRST  
        tlnmas = RS("TOT") 
    ELSE 
        tlnmas= 0 
    END IF%>
        <td >&nbsp;<%=formatnumber(tlnmas   ,2,,true) %>&nbsp;</td>
        <%RS.CLOSE
    CAD = "exec SP_CAJA_DIA_TOT '"&tienda&"', '"&fec&"', '"&fec2&"', 'NCR' " 
    RS.OPEN CAD,CNN
    IF RS.RECORDCOUNT > 0 THEN 
        RS.MOVEFIRST  
        tlnncr = RS("TOT") 
    ELSE 
        tlnncr= 0 
    END IF%>
        <td >&nbsp;<%=formatnumber(tlnncr   ,2,,true) %>&nbsp;</td>
        <td >&nbsp;<%=formatnumber(tota   ,2,,true) %>&nbsp;</td>
	</tr>




	</tr>

</table>
<p>&nbsp;</p>
<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="Table1" name="listado"  >
   <tr> <td align="center" class="Estilo8" colspan="3">RESUMEN:</td>   </tr>
   <tr> <td align="center" class="Estilo8">Tipo</td>  
    <td align="center" class="Estilo8">Soles</td>  
    <td align="center" class="Estilo8">Dolares</td>   </tr>

<%RS.CLOSE%>
<tr bgcolor='<%=application("color2") %>' > 
    <td align="left" class="EstiloT">SOLES</td>  
    <td align="right" class="EstiloT"><%=formatnumber(TLNSOL   ,2,,true) %></td>  
    <td align="center" class="EstiloT">&nbsp;</td>  
</tr>
<tr bgcolor='<%=application("color2") %>' > 
    <td align="left" class="EstiloT">VISA</td>  
    <td align="right" class="EstiloT"><%=formatnumber(TLNVIS   ,2,,true) %></td>  
    <td align="center" class="EstiloT">&nbsp;</td>  
</tr>
<tr bgcolor='<%=application("color2") %>' > 
    <td align="left" class="EstiloT">MASTER</td>  
    <td align="right" class="EstiloT"><%=formatnumber(TLNMAS   ,2,,true) %></td>  
    <td align="center" class="EstiloT">&nbsp;</td>  
</tr>
 <tr bgcolor='<%=application("color2") %>' > 
    <td align="left" class="EstiloT">N. CRE.</td>  
    <td align="right" class="EstiloT"><%=formatnumber(TLNNCR   ,2,,true) %></td>  
    <td align="center" class="EstiloT">&nbsp;</td>  
</tr>
<tr> 
    <td align="left" class="Estilo8">Total S/.</td>  
    <td align="right" class="Estilo8"><%=formatnumber(tota   ,2,,true) %></td>  
    <td align="center" class="Estilo8">&nbsp;</td>  
</tr>
</table>

</center>
</body>
</html>
