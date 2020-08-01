<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<% tienda = Request.Cookies("tienda")("pos") %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="COMUN/FUNCIONESCOMUNES.ASP"-->
<!--#include file="COMUN/COMUNqry.ASP"-->
<!--#include file="includes/Cnn.inc"-->
<% 
opr = request.querystring("op")
 tda = tienda
'**************************************************************************************
' muestra los  datos a imprimir
'**************************************************************************************
' SE HICIERON LOS CAMBIOS PARA QUE TOME LA COLUMNA ADIVCIONAL DE TEXTO
' EN MOVIMIENTOS DETALLE A SOLICITUD DE CECILIA PARA ;LAS NOTAS ESPECIALES
' QUE NO MUEVEN STOCK
' 20-6-2014  --> MMB
CAD=    " SELECT M1.CLIENTE, C1.NOMBRE, C1.DIRECCION, M2.CODART, M1.FECDOC, " & _
        " M1.SERIE, M1.NUMDOC ,M2.ENTRA, LTRIM(RTRIM(V1.DESCRI)) COLLATE    " & _
        " Modern_Spanish_CI_AI + ' ' +    " & _
        " LTRIM(RTRIM(M2.TEXTO)) AS DESCRI,                                 " & _ 
        " ROUND(M2.PRECIO,2) AS PRECIO, M2.PORDES, m1.pvp as bruto,         " & _
        " (M2.PRECIO - M2.IGV + M2.DESCUENTO) / CASE WHEN M2.ENTRA = 0      " & _
        " THEN 1 ELSE ENTRA END AS PVT,                                     " & _
        " M1.TOTAL, M1.DESCUENTO, m1.subtot, m1.igv, m1.docori,             " & _
        " m1.serori, m1.numori                                              " & _
        " FROM MOVIMCAB AS M1                                               " & _
        " INNER JOIN CLIENTES AS C1 ON C1.CLIENTE = M1.CLIENTE              " & _
        " INNER JOIN MOVIMDET AS M2 ON M1.OPERACION = M2.OPERACION          " & _
        " inner join view_ARTICULOS_TIENDA V1 ON M2.CODART = V1.CODIGO      " & _
        " WHERE M1.OPERACION ='"&OPR&"' AND M2.ENTRA > = 0                  " & _
        " AND V1.TIENDA = '"&TDA&"'                                         " & _
        " ORDER BY M2.ITEM                                                  "
'response.Write(cad)
'RESPONSE.ENd
RS.OPEN CAD,CNN

RS.MOVEFIRST         
total = round(cdbl(rs("total")),2)
descuento = round(cdbl(rs("descuento")),2)
bruto = round(cdbl(rs("bruto")),2)
IGV = round(cdbl(rs("IGV")),2)
subtot = round(cdbl(rs("subtot")),2)

%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>

<style type="text/css">
body { font-family:Century Gothic;  font-size:10px;  color:Black; padding-left:14px;  padding-top:150px; width:750px}
body2 { font-family:Century Gothic;  font-size:11px;  color:Black; padding-left:14px; padding-top:150px;  }
</style>

</head>
<%DIM AMES(12) 
AMES(1) = "ENERO"
AMES(2) = "FEBRERO"
AMES(3) = "MARZO"
AMES(4) = "ABRIL"
AMES(5) = "MAYO"
AMES(6) = "JUNIO"
AMES(7) = "JULIO"
AMES(8) = "AGOSTO"
AMES(9) = "SETIEMBRE"
AMES(10) = "OCTUBRE"
AMES(11) = "NOVIEMBRE"
AMES(12) = "DICIEMBRE" %>
<body>
<table width="100%" align="left" border="0">
<tr>
<td align="center">
	<img src="images/print.jpg" border="0" id="prn" name="prn" onClick="printa();" style="display:block; cursor:pointer" />
</td>
</tr>
<tr>
<td>
	<table border="0"   align="right" cellpadding="0"cellspacing="0" width="100%">
    	<tr height="20px"  valign="bottom"><td align="right"><%=trim(rs("serie")) %> - <%=trim(rs("numdoc")) %></td><td width="18%">&nbsp; </td></tr>
    	<tr height="20px" ><td colspan="2">&nbsp;</td></tr>
	</table>
</td>
</tr>
<tr>
<td>
	<table border="0"   align="left" cellpadding="0"cellspacing="0" width="700px">
    	<tr  valign="bottom">
        	<td width="60px">&nbsp;</td>
        	<td  align="left" ><%=RS("CLIENTE") %>&nbsp;&nbsp;<%=left(trim(RS("NOMBRE")),65) %></td>
        	<td align="left">&nbsp;</td>
        	<td align="right" >&nbsp;</td>
    	</tr>
    	<tr>
        	<td >&nbsp;</td>
        	<td align="center">&nbsp;</td>
        	<td align="left">&nbsp;</td>
        	<td align="right" ><%=DAY(rs("fecdoc")) %> de <%=AMES(MONTH(RS("FECDOC"))) %> del <%=YEAR(RS("FECDOC")) %></td>
    	</tr>
    	<tr valign="bottom" height="8px">
        	<td >&nbsp;</td>
        	<td  align="left"><%=TRIM(RS("docori")) %>  <%=TRIM(RS("serori")) %>-<%=TRIM(RS("numori")) %></td>
        	<td align="left">&nbsp;</td>
        	<td align="center">&nbsp;</td>
    	</tr>
	</table>
</td>
</tr>
<tr>
<td>
	<table border="0"   align="left" cellpadding="0"cellspacing="0" width="100%">
    	<tr height="45px" valign="bottom"><td>&nbsp;</td></tr>
	</table>
</td>
</tr>
<tr>
<td>
	<table border="0"   align="left" cellpadding="0"cellspacing="0" WIDTH="700px">
		<%do while not rs.eof %>
            <tr  valign="bottom">
                <td width="28px" align="left"><%=ucASE(left(trim(RS("codart")),10)) %></td>
                <td width="60px" align="right"><%=RS("entra") %>&nbsp;</td>
                <td width="475px" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <%=left(trim(RS("DESCRI")),65) %></td>
                <td width="40px" align="right"><%=formatnumber(RS("PVT"),2,,true) %>&nbsp;</td>
                
                <td width="125px" align="right"><%=formatnumber(PRECIO,2,,true)%>&nbsp;</td>
                <%i = i + 1 %>
        </tr>
            <%rs.movenext %>
        <%loop%>
        <%for m=i to 14%>
        <tr  valign="bottom">
            <td align="left">&nbsp;</td>
            <td align="right">&nbsp;</td>
            <td align="left">&nbsp;</td>
            <td align="right">&nbsp;</td>
            <td align="right">&nbsp;</td>
        </tr>
        <%next %>
	</table>
</td>
</tr>
<tr>
<td>
	<table border="0"   align="left" cellpadding="0"cellspacing="0" width="700px" >
    	<tr  valign="bottom" height="65px" >
        	<td width="20px">&nbsp;</td>
        	<td id="sonn"></td>
        	<td align="right" width="80px"><%=formatnumber(SUBTOT,2,,true) %>&nbsp;</td>
       		<td align="right" width="80px"><%=formatnumber(IGV,2,,true) %>&nbsp;</td>
        	<td align="right" width="80px"><%=trim(formatnumber(TOTAL,2,,true)) %>&nbsp;</td>
    	</tr>
	</table>
</td>
</tr>
</table>

<script language="jscript" type="text/jscript">
    document.getElementById('sonn').innerText = 'Son : ' + FComson('<%=total%>')

function printa() {
    document.all.prn.style.display = 'none'
    window.print()
    this.window.close()
}
</script> 
</body>
</html>
