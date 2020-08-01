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
pos = request.QueryString("pos")
tip = request.QueryString("tipo")	'FAC 20130108 tipo: PG ó PP
ini = request.QueryString("ini")
fin = request.QueryString("fin")

'*************************************************************

cad =   " set dateformat dmy; " & _
        " SELECT CODDOC, SERIE, NUMDOC, CLIENTES.CLIENTE, CLIENTES.NOMBRE, MOVIMCAB.FECDOC AS FECHA,  " & _
        " PVP, DESCUENTO, IGV, TOTAL FROM MOVIMCAB INNER JOIN CLIENTES ON                   " & _ 
        " CLIENTES.CLIENTE = MOVIMCAB.CLIENTE WHERE                      " & _
        " movimcab.FECdoc between '"&INI&"' AND DateAdd(day,1,'"&FIN&"') and coddoc in ('BL','FC','NC') "
if pos<>"TT" then cad = cad&" and TIENDA ='" & POS & "' "
cad = cad&"order by coddoc,numdoc,operacion"
'*************************************************************

rs.open cad,cnn
if rs.recordcount <=0 then RESPONSE.End

    archivo = "c:\temp\registroshort.xls"
    Response.Charset = "UTF-8"
    Response.ContentType = "application/vnd.ms-excel" 
    Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo  %>

<body>
<table>
	<tr> 
        <td align="center" >DOCUMENTO<br>Articulo</td>
	    <td align="center" >CLIENTE<br>Descripción</td>
        <td align="center" >FECHA<br />Unds.</td>
        <td align="center" >P.Vta.</td>
        <td align="center" >Dscto.</td>
        <td align="center" >I.G.V</td>
        <td align="center" >Precio</td>
	</tr >

    <%RS.MOVEFIRST

		tlis=0	  
		tdct=0
		tigv=0
		tprecio=0  
	  %>
    <%do while not rs.eof %>
        <tr bgcolor='<%=application("color2") %>' align="left">
            <td>&nbsp;<%=rs("coddoc")&" - "&rs("SERIE") & " - " & rs("NUMDOC")%></td>
            <td>&nbsp;<%=rs("cliente")&" - "&rs("nombre")%></td>
            <td>&nbsp;<%=formatdatetime(rs("Fecha"),2)%>&nbsp;</td>
    		<td align="right"><%=formatnumber(rs("PVP"),2,,true)	%></td>
            <td align="right"><%=formatnumber(rs("DESCUENTO"),2,,true)		%></td>  
            <td align="right"><%=formatnumber(rs("IGV"),2,,true)		%></td>
            <td align="right"><%=formatnumber(rs("TOTAL"),2,,true)	%></td>
		</tr> 
			<% if rs("coddoc")<>"NC" then
				tlis = tlis + cdbl(rs("PVP"))
				tdct= tdct + cdbl(rs("dESCUENTO"))
		   		tigv= tigv + cdbl(rs("IGV"))
				tprecio= tprecio + cdbl(rs("TOTAL"))
				else
				tlis = tlis - cdbl(rs("PVP"))
				tdct= tdct - cdbl(rs("DESCUENTO"))
		   		tigv= tigv - cdbl(rs("IGV"))
				tprecio= tprecio - cdbl(rs("TOTAL"))
                end if
			%>
           <%rs.movenext%>

        <%LOOP%>

	<tr bgcolor= "gainsboro"align="right">
        <td colspan="3"><strong>&nbsp;Total Rango de Fechas &nbsp;</strong></td>
        <td><strong><%=formatnumber(tlis		,2,,true)%></strong></td>        
        <td><strong><%=formatnumber(tdct   	,2,,true)%></strong></td>
        <td><strong><%=formatnumber(tigv   	,2,,true)%></strong></td>
        <td><strong><%=formatnumber(tprecio	,2,,true)%></strong></td>
	</tr>

</table>
</center>
</body>
<%RS.CLOSE %>
</html>
