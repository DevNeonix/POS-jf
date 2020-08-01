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
IF  request.QueryString("EXCEL") = "1" THEN
  archivo = "c:\temp\RESUMEN_DETAexcel.xls"
    Response.Charset = "UTF-8"
    Response.ContentType = "application/vnd.ms-excel" 
    Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo 
END IF

'"RESUMENdeta.asp?pos=" + tienda + '&tipo=' + had + '&ini=' + document.all.ini.value + '&fin=' + document.all.fin.value
pos = request.QueryString("pos")
tip = request.QueryString("tipo")	'FAC 20130108 tipo: PG ó PP
ini = request.QueryString("ini")
fin = request.QueryString("fin")

'*************************************************************
'FAc 20131209 condicion de tienda
cad =   " set dateformat dmy; SELECT sale+entra as cant,*, convert(char(5),vv.fecha, 108) [hora], vv.fecha" & _
        " FROM VIEW_VENTAS_ARTICULO vv inner join tiendas tt on vv.tienda = tt.codigo           " & _
        " WHERE VV.FECHA between '"&INI&"' AND DateAdd(day,1,'"&FIN&"') and tipdoc in ('BL','FC','NC') "
if pos<>"TT" then cad = cad&" and TIENDA ='" & POS & "' "
cad = cad&"order by tipdoc,numdoc,operacion,item"
'*************************************************************
'FAC 20130108 RESPONSE.WRITE(CAD)
'RESPONSE.WRITE(CAD)
rs.open cad,cnn
if rs.recordcount <=0 then RESPONSE.End
%>

<body onload="AGRANDA()">
<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="Table1" name="listado"  >
<tr>
<td><input type="button" value="Exportar a Excel" onclick="REPORTE(2)" /></td>
</tr>
</table>



<center>

<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="listado" name="listado"  >
	<tr>
        <td align="center" class="Estilo8">TIENDA</td> 
        <td align="center" class="Estilo8">DOCUMENTO</td> 
        <td align="center" class="Estilo8">FECHA</td>
        <td align="center" class="Estilo8">HORA</td>
        <td align="center" class="Estilo8">CLIENTE</td>
        <td align="center" class="Estilo8">ARTICULO</td>
	    <td align="center" class="Estilo8">DESCRIPCION</td>
        <td align="center" class="Estilo8">CANT</td>
        <td align="center" class="Estilo8">PRECIO</td>
	</tr >

    <%I=0 
      tota = 0
      RS.MOVEFIRST
      F=0     
      MM=0
		ttlis=0	  
		ttdct=0
		ttigv=0
		ttprecio=0  
	  %>
    <%do while not rs.eof %>            
        <tr bgcolor='<%=application("color2") %>'  class="EstiloT" align="left">
            <td>&nbsp;<%=rs("tienda")&" - "&rs("descripcion")%></td>
             <td>&nbsp;<%=rs("tipdoc")&" - "&rs("numdoc")%></td>
             <td>&nbsp;<%=formatdatetime(rs("Fecha"),2)%>&nbsp;</td>
             <td align="center"><%=rs("hora")  						   %></td>
             <td>&nbsp;<%=rs("cliente")&" - "&rs("nombre")             %></td>
			 <td align="center">&nbsp;<%=ucase(RS("CodArt"))           %></td>
			 <td>&nbsp;<%=RS("descri")                                 %></td>
			 <td align="center"><%=rs("cant")                          %></td>
             <td align="right" ><%=formatnumber(rs("precio"),2,,true)  %></td>
		</tr> 
        <%rs.movenext%>
	<%loop %>
</table>
</center>
</body>
<%RS.CLOSE %>
<script language="jscript" type="text/jscript">
function REPORTE(op) { 
window.open('COMERCIOdeta.asp?pos=' + '<%=pos %>' + '&tipo=' + '<%=tip %>' + '&ini=' + '<%=ini %>' + '&fin=' + '<%=fin%>'+'&EXCEL=1')

}
</script>
</html>
