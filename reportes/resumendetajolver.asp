<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%'Response.Buffer = false %>
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

cadCab = "select *,nomcli = (select top 1 NOMBRE from CLIENTES x where x.CLIENTE=c.CLIENTE) from JACINTA..movimcab c where TIENDA='06' and FECDOC between '09/05/2020' AND DateAdd(day,1,'11/05/2020') and CODDOC in ('BL','FC','NC') order by operacion"
rs.open cadCab,cnn
%>

<body onload="AGRANDA()">
<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="Table1" name="listado"  >
<tr>
<td><input type="button" value="Pantalla Completo" onclick="REPORTE(1)" /></td>
<td><input type="button" value="Pantalla Resumen" onclick="REPORTE(2)" /></td>
<td><input type="button" value="Excel Completo" onclick="REPORTE(3)" /></td>
<td><input type="button" value="Excel Resumen" onclick="REPORTE(4)" /></td>
</tr>
</table>



<center>

<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="listado" name="listado"  >
	<tr> 
        <td align="center" class="Estilo8">DOCUMENTO<br>Articulo</td>
	    <td align="center" class="Estilo8">CLIENTE<br>Descripción</td>
        <td align="center" class="Estilo8">FECHA<br />Unds.</td>
        <td align="center" class="Estilo8">P.Vta.</td>
        <td align="center" class="Estilo8">Dscto.</td>
        <td align="center" class="Estilo8">%Dscto</td>
        <td align="center" class="Estilo8">I.G.V</td>
        <td align="center" class="Estilo8">I.S.C</td>
        <td align="center" class="Estilo8">Precio</td>
        <td align="center" class="Estilo8">Tot. Doc/OP</td>
        <td align="center" class="Estilo8">Hora</td>
	</tr >
    <%  granSubtotal = 0%>

    <%if rs.recordcount > 0 then%>
    <%rs.movefirst%>
    <%for i=0 to rs.recordcount-1%>


    <%'CABECERA%>
    <TR class=EstiloT align=left bgColor=#ffffff>
        <TD>&nbsp;<%=rs("CODDOC")%> - <%=rs("serie")%>-<%=rs("numdoc")%> </TD>
        <TD>&nbsp;<%=rs("CLIENTE")%> - <%=rs("nomcli")%></TD>
        <TD>&nbsp;<%=left(rs("fecdoc"),10)%>&nbsp;</TD>
        <TD colSpan=6></TD>
        <TD align=center>&nbsp;<%=rs("operacion")%>&nbsp;</TD>
        <TD align=center><%=right(rs("fecdoc"),14)%></TD>
    </TR>
    <%'DETALLE%>

    <%
    set rsDet = RsNuevo

    cadDet = "SELECT * FROM VIEW_MOVIMDETART WHERE OPERACION = '"&rs("operacion")&"' order by item"

  


    rsDet.open cadDet,cnn
    %>
        
        <%if rsDet.recordcount > 0 then%>
            <%rsDet.movefirst%>
            <%subtotal = 0%>
            <%for d=0 to rsDet.recordcount-1%>

                <%subtotal = cdbl(rsDet("pventa")) + subtotal%>


                <TR class=EstiloT align=left bgColor=#ffffff>
                    <TD align=left>&nbsp;<%=rsDet("codart")%> &nbsp;</TD>
                    <TD>&nbsp;<%=rsDet("descri")%>&nbsp;</TD>
                    <TD align=center><%=rsDet("CANT")%>&nbsp;</TD>
                    <TD align=right><%=formatnumber(rsDet("pventa"),2,,true)%></TD>
                    <TD align=right><%=rsDet("descuento")%></TD>
                    <TD align=center><%=rsDet("pordes")%></TD>
                    <TD align=right><%=rsDet("igv")%></TD>
                    <TD align=right><%=formatnumber((CDBL(rsDet("ISC"))*cdbl(rs("isc")))*CDBL(rsDet("CANT")),2,,true)%></TD>
                    <TD align=right>0.00</TD>
                    <TD>&nbsp;</TD>
                    <TD align=center><%=right(rs("fecdoc"),14)%></TD>
                </TR>

                
            <%rsDet.movenext%>
            <%next%>

            

            <TR class=Estilo0 align=right bgColor=#dbdbdb><TD colSpan=3><STRONG>&nbsp;Total Documento&nbsp;</STRONG></TD>
                <TD><STRONG><%=formatnumber(subtotal,2)%></STRONG></TD>
                <TD><STRONG>0.00</STRONG></TD>
                <TD></TD>
                <TD><STRONG>39.18</STRONG></TD>
                <TD><STRONG>0.20</STRONG></TD>
                <TD><STRONG>256.90</STRONG></TD>
                <TD><STRONG>257.10</STRONG></TD>
            </TR>

            <%
            'AUMENTA GRAN SUBTOTAL'
            granSubtotal = granSubtotal + subtotal
            %>
            
        <%end if%>


    <%rs.movenext%>
    <%next%>


    <TR class=Estilo3 align=right bgColor=#f9c1d9><TD colSpan=3><STRONG>&nbsp;Total Documento&nbsp;</STRONG></TD>
    <TD><STRONG><%=granSubtotal%></STRONG></TD>
    <TD><STRONG>0.00</STRONG></TD>
    <TD></TD>
    <TD><STRONG>712.84</STRONG></TD>
    <TD><STRONG>3.80</STRONG></TD>
    <TD><STRONG>4,673.38</STRONG></TD>
    <TD><STRONG>4,677.18</STRONG></TD></TR>

    <%end if%>


    </table>
</html>
