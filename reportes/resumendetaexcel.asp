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
'"RESUMENdeta.asp?pos=" + tienda + '&tipo=' + had + '&ini=' + document.all.ini.value + '&fin=' + document.all.fin.value
pos = request.QueryString("pos")
tip = request.QueryString("tipo")	'FAC 20130108 tipo: PG ó PP
ini = request.QueryString("ini")
fin = request.QueryString("fin")

'*************************************************************

cad = "set dateformat dmy; SELECT sale+entra as cant,* FROM VIEW_VENTAS_ARTICULO WHERE FECHA between '"&INI&"' AND DateAdd(day,1,'"&FIN&"') and tipdoc in ('BL','FC','NC') "
if pos<>"TT" then cad = cad&" and TIENDA ='" & POS & "' "
cad = cad&"order by tipdoc,numdoc,operacion,item"
'*************************************************************
'FAC 20130108 RESPONSE.WRITE(CAD)
rs.open cad,cnn
if rs.recordcount <=0 then RESPONSE.End

    archivo = "c:\temp\registroDETA.xls"
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
        <td align="center" >%Dscto</td>
        <td align="center" >I.G.V</td>
        <td align="center" >Precio</td>
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
    	<%OPER = trim(RS("OPERACION")) %>
    	<%	F=0     
      		MM=0
			tlis=0
	  		tdct=0
	  		tigv=0
	  		tprecio=0 %>
            
        <tr bgcolor='<%=application("color2") %>' align="left">
            <td>&nbsp;<%=rs("tipdoc")&" - "&rs("numdoc")%></td>
            <td>&nbsp;<%=rs("cliente")&" - "&rs("nombre")%></td>
            <td>&nbsp;<%=formatdatetime(rs("Fecha"),2)%>&nbsp;</td>
        </tr>
        
        <%do while not rs.eof AND TRIM(RS("OPERACION")) = OPER%>
			<tr bgcolor='<%=application("color2") %>'  align="left">
				<td align="center">&nbsp;<%=ucase(RS("CodArt"))%>&nbsp;</td>
				<td>&nbsp;<%=RS("descri")%>&nbsp;</td>
				<td align="center">&nbsp;<%=rs("cant")%>&nbsp;</td>
				<td align="right">&nbsp;<%=formatnumber(rs("lista1"),2,,true)	%>&nbsp;</td>
                <td align="right">&nbsp;<%=formatnumber(rs("dct"),2,,true)		%>&nbsp;</td>
                <td align="center">&nbsp;<%=rs("pordes")%> 						%&nbsp;</td>
                <td align="right">&nbsp;<%=formatnumber(rs("IGV"),2,,true)		%>&nbsp;</td>
                <td align="right">&nbsp;<%=formatnumber(rs("precio"),2,,true)	%>&nbsp;</td>
			</tr> 
			<% if rs("tipdoc")<>"NC" then
				tlis = tlis + (cdbl(rs("lista1")) * cdbl(rs("cant")))
				tdct= tdct + cdbl(rs("dct"))
		   		tigv= tigv + cdbl(rs("IGV"))
				tprecio= tprecio + cdbl(rs("precio"))
				else
				tlis = tlis - (cdbl(rs("lista1")) * cdbl(rs("cant")))
				tdct= tdct - cdbl(rs("dct"))
		   		tigv= tigv - cdbl(rs("IGV"))
				tprecio= tprecio - cdbl(rs("precio"))
                end if
			%>
           <%I= I + 1%>
           <%F= F + 1%>
           <%rs.movenext%>
           <%IF RS.EOF THEN EXIT DO %>
        <%LOOP%>

        <tr align="right" bgcolor="gainsboro">
			<td colspan="3"><strong>&nbsp;Total Documento&nbsp;</strong></td>
            <td><strong>&nbsp;<%=formatnumber(tlis		,2,,true)%>&nbsp;</strong></td>
			<td><strong>&nbsp;<%=formatnumber(tdct		,2,,true)%>&nbsp;</strong></td>
			<td></td>
			<td ><strong>&nbsp;<%=formatnumber(tigv		,2,,true)%>&nbsp;</strong></td>
			<td ><strong>&nbsp;<%=formatnumber(tprecio	,2,,true)%>&nbsp;</strong></td>
        </tr>
        <%	
			ttlis= ttlis + tlis
			ttdct= ttdct + tdct
	   		ttigv= ttigv + tIGV
			ttprecio= ttprecio + tprecio
        	IF RS.EOF THEN EXIT DO
		%>
	<%loop %>
	<tr bgcolor= "<%=application("barra")%>"align="right">
        <td colspan="3"><strong>&nbsp;Total Documento&nbsp;</strong></td>
        <td><strong>&nbsp;<%=formatnumber(ttlis		,2,,true)%>&nbsp;</strong></td>        
        <td><strong>&nbsp;<%=formatnumber(ttdct   	,2,,true)%>&nbsp;</strong></td>
        <td></td>
        <td><strong>&nbsp;<%=formatnumber(ttigv   	,2,,true)%>&nbsp;</strong></td>
        <td><strong>&nbsp;<%=formatnumber(ttprecio	,2,,true)%>&nbsp;</strong></td>
	</tr>

</table>
</center>
</body>
<%RS.CLOSE %>
</html>
