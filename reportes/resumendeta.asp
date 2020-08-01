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

'*************************************************************
'FAc 20131209 condicion de tienda
cad =   " set dateformat dmy; SELECT sale+entra as cant,*, convert(char(5),fecha, 108) [hora]" & _
        " FROM [VIEW_VENTAS_ARTICULO_2] WHERE FECHA between '"&INI&"' AND DateAdd(day,1,'"&FIN&"') and tipdoc in ('BL','FC','NC') "
if pos<>"TT" then cad = cad&" and TIENDA ='" & POS & "' "
cad = cad&"order by tipdoc,numdoc,operacion,item"
'*************************************************************
'FAC 20130108 RESPONSE.WRITE(CAD)
'RESPONSE.WRITE(CAD)
'response.end
rs.open cad,cnn
if rs.recordcount <=0 then RESPONSE.End
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

    <%I=0 
      tota = 0
      RS.MOVEFIRST
      F=0     
      MM=0
      TTTOT = 0
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
            TISC= 0
	  		tprecio=0 

        %>
            
        <tr bgcolor='<%=application("color2") %>'  class="EstiloT" align="left">

            <%
            doccccori = ""
            if rs("numori") <> "" AND ucase(rs("tipdoc")) = "NC" then
                Set RSx = Server.CreateObject("ADODB.Recordset")
                RSx.ActiveConnection = Cnn
                RSx.CursorType       = 3 'CONST adOpenStatic = 3
                RSx.LockType         = 1 'CONST adReadOnly = 1
                RSx.CursorLocation   = 3 'CONST adUseClient = 3

                if rsx.state>0 then
                    rsx.close
                end if
                oricad="exec SP_DOCORI '"&rs("DOCori")&rs("numori")&"'"
                'response.write(oricad)
                rsx.open oricad,cnn
                if rsx.recordcount>0 then
                    rsx.movefirst
                    doccccori = formatdatetime(rsX("FECDOC"),2)
                end if

            end if
            %>

            <td>&nbsp;<%=rs("tipdoc")&" - "&rs("numdoc")%></td>
            <td>&nbsp;<%=rs("cliente")&" - "&rs("nombre")%></td>
            <td>&nbsp;<%=formatdatetime(rs("Fecha"),2)%>&nbsp;<%IF doccccori<>"" THEN%>(FEC-ORI:<%=doccccori%>)&nbsp;<%END IF%></td>
            <%doccccori=""%>
            <td colspan="6">
            <%if isnull(rs("docori")) then %>
            <%else%>
            <%=trim(rs("docori"))+" --> "+trim(left(rs("serori"),3))+"-"+ (trim(rs("numori")))%>      
            <%end if %></td>
            <td align="center">&nbsp;<%=rs("operacion")%>&nbsp;</td>
            <td align="center"><%=rs("hora")  						%></td>
        </tr>
        <% if ucase(rs("tipdoc")) = "NC" then   TOTAL = CDBL(RS("TOTAL"))*-1 ELSE   TOTAL = CDBL(RS("TOTAL"))%>
        <%do while not rs.eof AND TRIM(RS("OPERACION")) = OPER%>
        <% if ucase(rs("tipdoc")) = "NC" then  precios = cdbl(rs("precio"))*-1  else  precios = cdbl(rs("precio")) %>
			<tr bgcolor='<%=application("color2") %>'  class="EstiloT" align="left">
				<td align="center">&nbsp;<%=ucase(RS("CodArt"))                 %>&nbsp;</td>
				<td>&nbsp;<%=RS("descri")                                       %>&nbsp;</td>
				<td align="center"><%=rs("cant")                          %>&nbsp;</td>
				<td align="right" ><%=formatnumber(rs("lista1"),2,,true)	%></td>
                <td align="right" ><%=formatnumber(rs("dct"),2,,true)		%></td>
                <td align="center"><%=rs("pordes")  						%></td>
                <td align="right" ><%=formatnumber(rs("IGV"),2,,true)		%></td>
                <td align="right" ><%=formatnumber((CDBL(rs("ISC"))*cdbl(rs("ISCprice")))*CDBL(RS("CANT")),2,,true)		%></td>
                <td align="right" ><%=formatnumber(precio,2,,true) 	%></td>
                <td >&nbsp;</td>
                <td align="center"><%=rs("hora")  						%></td>
			</tr> 
			<% 'if ucase(rs("tipdoc"))<>"NC" then
				    tlis = tlis + (cdbl(rs("lista1")) * cdbl(rs("cant")))
				    tdct= tdct + cdbl(rs("dct"))
		   		    tigv= tigv + cdbl(rs("IGV"))
                    tiSC= tiSC + (CDBL(rs("ISC"))*cdbl(rs("ISCprice")))*CDBL(RS("CANT"))
				    tprecio= tprecio + precios
				'else
				'    tlis = tlis - (cdbl(rs("lista1")) * cdbl(rs("cant")))
				'    tdct= tdct - cdbl(rs("dct"))
		   		'    tigv= tigv - cdbl(rs("IGV"))
                '    TISC = TISC - (CDBL(rs("ISC"))*0.1)*CDBL(RS("CANT"))
				'    tprecio= tprecio - cdbl(rs("precio"))
                'end if
			%>
           <%I= I + 1%>

           <%F= F + 1%>
           <%rs.movenext%>
           <%IF RS.EOF THEN EXIT DO %>
        <%LOOP%>
        
        <tr bgcolor= "<%=application("color1")%>" class="Estilo0" align="right">
			<td colspan="3"><strong>&nbsp;Total Documento&nbsp;</strong></td>
            <td><strong><%=formatnumber(tlis		,2,,true)%></strong></td>
			<td><strong><%=formatnumber(tdct		,2,,true)%></strong></td>
			<td></td>
			<td ><strong><%=formatnumber(tigv		,2,,true)%></strong></td>
            <td ><strong><%=formatnumber(tiSC		,2,,true)%></strong></td>
			<td ><strong><%=formatnumber(tprecio	,2,,true)%></strong></td>
            <td><strong><%=formatnumber(CDBL(TOTAL)	,2,,true)%></strong></td>
        </tr>
        <%	
			ttlis= ttlis + tlis
			ttdct= ttdct + tdct
	   		ttigv= ttigv + tIGV
            ttiSC= ttiSC + tISC
			ttprecio= ttprecio + tprecio
            TTTOT = TTTOT + TOTAL
        	IF RS.EOF THEN EXIT DO
		%>
	<%loop %>
	<tr bgcolor= "<%=application("barra")%>" class="Estilo3" align="right">
        <td colspan="3"><strong>&nbsp;Total Documento&nbsp;</strong></td>
        <td><strong><%=formatnumber(ttlis		,2,,true)%></strong></td>        
        <td><strong><%=formatnumber(ttdct   	,2,,true)%></strong></td>
        <td></td>
        <td><strong><%=formatnumber(ttigv   	,2,,true)%></strong></td>
        <td><strong><%=formatnumber(ttiSC   	,2,,true)%></strong></td>
        <td><strong><%=formatnumber(ttprecio	,2,,true)%></strong></td>
        <td><strong><%=formatnumber(TTTOT    	,2,,true)%></strong></td>
	</tr>

</table>
</center>
</body>
<%RS.CLOSE %>
<script language="jscript" type="text/jscript">
function REPORTE(op) { 
if (op == '3')
    window.open('resumendeta.asp?pos=' + '<%=pos %>' + '&tipo=' + '<%=tip %>' + '&ini=' + '<%=ini %>' + '&fin=' + '<%=fin%>'+'&EXCEL=1')
else if (op == '4')
    window.open('resumenexcelSHORT.asp?pos=' + '<%=pos %>' + '&tipo=' + '<%=tip %>' + '&ini=' + '<%=ini %>' + '&fin=' + '<%=fin%>')
else if (op == '2')
    window.open('resumenSHORT.asp?pos=' + '<%=pos %>' + '&tipo=' + '<%=tip %>' + '&ini=' + '<%=ini %>' + '&fin=' + '<%=fin%>')
return true;
    
    }
</script>
</html>
