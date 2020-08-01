<%@ Language=VBScript %>
<%TDA = Request.Cookies("TIENDA")("POS")%>
<%Response.Buffer = TRUE %>
<!--#include file="../includes/Cnn.inc"-->
<!--#include file="../comun/funcionescomunes.asp"-->
<link REL="stylesheet" TYPE="text/css" HREF="../ventas.CSS" >
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<body>

<%
'*********************** OJO **********************
'REVISAR EL LARGO DE LAS COLUMNAS de la tabla
'*********************** OJO **********************
pos = ucase(TRIM(Request.QueryString("pos")))
OP = ucase(TRIM(Request.QueryString("OP")))
marka = 0
    CAD =   " SELECT * FROM  VIEW_ARTICULOS_TIENDA   WHERE   " & _
            " tienda = '"&TDA&"' AND codigo like '"&POS&"%' and estado = 'A' and stock>0  " & _
			" order by codigo; "  ' Fac 20121229
           ' response.write(cad)
           ' response.write("<br>")
    RS.OPEN CAD,CNN
    if rs.recordcount > 0 then%>
    <table width="100%">
	<tr><td  align="center" class="estilo6">Artículos <%=ucase(destda) %></td></tr>
	<tr><td><hr /></td></tr>
</table>

<table id="TABLA" align="center"  cellpadding="0" cellspacing="2" bordercolor='<%=application("color2") %>' border="1" >
<%'**************************%>
<%'LINEA DE CABECERA STANDAR %>
<%'**************************%>
<%cont = 1
rs.movefirst %>
    <tr>
	    <td align="center" class="Estilo8">Codigo</td>
        <td align="center" class="Estilo8">Descripcion</td>
        <td align="center" class="Estilo8">Cant.</td>
    </tr>
    <%IF NOT RS.EOF THEN%>
        <%DO WHILE NOT RS.EOF%>
            <tr  bgcolor="<% if CONT mod 2  = 0 THEN 
                    response.write(Application("color2"))
                    else
	                response.write(Application("color1"))
	                end IF%>"
	                 id="fila<%=Trim(Cstr(cont))%>" onDblClick="selecta('<%=rs("codigo")%>', '<%=rs("descri") %>')">
	            <td class="Estilo5" align="center"><%=rs("codigo") %></td>
                <td class="Estilo5" align="center"><%=ucase(rs("descri")) %></td>
                <td class="Estilo5" align="center"><%=rs("stock") %></td>
                <%cont =cont +1 %>
            </tr>
            <%rs.movenext %>
        <%loop%>
    <%end if %>
</table>
<%END IF%>
<script language="jscript" type="text/jscript">
function selecta(pos, des) {
    po = parseInt('<%=op%>', 10)
    //alert(op)
    eval("window.opener.document.all.COD" + po + ".value=pos")
    eval("window.opener.document.all.DES" + po + ".value=des")
    window.opener.carga('<%=op%>')
    this.window.close();
}
    

    </script>
</BODY>
</HTML>
