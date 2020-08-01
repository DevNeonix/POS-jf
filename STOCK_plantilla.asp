<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%session.LCID = 2057 %>
<% tienda = Request.Cookies("tienda")("pos") %>
<% destda = Request.Cookies("tienda")("tda") %>
<!--#include file="includes/Cnn.inc"-->
<!--#include file="comun/funcionescomunes.asp"-->
<%
lin= request.querystring("lin")
CAD =	" select DISTINCT CODIGO, DESCRI, STOCK, MINIMO, PLANILLA AS PLA " & _
        " from  view_ARTICULOS_tienda           " & _
        " where  planilla ='1'                  " & _
        " and tienda = '"&TIENDA&"'             " & _
		" order by codigo "	'Fac 20121229
'RESPONSE.WRITE(CAD)
RS.OPEN CAD ,Cnn
		
If rs.eof or rs.bof then
	Response.Write("Stock OK, no se requiere Reposición")		
	Response.End
end if	
rs.movefirst
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=titulo%></title>
<link rel="stylesheet" type="text/css" href="ventas.CSS" />

</head>

<body topmargin="0" leftmargin="0" rightmargin="0" border="0">

<table width="100%">
	<tr><td  align="center" class="estilo6">Artículos Plantilla --> tienda: <%=ucase(destda) %></td></tr>
	<tr><td><hr /></td></tr>
</table>

<table id="TABLA" align="center"  cellpadding="2" cellspacing="2" bordercolor='<%=application("color2") %>' border="0" >
<%'**************************%>
<%'LINEA DE CABECERA STANDAR %>
<%'**************************%>
<%cont = 1%>
<tr>
    <td align="center" class="Estilo8">IT</td>
	<td align="center" class="Estilo8">Codigo</td>
    <td align="center" class="Estilo8">Descripcion</td>
    <td align="center" class="Estilo8">Cant.</td>
    <td align="center" class="Estilo8">Min.</td>
    <td align="center" class="Estilo8">Pla</td>
</tr>
<%IF NOT RS.EOF THEN%>
    <%DO WHILE NOT RS.EOF%>
        <tr  bgcolor="<% if CONT mod 2  = 0 THEN 
                response.write(Application("color2"))
                else
	            response.write(Application("color1"))
	            end IF%>" id="fila<%=Trim(Cstr(cont))%>" name="fila<%=Trim(Cstr(cont))%>" 
            class="<%if cdbl(rs("stock")) < cdbl(rs("minimo")) then 
                    response.write("Estilo23") 
                 else 
                    response.write("estilo5") 
                 end if %>">
            <td  align="center"><%=Cont%></td>
	        <td  align="center"><%=rs("codigo") %></td>
            <td align="left"><%=ucase(rs("descri")) %></td>
            <td align="center"><%=rs("stock") %></td>
            <td  align="center"><%=rs("minimo") %></td>

            <td  align="center">
                <select id="pl<%=cont%>" name="pl<%=cont%>" onchange="graba('<%=rs("codigo") %>','<%=cont%>', this.value)">
                    <option value="0" <%if rs("pla") = "0" then%> selected <%end if%>>No</option>
                    <option value="1" <%if rs("pla") = "1" then%> selected <%end if%>>Si</option>
                </select>
                
            </td>
            <%cont =cont +1 %>
        </tr>
        <%rs.movenext %>
    <%loop%>
<%end if %>
</table>


<iframe frameborder="1" style="visibility:hidden" height="100" width="100%" id="ACTIV" name="ACTIV"></iframe>
 <script type="text/jscript" language="jscript">

function graba(cod, ff,op) {
    var pos = parseInt(ff, 10)
    cad  = "comun/updateplantilla.asp?lin=" + pos + '&cod=' + trim(cod)
    cad += '&pla=' + op
    document.all.ACTIV.src=cad
    

}
 </script>
</body>

</html>
