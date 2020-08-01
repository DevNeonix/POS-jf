<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.Buffer = FALSE %>
<%Session.LCID=2058%>
<!--#include file="../includes/cnn.inc"-->
<%
	fechai 	= request.queryString("ini")
	fechaf 	= request.queryString("fin")
	excel 	= request.queryString("excel")
	agru 	= request.queryString("agru")
	cad 	= "exec sp_reporte_ventas_ripley '"&fechai&"','"&fechaf&"'"
	if agru="true" then
		cad 	= "exec sp_reporte_ventas_ripley_agrupado '"&fechai&"','"&fechaf&"'"
	end if
	'response.write(cad)
	rs.open cad,cnn

	if excel = "1" then
		archivo = "c:\temp\VTS_DETA_ripley_excel.xls"
  		Response.Charset = "UTF-8"
    	Response.ContentType = "application/vnd.ms-excel" 
   		Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo 
	end if

%>

<%
if rs.recordcount > 0 then
	rs.movefirst
%>

	<table align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="2"  cellspacing="4"  border="0">
		<thead>
			<tr style="background:#003366 ">
				<td style="font-family: sans-serif;color:#fff">SKU</td>
				<td style="font-family: sans-serif;color:#fff">CÓDIGO</td>
				<td style="font-family: sans-serif;color:#fff">CODCOL</td>
				<td style="font-family: sans-serif;color:#fff">DESCOL</td>
				<td style="font-family: sans-serif;color:#fff">CODTAL</td>
				<td style="font-family: sans-serif;color:#fff">DESTAL</td>
				<%if agru = "false" then%><td style="font-family: sans-serif;color:#fff">MES</td><%end if%>
				<%if agru = "false" then%><td style="font-family: sans-serif;color:#fff">AÑO</td><%end if%>
				<td style="font-family: sans-serif;color:#fff">DESCRIPCIÓN</td>
				<td style="font-family: sans-serif;color:#fff">CANTIDAD</td>
				<td style="font-family: sans-serif;color:#fff">COSTO UNIT</td>
				<td style="font-family: sans-serif;color:#fff">COSTO TOT</td>
				<td style="font-family: sans-serif;color:#fff">PRECIO UNIT</td>
				<td style="font-family: sans-serif;color:#fff">PRECIO TOT</td>
				
				
			</tr>
		</thead>
		<tbody>
			<%for i=0 to rs.recordcount-1 %>
			<tr>
				<td style="color:#333;text-align: left">&nbsp;<%=rs("f6_ccodigo")%></td>
				<td style="color:#333;text-align: left">&nbsp;<%=LEFT(TRIM(rs("f6_ccodigo")),5)%></td>
				<td style="color:#333;text-align: left"><%=rs("CODCOL")%></td>
				<td style="color:#333;text-align: left"><%=rs("DESCOL")%></td>
				<td style="color:#333;text-align: left"><%=rs("CODTAL")%></td>
				<td style="color:#333;text-align: left"><%=rs("DESTAL")%></td>
				<%if agru = "false" then%><td style="color:#333;text-align: left"><%=rs("MES")%></td><%end if%>
				<%if agru = "false" then%><td style="color:#333;text-align: left"><%=rs("ANO")%></td><%end if%>
				<%if agru="true" then%>
				<td style="color:#333;text-align: left"><%=rs("ar_cdescri")%></td>
				<td style="color:#333"><%=rs("vendidos")%></td>
				<td style="color:#333"><%=rs("costo")%></td>
				<td style="color:#333"><%=CDBL(rs("costo")) * CDBL(rs("vendidos"))%></td>
				<td style="color:#333"><%=CDBL(rs("precio"))%></td>
				<td style="color:#333"><%=CDBL(rs("precio")) * CDBL(rs("vendidos"))%></td>
				<td style="color:#333"><%=cdbl(   CDBL(rs("costo")) /  CDBL(rs("precio")) * 100 )%>%</td>
				<%else %>
				<td style="color:#333;text-align: left"><%=rs("f6_cdescri")%></td>
				<td style="color:#333"><%=rs("vendidos")%></td>
				<td style="color:#333"><%=rs("costo")%></td>
				<td style="color:#333"><%=CDBL(rs("costo")) * CDBL(rs("vendidos"))%></td>
				<td style="color:#333"><%=CDBL(rs("f6_nprecio"))%></td>
				<td style="color:#333"><%=CDBL(rs("f6_nprecio")) * CDBL(rs("vendidos"))%></td>
				<td style="color:#333"><%=cdbl(   CDBL(rs("costo")) /  CDBL(rs("f6_nprecio")) * 100 )%>%</td>
				<%end if%>
				
				
			</tr>
			<%
			rs.movenext
			next
			%>
		</tbody>
	</table>

<%

else
response.write("no se encontraron registros con esa fecha")
end if%>