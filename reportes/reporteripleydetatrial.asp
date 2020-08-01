<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<!--#include file="../includes/cnn.inc"-->
<%
	fechai 	= request.queryString("ini")
	fechaf 	= request.queryString("fin")
	excel 	= request.queryString("excel")
	cad 	= "exec sp_reporte_ventas_trial '"&fechai&"','"&fechaf&"'"
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
				<td style="font-family: sans-serif;color:#fff">EMP</td>
				<td style="font-family: sans-serif;color:#fff">TIPDOC</td>
				<td style="font-family: sans-serif;color:#fff">NUMDOC</td>
				<td style="font-family: sans-serif;color:#fff">CODIGO</td>
				<td style="font-family: sans-serif;color:#fff">DESCRIPCION</td>
				<td style="font-family: sans-serif;color:#fff">PRECIO</td>
				<td style="font-family: sans-serif;color:#fff">CANTIDAD</td>
				<td style="font-family: sans-serif;color:#fff">SUBTOTAL</td>
				<td style="font-family: sans-serif;color:#fff">FECH DOC</td>
				
			</tr>
		</thead>
		<tbody>
			<%for i=0 to rs.recordcount-1 %>
			<tr>
				<td style="color:#333;text-align: left"><%=rs("emp")%></td>
				<td style="color:#333;text-align: left"><%=rs("f6_ctd")%></td>
				<td style="color:#333;text-align: left"><%=rs("numdoc")%></td>
				<td style="color:#333;text-align: left">&nbsp;<%=rs("f6_ccodigo")%></td>
				<td style="color:#333;text-align: left"><%=rs("f6_cdescri")%></td>
				<td style="color:#333;text-align: left"><%=rs("F6_NPRECIO")%></td>
				<td style="color:#333;text-align: left"><%=rs("F6_Ncantid")%></td>
				<td style="color:#333;text-align: left"><%=rs("F6_NIMPMN")%></td>
				<td style="color:#333"><%=rs("f6_dfecdoc")%></td>
				
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