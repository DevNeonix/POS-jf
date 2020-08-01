<%@LANGUAGE="VBSCRIPT" %>
<%Response.Buffer = true %>
<%Session.LCID=2058%>

<!--#include file="../includes/Cnn.inc"-->
<!--#include file="../comun/funcionescomunes.asp"-->
<link REL="stylesheet" TYPE="text/css" HREF="../ventas.CSS" >
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>

<%
IF  request.QueryString("EXCEL") = "1" THEN
  archivo = "c:\temp\cajaexcel.xls"
    Response.Charset = "UTF-8"
    Response.ContentType = "application/vnd.ms-excel" 
    Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo 
END IF

%>

<body>
<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="Table2" name="listado"  >
<tr>
<td><input type="button" value="Excel " onclick="REPORTE(1)" /></td>
</tr>
</table>
<center>
<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="listado" name="listado"  >
	<tr> 
   		<td align="center" class="Estilo8">Tienda</td>
        <td align="center" class="Estilo8">Año</td>
        <td align="center" class="Estilo8">ENE</td>
        <td align="center" class="Estilo8">FEB</td>
        <td align="center" class="Estilo8">MAR</td>
        <td align="center" class="Estilo8">ABR</td>
        <td align="center" class="Estilo8">MAY</td>
        <td align="center" class="Estilo8">JUN</td>
        <td align="center" class="Estilo8">JUL</td>
        <td align="center" class="Estilo8">AGO</td>
        <td align="center" class="Estilo8">SET</td>
        <td align="center" class="Estilo8">OCT</td>
        <td align="center" class="Estilo8">NOV</td>
        <td align="center" class="Estilo8">DIC</td>
	</tr>
<% cad =  " SELECT   *   FROM  ( SELECT  TIENDA, ANO AS [AÑO], VENTA, MES  FROM jacinta..historia  ) AS T            " & _
          " PIVOT (  SUM(Venta)  FOR mes IN ([01],[02],[03],[04], [05],[06],[07],[08],[09],[10],[11],[12] ) " & _
          " ) AS P order by  tienda, ano desc                                                               "

rs.open cad,cnn

if rs.recordcount <=0 then RESPONSE.End %>

	<%do while not rS.eof %>
		<tr bgcolor='<%=application("color2") %>'  class="EstiloT" align="right">
        	<td>&nbsp;<%=(rs("TIENDA"))%>&nbsp;</td>
            <td>&nbsp;<%=rs("AÑO")%>&nbsp;</td>
            <td style="padding-right:10px"><%if isnull(rs("01")) then response.write("0.00") else response.write(FORMATNUMBER(rs("01"), 2,,,TRUE))%></td>
			<td style="padding-right:10px"><%if isnull(rs("02")) then response.write("0.00") else response.write(FORMATNUMBER(rs("02"), 2,,,TRUE))%></td>
			<td style="padding-right:10px"><%if isnull(rs("03")) then response.write("0.00") else response.write(FORMATNUMBER(rs("03"), 2,,,TRUE))%></td>
			<td style="padding-right:10px"><%if isnull(rs("04")) then response.write("0.00") else response.write(FORMATNUMBER(RS("04"), 2,,,TRUE))%></td>
			<td style="padding-right:10px"><%if isnull(rs("05")) then response.write("0.00") else response.write(FORMATNUMBER(RS("05"), 2,,,true))%></td>
            <td style="padding-right:10px"><%if isnull(rs("06")) then response.write("0.00") else response.write(FORMATNUMBER(RS("06"), 2,,,true))%></td>
            <td style="padding-right:10px"><%if isnull(rs("07")) then response.write("0.00") else response.write(FORMATNUMBER(RS("07"), 2,,,true))%></td>
            <td style="padding-right:10px"><%if isnull(rs("08")) then response.write("0.00") else response.write(FORMATNUMBER(RS("08"), 2,,,true))%></td>
            <td style="padding-right:10px"><%if isnull(rs("09")) then response.write("0.00") else response.write(FORMATNUMBER(RS("09"), 2,,,true))%></td>
            <td style="padding-right:10px"><%if isnull(rs("10")) then response.write("0.00") else response.write(FORMATNUMBER(RS("10"), 2,,,true))%></td>
            <td style="padding-right:10px"><%if isnull(rs("11")) then response.write("0.00") else response.write(FORMATNUMBER(RS("11"), 2,,,true))%></td>
            <td style="padding-right:10px"><%if isnull(rs("12")) then response.write("0.00") else response.write(FORMATNUMBER(RS("12"), 2,,,true))%></td>
		</tr>                                                                                                                               
        <%RS.movenext%>
	<%loop %>

</table>

</center>
</body>
<script language="jscript" type="text/jscript">
    function REPORTE(op) {
        if (op == '1')
            window.open('comparativo.asp?EXCEL=1')
    }
</script>
</html>
