<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = true %>
<%Session.LCID=2058%>

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
fec = request.QueryString("fec")
fec2 = request.QueryString("fec2")

 cad = "exec SP_ventas_DIA  '"&fec&"', '"&fec2&"' "
'RESPONSE.WRITE (CAD)
rs.open cad,cnn
if rs.recordcount <=0 then RESPONSE.End
IF  request.QueryString("exl") = "1" THEN
  archivo = "c:\temp\cajaexcel.xls"
    Response.Charset = "UTF-8"
    Response.ContentType = "application/vnd.ms-excel" 
    Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo 
END IF

%>

<body onload="AGRANDA()">
<center>
<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="listado" name="listado"  >
	<tr> 
    <%for i = 0 to rs.fields.count-1%>
   		<td align="center" class="Estilo8"><%=RS.FIELDS(I).name %></td>
    <%NEXT%>        
	</tr>
	
	<%do while not rs.eof%>
		<tr bgcolor='<%=application("color2") %>'  class="EstiloT" align="right">
            <td  align="center" class="EstiloT"><%=rs.fields.item(0)%></td>
            <%for i = 1 to rs.fields.count-1%>
   		        <td align="right" class="EstiloT"><%=formatnumber(cdbl(RS.FIELDS.Item(i)),2,,,true) %></td>
            <%NEXT%>        
        </tr>
        <%rs.movenext%>        	
    <%LOOP%>
</table>



</body>

</html>
