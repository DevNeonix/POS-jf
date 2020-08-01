<%@ Language=VBScript %>
<%TDA = Request.Cookies("TIENDA")("POS")%>
<%Response.Buffer = TRUE %>
<!--#include file="../includes/Cnn.inc"-->
<!--#include file="../comun/funcionescomunes.asp"-->
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
des = ""
CAD = " SELECT distinct  DESCRI " & _
      " FROM  VIEW_GRUPOS       " & _
      " WHERE grupo ='"&POS&"'  "
      response.write(cad)
      response.write("<br>")
rs.open cad, cnn
IF rs.recordcount > 0 THEN     des = rs("descri")

Cnn.Close	
set Cnn = Nothing
SET RS = Nothing	
'regresa a la página de donde fué llamado, para que vea que agregó el registro
%>
<script language="javascript" type="text/jscript">
    parent.window.document.all.descri.innerText = '<%=trim(des) %>'
</script>
</BODY>
</HTML>
