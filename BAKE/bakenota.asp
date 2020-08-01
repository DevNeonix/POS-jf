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
ser = right(ucase(TRIM(Request.QueryString("ser"))),4)
ncr = right("0000000000000"+ucase(TRIM(Request.QueryString("ncr"))),7)
NOTA = SER + "-" + NCR
      response.write("<br>")
CAD = " SELECT NOTA       " & _
      " FROM  CAJA " & _
      " WHERE LTRIM(RTRIM(NOTA)) ='"&NOTA&"' "
 RESPONSE.WRITE(cad)  
 rs.open cad, cnn

    
IF rs.recordcount > 0 THEN
    
   RS.CLOSE%>
   <script language="javascript" type="text/jscript">
       alert("Documento ya aplicado a otra venta")
       window.parent.document.all.sr1.value = ''
       window.parent.document.all.ncr.value = ''
 </script>
 
<%END IF	
Cnn.Close	
set Cnn = Nothing
SET RS = Nothing	
'regresa a la página de donde fué llamado, para que vea que agregó el registro
%>
 


</BODY>
</HTML>
