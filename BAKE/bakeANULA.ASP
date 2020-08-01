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
ser = ucase(TRIM(Request.QueryString("ser")))
NRO = ucase(TRIM(Request.QueryString("NRO")))
TIP = ucase(TRIM(Request.QueryString("TIP")))
      response.write("<br>")
CAD = " SELECT OPERACION                      " & _
      " FROM  movimcab                        " & _
      " WHERE TIENDA = '"&TDA&"' AND          " & _
      " SERIE ='"&ser&"' and                  " & _
      " NUMDOC = '"&NRO&"'                    " & _
      " AND CODDOC = '"&TIP&"'  order by 1 desc;               "
 rs.open cad, cnn

    RESPONSE.WRITE(cad)  
IF rs.recordcount <= 0 THEN%>

   <script language="javascript" type="text/jscript">
       alert("Documento NO EXISTE para esta tienda")
       window.parent.document.all.SER.value = ''
       window.parent.document.all.NRO.value = ''
       window.parent.document.all.TIP.selectedIndex = -1
 </script>
 <%else %>
 <script language="javascript" type="text/jscript">
 
    top.window.frames[0].window.frames[1].window.location.replace('../anuladeta.asp?pos='+'<%=rs("operacion") %>')
     
 </script>
<%END IF	
rs.close
Cnn.Close	
set Cnn = Nothing
SET RS = Nothing	
'regresa a la página de donde fué llamado, para que vea que agregó el registro
%>
 


</BODY>
</HTML>
