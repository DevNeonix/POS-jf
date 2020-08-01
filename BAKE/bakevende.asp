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
RESPONSE.WRITE(POS)
      response.write("<br>")
CAD = " SELECT *                  " & _
      " FROM  usuarios            " & _
      " WHERE usuario ='"&POS&"'  "
 rs.open cad, cnn

    RESPONSE.WRITE(cad)  
IF rs.recordcount > 0 THEN
    des = rs("NOMBREs")   
   RS.CLOSE%>
   <script language="javascript" type="text/jscript">
       window.parent.document.all.des.value = '<%=trim(des)%>'
       window.parent.document.getElementById('prn').style.display = 'block'
       setCookie('vende', '<%=trim(pos)%>')
 </script>
<% else%>
     <script language="javascript" type="text/jscript">
         alert("vendedor no Registrado")   
    </script>
<%END IF	
Cnn.Close	
set Cnn = Nothing
SET RS = Nothing	
'regresa a la página de donde fué llamado, para que vea que agregó el registro
%>

</BODY>
</HTML>
