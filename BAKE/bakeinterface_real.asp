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

<body style="margin-top:0">

<%
'*********************** OJO **********************
'REVISAR EL LARGO DE LAS COLUMNAS de la tabla
'*********************** OJO **********************
TDA = ucase(TRIM(Request.QueryString("pos")))
INI = ucase(TRIM(Request.QueryString("INI")))
FIN = ucase(TRIM(Request.QueryString("FIN")))

CAD =   " SET DATEFORMAT DMY; exec INTERFACE_REAL '"&TDA&"', '"&INI&"', '"&FIN&"'; "
 
 rs.open cad, cnn

' RESPONSE.Write(CAD)
IF rs.recordcount > 0 THEN
   crea = 1
ELSE
   crea = 0
END IF	

'regresa a la página de donde fué llamado, para que vea que agregó el registro
%>

<%if crea = 1 then %>
<table align="left" cellpadding="2" cellspacing="1" bordercolor='<%=application("color1") %>' border="1" id="MSG" name="MSG" >
  <tr><td colspan="9"  class="Estilo3">DOCUMENTOS GENERADOS EN EL REAL ...</td></tr>
  <tr class="Estilo0">
        <%for i=0 to rs.fieLds.count-1 %>
            <td><%=rs.fields(i).name%></td>
        <%next %>
    </tr>
  <%rs.movefirst%>
  <%do while not rs.eof %>
    <tr class="Estilo0">
        <%for i=0 to rs.fieLds.count-1 %>
            <td><%=rs.fields.item(i)%></td>
        <%next %>
    </tr>
    <%rs.movenext %>
<%loop %>
</table>
<script language="javascript" type="text/jscript">
    window.parent.document.all.espera3.style.display = 'none'
    window.parent.document.all.ok3.style.display = 'block'
    parent.window.stocks()    
</script>




<%else%>
<table align="center" cellpadding="2" cellspacing="0" border="1">
   <tr> <td align="center" class="Estilo3" colspan="3 ">No hay documentos que transferir</td>
   </tr>
</table>
<%end if%>




</BODY>
</HTML>
