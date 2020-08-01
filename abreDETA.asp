<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<% tienda = Request.Cookies("tienda")("pos") %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="COMUN/FUNCIONESCOMUNES.ASP"-->
<!--#include file="COMUN/COMUNqry.ASP"-->
<!--#include file="includes/Cnn.inc"-->
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>detalle</title>
</head>
<script language="jscript" type="text/jscript">
function agranda(pos) {
pos = 'trim(request.QueryString("pos") )'
if (trim(pos) == '')
    top.parent.window.document.getElementById('body0').height = 400
   
}
</script>
<body onload="agranda()">
<%pos = trim(request.QueryString("pos") )
TDA = left(pos,2)
' la primera vez que se muestra la pantalla no tiene querystring y se deb mostratr en blanco la pantalla
if len(trim(pos)) <=0 then  response.End  

tipo = right(pos,2)
pos = left(pos,4)

' ****************************************************************************
' HAY QUE VER LA FORMA DE AUMENTAR EL DELAY PARA LA BASE DE DATOS
' CON TIENDA 2, DEMORA 5 SEGS EN EL QUERY ANALIZER Y FUNCIONA
' MAS DE 6 SEGS EN EL QUERY ANALIZER Y SE CUELGA POR EXCESO DE TIEMPO ....
' ****************************************************************************
CAD = "EXEC SP_LLENA_TIENDA '"&POS&"' "

   cnn.execute CAD
 %> 
 PASA EL PRIMER SP
 <%=TIME()%>

 <%

 CAD = "EXEC SP_LLENA_TIENDA2 '"&POS&"' "

   cnn.execute CAD

%>
<br />
<%=TIME() %>
<br />
PASA EL SEGUNDO
<%   CAD = "SELECT DISTINCT GRUPO, DESCRI FROM VIEW_GRUPOS WHERE TIENDA = '"&POS&"'"
     response.write(CAD)


 RS.OPEN CAD,CNN

 If rs.recordcount<=0 then   
    RESPONSE.Write("SIN REGISTROS....")
    response.End  
 END IF      
 rs.movefirst
 
 
%>
<table align="center" cellpadding="0" cellspacing="1" bordercolor='<%=application("color2") %>' border="0" >
   <tr> <td align="center" class="Estilo8">IT</td>
        <td align="center" class="Estilo8">&nbsp;GRUPO&nbsp;</td>
	    <td align="center" class="Estilo8">DESCRIPCION</td>
   </tr>
 
    <%I=0 %>
    <%do while not rs.eof %>
        <tr bgcolor="<%IF I MOD 2 = 0 THEN RESPONSE.WRITE(application("color1")) ELSE RESPONSE.WRITE(application("color2"))%>" class="Estilo19">
            <td>&nbsp;<%=I%>&nbsp;</td>
            <td align="center" class="Estilo21"><%=trim(rs("grupo"))%></td>
            <td class="Estilo19"><%=trim(rs("descri"))%></td>
       </tr> 
       <%I= I + 1%>
       <%rs.movenext %>
    <%loop%>
</table>


</body>
<%rs.close
set rs = nothing
set cnn = nothing
 %>

</html>
