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

<%pos = trim(request.QueryString("pos") )%>

<script language="jscript" type="text/jscript">
function agranda(pos) {
pos = 'trim(request.QueryString("pos") )'
if (trim(pos) == '')
    top.parent.window.document.getElementById('body0').height = 400
   
}
</script>
<body onload="agranda()">
  <%
 cad =  " select distinct m1.ITEM AS IT, V1.CODIGO, V1.DESCRI, M1.PRECIO, cant = case       " & _
        " WHEN M1.SALE >0 THEN M1.SALE WHEN M1.ENTRA > 0 THEN M1.ENTRA ELSE M1.VALE END     " & _
        " , v1.tienda  as TDA, operacion as OPE                                             " & _
        " from movimdet  m1 inner join view_articulos_tienda v1 on v1.codigo = m1.codart    " & _
        " and m1.tienda = v1.tienda                                                         " & _
        " where operacion ='"&pos&"'                                                        " & _
        " order by 1                                                                        "

' response.Write(cad)
 
    RS.OPEN CAD,CNN

 If rs.recordcount<=0 then   
    RESPONSE.Write("SIN REGISTROS....")
    response.End  
 END IF      
 rs.movefirst
 %>
<table align="center" cellpadding="2" cellspacing="1" bordercolor='<%=application("color2") %>' border="0" >
   <tr> 
        <%for j=0 to rs.fields.count-1 %>
            <td class="Estilo8">&nbsp;<%=RS.FIELDS(j).NAME%>&nbsp;</td>
        <%next %>
         
   </tr>
    <%I=1 %>
    <%do while not rs.eof %>
        <tr bgcolor="<%IF I MOD 2 = 0 THEN RESPONSE.WRITE(application("color1")) ELSE RESPONSE.WRITE(application("color2"))%>" class="Estilo19">
            <%for j=0 to rs.fields.count-3 %>
            <td id='COL<%=i%><%=j%>' name='COL<%=i%><%=j%>'  ><%=RS.FIELDs.item(j)%></td>
            <%next %>
            <%for h=j to rs.fields.count-1 %>
            <td align="right" id='COL<%=i%><%=h%>' name='COL<%=i%><%=h%>'><%=trim(RS.FIELDs.item(h))%>&nbsp;&nbsp;</td>
            <%next %>
           
       </tr> 
       <%I= I + 1%>
       <%rs.movenext %>
    <%loop%>
    <tr>
    <td colspan="<%=rs.fields.count%>" align="center"><input type="button" id="manda" name="manda" onclick="manda()" value="ACEPTAR"/></td>
    </tr>
</table>
<iframe width="100%"  src="" id="cok" name="cok"  height="100" style="display:none;" align="middle">
</iframe>

</body>
<%rs.close
set rs = nothing
set cnn = nothing
 %>
 <script language=jscript type="text/jscript">
function manda() {
    var urlElimina = './ppbaja.asp?ope=<%=pos%>';
    alert("Eliminando de la Sunat NO CERRAR NINGUNA VENTANA");
    window.open(urlElimina);
    setTimeout(function(){
        window.location.replace('comun/anulacion.asp?pos=' + '<%=pos %>')   
    }, 5000);

}   
 
 </script>
</html>
