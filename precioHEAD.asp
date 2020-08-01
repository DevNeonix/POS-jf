<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<% tienda = Request.Cookies("tienda")("pos") %>
<% tda = Request.Cookies("tienda")("tda") %>
<!--#include file="includes/Cnn.inc"-->
<!--#include file="COMUN/FUNCIONESCOMUNES.ASP"-->
<!--#include file="COMUN/COMUNqry.ASP"-->
<script type="text/jscript" language="jscript">
function calcHeight()
{
  //find the height of the internal page
  var the_height=
    document.getElementById('mirada').contentWindow.
      document.body.scrollHeight;

  //change the height of the iframe
  document.getElementById('mirada').height=
      the_height+20;
}

</script>
<% CAD =   " SELECT * FROM TIENDAS WHERE ESTADO = 'A' order by descripcion  "
          '  response.write(cad)
          '  response.write("<br>")
    RS.OPEN CAD,CNN
    IF rs.recordcount > 0 THEN rs.movefirst
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
    
</head>

<body >

<table width="100%">
	<tr><td colspan="3"><hr /></td></tr>
    <tr><td align="center" class="Estilo6">Listas de Precio</td></tr>
	
</table>
<table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="2"  cellspacing="4"  border="0" align="center" >
    <tr valign="middle" >
        <td class="Estilo11" valign="middle" align="right" rowspan="2">
            <label for="Radio">Tiendas:&nbsp;</label></td> 
        <td  class="Estilo12" align="left"  rowspan="2">
            <select  name="miRadio" id="miRadio">
                <option value = "TT" selected>TODAS</option>
                <%do while not rs.eof %>
                    <option value="<%=TRIM(RS("CODIGO"))%>"><%=TRIM(RS("DESCRIPCION")) %></option>
                    <%rs.movenext %>
                <%loop %>
            </select>
            
        </td>
        <td width="15px;"></td>
        <td class="Estilo11" valign="middle" align="right" rowspan="2">
            <label for="Radio">Rango:&nbsp;</label></td> 
        <td  class="Estilo12" align="left"  rowspan="2">
            <input type="Radio" name="miRadio1" id="miRadio1" value="PG"  checked>Por Grupo
            <input type="Radio" name="miRadio1" id="miRadio1" value="PP"  >Por Producto
        </td>
        <td width="15px;"></td>
        <td class="Estilo11" valign="middle" align="right" rowspan="2">
            <label for="Radio">Tipo:&nbsp;</label></td> 
        <td  class="Estilo12" align="left"  rowspan="2">
            <input type="Radio" name="Radio2" id="Radio2" value="TL" checked>Listado
            <input type="Radio" name="Radio2" id="Radio2" value="TI">Individual
            
        </td>


         <td><img src="images/ok.gif" onclick="strSQL()" style="cursor:pointer;"/></td>            
    </tr>
   
</table>
<iframe src="" id="mirada" name="mirada" style="display:none" width="100%"></iframe>

</body>
<script type="text/jscript" language="jscript">
function strSQL() {
    cad = document.all.miRadio.value
    
    if (document.all.miRadio1[0].checked == true)
        cad += document.all.miRadio1[0].value
    if (document.all.miRadio1[1].checked == true)
        cad += document.all.miRadio1[1].value

    if (document.all.Radio2[0].checked == true)
        cad += document.all.Radio2[0].value
    if (document.all.Radio2[1].checked == true)
        cad += document.all.Radio2[1].value
    //parent.window.frames[1].window.document.all.Text1.value = cad
    parent.window.frames[1].window.location.replace = "preciodeta.asp?pos=" + cad
}

</script>

</html>
