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
top.window.document.getElementById('body0').height = 425  
</script>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title> 
</head>
<body >
<table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="1"  cellspacing="2"  border="0" align="center" >
    <tr valign="middle"><td class="Estilo1">Número :</td><td class="Estilo1">Motivo</td><td class="Estilo1">Destino</td></tr>
    <tr valign="middle" >
    <% cad = "select serie, correl from documento where cia = '"&tienda&"' and codigo = 'TR'"
        rs.open  cad, cnn
        rs.movefirst%>
        <td><input  type="text" class="Estilo0" name="DOC" id="DOC" readonly tabindex="-1" value ="<%=rs("serie")%>-<%=right("0000000"+cstr(cdbl(rs("correl"))+1),7)%>"></td>
        <td class="Estilo12" align="left"  rowspan="2">
        <select id="TIP" name="TIP" class="Estilo12">
        <%rs.close
        CAD = "select distinct codigo, descripcion from MOTITRANS where estado ='a' AND CODIGO < '50'order by 2" 
        RS.OPEN CAD,CNN%>
        <%IF RS.RECORDCOUNT <= 0 THEN%>
            <option value="">no Hay Motivos definidos</option>
        <%END IF%>
        <%rs.movefirst
        do while not rs.eof%>
            <option value="<%=TRIM(RS("CODIGO"))%>"><%=TRIM(RS("DESCRIPCION")) %></option>
            <%rs.movenext %>
        <%loop%>
        <%rs.close%>
        </select>
        </td>
        
        <td class="Estilo12" align="left"  rowspan="2" width="150px"><input  name="NRO" id="NRO" onkeyup="this.value=toAlpha(this.value)" maxlength="50" size="50"></td>
        
         <td><img src="images/ok.gif" onclick="manda()" style="cursor:pointer;"/></td>            
    </tr> 
</table>

</body>

<script type="text/jscript" language="jscript">
window.document.all.TIP.selectedIndex= -1;
window.document.all.TIP.focus();
function manda() {
    if (trim(window.document.all.TIP.value) == "") {
        alert('Debe Informar el tipo de Transferencia');
        window.document.all.TIP.focus();
        return false;
    }

    if (trim(window.document.all.NRO.value) == "") {
        alert('Debe Informar el Destinatario');
        window.document.all.NRO.focus();
        return false;
    }
    parent.window.frames[1].window.location.replace("transferdeta.asp")
 return true
}

</script>

</html>
