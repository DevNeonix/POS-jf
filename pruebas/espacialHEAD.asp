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
    <tr valign="middle"><td class="Estilo1">Número :</td>
    <% cad = "select serie, correl from documento where cia = '"&tienda&"' and codigo = 'NC'"
        rs.open  cad, cnn
        rs.movefirst%>
        <td><input  type="text" class="Estilo0" name="DOC" id="DOC" readonly tabindex="-1" value ="<%=rs("serie")%>-<%=right("0000000"+cstr(cdbl(rs("correl"))+1),7)%>"></td>
        <td class="Estilo12" align="left"  rowspan="2">
         <td class="Estilo11" valign="middle" align="right"  rowspan="2">Cliente :&nbsp;</td> 
        <td align="left"><input type="text" name="CLI" id="CLI" value="" class="Estilo12" onchange="cliente(this.value)"  size="20" maxlength="11">
        <input type="text" name="DES" id="DES" value="" maxlength="50" size="50" class="Estilo12" readonly tabindex="-1">
        <br /><input type="text" name="DIR" id="DIR" value="" maxlength="100" size="80" class="Estilo12"  readonly tabindex="-1"></td>
        <td><img src="images/ok.gif" onclick="manda()" style="cursor:pointer;"/></td>            
    </tr> 
</table>
<iframe src="" allowScriptAccess='always'  id="mirada" name="mirada" style="display:none"></iframe>

</body>

<script type="text/jscript" language="jscript">
function cliente(dato) {

        ss = document.all.CLI.value
        cad = "bake/bakecliente.asp?pos=" + trim(dato)
        //   document.all.mirada.style.display='block'
        document.all.mirada.src = cad

        if (ss.length < 8) {
            document.all.CLI.value = strzero(trim(document.all.CLI.value), 11)
            dato = strzero(dato, 11)
        }

    }
function manda() {
    if (trim(document.all.CLI.value) == "") {
        alert('Debe Informar el Cliente');
        document.all.CLI.focus();
        return false;
    }

    parent.window.frames[1].window.location.replace("espacialdeta.asp")
 return true
}

</script>

</html>
