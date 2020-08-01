<%@Language=VBScript %>
<%Response.Buffer = true%>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos")%>
<%tda = Request.Cookies("tienda")("tda")%>
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

<% 
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body >

<iframe src="" id="Iframe1" name="mirada" style="display:none" width="100%" height="20"></iframe>

<table width="100%" >
	<tr><td colspan="3"><hr /></td></tr>
    <tr><td align="center" class="Estilo6">Ingrese el número de Documento de Venta Para NC CONTINGENCIA</td></tr>
	
</table>
<table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="1"  cellspacing="2"  border="0" align="center" >
    <tr valign="middle">
    	<td class="Estilo1">Doc</td>
        <td class="Estilo1">Serie</td>
        <td class="Estilo1">Numero</td>
        <td class="Estilo1" colspan="2">Producto</td>
        <td>&nbsp;</td></tr>
    <tr valign="middle" >
        <td class="Estilo12" align="left"  rowspan="2">
        <select id="TIP" name="TIP" class="Estilo12">
        <%CAD = "select distinct codigo, descripcion from documento order by 2" 
        RS.OPEN CAD,CNN%>
        <%IF RS.RECORDCOUNT <= 0 THEN%>
            <option value="">no Hay Documentos definidos</option>
        <%END IF%>
        <%rs.movefirst
        do while not rs.eof%>
            <option value="<%=TRIM(RS("CODIGO"))%>"><%=TRIM(RS("DESCRIPCION")) %></option>
            <%rs.movenext %>
        <%loop%>
        </select>
        </td>
        <td class="Estilo12" align="left"  rowspan="2"><input  name="SER" id="SER" size="5"  onkeyup="this.value=(this.value)" maxlength="4"></td>
        <td class="Estilo12" align="left"  rowspan="2"><input  name="NRO" id="NRO" onKeyUp="this.value=toInt(this.value)" maxlength="7" size="10"></td>
        <td class="Estilo12" align="left"  rowspan="2"><input  name="COD" id="COD" size="12"></td>
        <td><input  name="DES" id="DES" class="Estilo1"  readonly tabindex="-1" size="30"></td>
         <td><img src="images/ok.gif" onClick="manda()" style="cursor:pointer;"/></td>            
    </tr>
   
</table>
</body>

<script type="text/jscript" language="jscript">
function manda() {
    if (trim(window.document.all.TIP.value) == "") {
        alert('Debe Informar el tipo de Documento');
        window.document.all.TIP.focus();
        return false;
    }

    if (trim(window.document.all.SER.value) == "") {
        alert('Debe Informar la Serie del documento');
        window.document.all.SER.focus();
        return false;
    }
    if (trim(window.document.all.NRO.value) == "") {
        alert('Debe Informar el número del Documento');
        window.document.all.NRO.focus();
        return false;
    }
    if (trim(window.document.all.COD.value) == "") {
        alert('Debe Informar el Código del Producto');
        window.document.all.COD.focus();
        return false;
    }

    document.all.SER.value = strzero(trim(document.all.SER.value), 3)
    document.all.NRO.value = strzero(trim(document.all.NRO.value), 7)
    cad = 'bake/bakedevuelveconti.asp?ser=' + strzero(trim(document.all.SER.value), 3) + '&nro=' + strzero(trim(document.all.NRO.value), 7) + '&cod=' + trim(document.all.COD.value)
    cad += '&tip=' + trim(document.all.TIP.value)
    //document.all.mirada.style.display='block'
    document.all.mirada.src=cad
    //alert(cad)
    //window.open("abredeta.asp?pos=" + cad)
    //parent.window.frames[1].window.location.replace = "abredeta.asp?pos=" + cad
}

</script>

</html>
