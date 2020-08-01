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
   top.parent.window.document.getElementById('body0').height = 400 
}
</script>
<body onload="agranda();document.all.COD.focus()">

<hr />
<table id="Table2" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="1"  cellspacing="2"  border="0" align="center" >
    <tr valign="middle">
    	<td class="Estilo1">Doc Origen</td>
        <td class="Estilo1">Serie</td>
        <td class="Estilo1">Numero</td>
       
    <tr valign="middle" >
        <td class="Estilo12" align="left"  rowspan="2">
        <select id="DOC" name="DOC" class="Estilo12">
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
        <td class="Estilo12" align="left"  rowspan="2"><input  name="SER" id="SER" size="5"  onkeyup="this.value=toInt(this.value)" maxlength="3"></td>
        <td class="Estilo12" align="left"  rowspan="2"><input  name="NRO" id="NRO" onKeyUp="this.value=toInt(this.value)" maxlength="7" size="10"></td>
      
    </tr>
   
</table>

<table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="3"  cellspacing="1"  border="1" align="center" >
   <tr>
     
	    <td align="center" class="Estilo8">Codigo</td>
        <td align="center" class="Estilo8">Concepto</td>
        <td align="center" class="Estilo8">Cant.</td>
        <td align="center" class="Estilo8">S/.</td>
   </tr>

    <tr id="lin" name="lin" >
     
        <td><input id="COD" name="COD" size="20" class="Estilo13" readonly tabindex="-1" value='ESPECIAL'  /></td>
        <td><input id="DES" name="DES" size="60" class="Estilo12" value=''/></td>
        <td><input id="CAN" name="CAN" size="8"  class="Estilo13" value='1' readonly tabindex="-1"/></td>
        <td><input id="SOL" name="SOL" size="10" class="Estilo12"   /></td>
    </tr>

</table>


<center><img src="images/print.jpg" onclick="graba()" style="cursor:pointer;" id="pr" name="pr"/></center>  

<iframe width="100%" src="" id="mirada" name="mirada"  scrolling="yes" frameborder="3" height="100" align="middle" style="display:none">
</iframe>
</body>


<script type="text/jscript" language="jscript">


function graba() {

    if (trim(document.all.SER.value) == '') {
        alert("Favor informar la Serie")
        return false;
    }

    if (trim(document.all.NRO.value) == '') {
        alert("Favor informar el nro dek¡l documento")
        return false;
    }
    else
    document.all.NRO.value = strzero(document.all.NRO.value,7)
    if (trim(document.all.DES.value) == '') {
        alert("Favor informar el concepto")
        return false;
    }
    if (trim(document.all.SOL.value) == '') {
        alert("Favor informar el valor")
        return false;
    }

    // Datos para grabar cabecera y detalle
        cad  = 'comun/NOTAESPACIAL.asp?cli=' + trim(parent.window.frames[0].window.document.all.CLI.value)
        cad += '&cod=ESPECIAL'
        cad += '&Can=1'
        cad += '&doc=NC'
        cad += '&mov=E'
        cad += '&ser=' +  Left(ltrim(rtrim(parent.window.frames[0].window.document.all.DOC.value)), 3)
        cad += '&nro=' + Right(ltrim(rtrim(parent.window.frames[0].window.document.all.DOC.value)), 7)
        cad += '&DES=' + ltrim(rtrim(document.all.DES.value))
        cad += '&sol=' + ltrim(rtrim(document.all.SOL.value))
        cad += '&ser=' + ltrim(rtrim(document.all.SER.value))
        cad += '&NRO=' + ltrim(rtrim(document.all.NRO.value))
        var opc = "directories=no,height=600,";
        opc += "hotkeys=no,location=no,";
        opc += "menubar=no,resizable=no,";
        opc += "left=0,top=0,scrollbars=yes,";
        opc += "status=no,titlebar=no,toolbar=no,";
        opc += "width=800";
        document.all.pr.style.display='none'
        //alert(cad)
        window.open(cad, '', opc)



    }
</script>
</html>
