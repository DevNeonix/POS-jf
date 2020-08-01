<%@ Language=VBScript %>
<%Response.Buffer = true %>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos") %>
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>
<body>
<table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="0"  cellspacing="1"  border="1" align="center" style="display:none">
   <tr>
        <td align="center" class="Estilo8"># Pdas.</td>
        <td align="center" class="Estilo8">P.Bruto</td>
        <td align="center" class="Estilo8">Dcto.</td>
        <td align="center" class="Estilo8" id="st">Sub Total</td>
        <td align="center" class="Estilo8" id="ig" >IGV</td>
        <td align="center" class="Estilo8" id="ic" >ISC</td>
        <td align="center" class="Estilo8">Total</td>
        <td align="center" rowspan="2" valign="middle" style="border:0">&nbsp;<img src="images/ok.gif" border="0" onclick="pagar()">&nbsp;</td>
   </tr>
   <tr>
        <td>         <input id="canti"  name="canti"     readonly tabindex="-1" size="8"  class="Estilo14" /></td>
        <td>         <input id="pvp"    name="pvp"       readonly tabindex="-1" size="8"  class="Estilo14" /></td>
        <td>         <input id="dcto"   name="dcto"      readonly tabindex="-1" size="8"  class="Estilo14 "/></td>
        <td id="sst"><input id="bruto"  name="bruto"     readonly tabindex="-1" size="8"  class="Estilo14" /></td>
        <td id="iig"><input id="igv"    name="igv"       readonly tabindex="-1" size="8"  class="Estilo14" /></td>
        <td id="iis"><input id="isc"    name="isc"       readonly tabindex="-1" size="8"  class="Estilo14" /></td>
        <td>         <input id="tota"   name="tota"      readonly tabindex="-1" size="8"  class="Estilo14" /></td>
   </tr>
   </table>

<table id="Table2" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="0"  cellspacing="1"  border="1" align="center" style = "display:block">
   <tr>
        <td align="center" class="Estilo8"># Pdas.</td>
        <td align="center" class="Estilo8">P.Bruto</td>
        <td align="center" class="Estilo8">Dcto.</td>
        <td align="center" class="Estilo8">ISC</td>
        <td align="center" class="Estilo8">Total</td>
        <td align="center" rowspan="2" valign="middle" style="border:0">&nbsp;<img src="images/ok.gif" border="0" onclick="pagar()">&nbsp;</td>
   </tr>
   <tr>
        <td><input id="CAN"  name="CAN" readonly tabindex="-1" size="8" class="Estilo14" /></td>
        <td><input id="PVP"  name="PVP" readonly tabindex="-1" size="8" class="Estilo14" /></td>
        <td><input id="DCT"  name="DCT" readonly tabindex="-1" size="8" class="Estilo14 "/></td>
        <td><input id="ISC"  name="ISC" readonly tabindex="-1" size="8" class="Estilo14 "/></td>
        <td><input id="TOT"  name="TOT" readonly tabindex="-1" size="8" class="Estilo14" /></td>
   </tr>
   </table>
</body>
<script language="jscript" type="text/jscript">
    var opc = "directories=no,height=500,";
    opc = opc + "hotkeys=no,location=no,";
    opc = opc + "menubar=no,resizable=no,";
    opc = opc + "left=100,top=100,scrollbars=yes,";
    opc = opc + "status=no,titlebar=no,toolbar=no,";
    opc = opc + "width=700";
function pagar() 
{   parent.window.frames[1].window.TOTALES()
    cambio = top.parent.window.document.all.cambio.value
    if (Math.round(document.all.tota.value, 2) == 0) 
    {   alert("No hay nada que facturar")
        return true
    }
    else {
        ven = 'vuelto.asp?pago=' + Math.round(document.all.tota.value * 100) / 100

        window.open(ven,"CANCELACION",opc)
    } 
    
       

}

</script>
</html>
