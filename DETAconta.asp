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
function agranda() {
    //    top.parent.window.document.getElementById('body0').height = 400 
return true 
}


function calcHeight(op) {
    if (op == '0') {
        var the_height =
    document.getElementById('Body10').contentWindow.
    document.body.scrollHeight;
        //change the height of the iframe
        document.getElementById('Body10').height =
    the_height + 20;
    }
    else if (op == '1') {
        var the_height =
    document.getElementById('Body11').contentWindow.
    document.body.scrollHeight;
        //change the height of the iframe
        document.getElementById('Body11').height =
    the_height + 20;
    }
    else {
        var the_height =
    document.getElementById('Body12').contentWindow.
    document.body.scrollHeight;
        //change the height of the iframe
        document.getElementById('Body12').height =
    the_height + 20;

    }
    top.parent.window.calcHeight()

}

</script>

<body onload="agranda()">
<%pos = trim(request.QueryString("pos") )
  INI = trim(request.QueryString("INI") )
  FIN = trim(request.QueryString("FIN") )
mensa= pos+ini+fin
CAD = "EXEC SP_revisa_fechas_INTERFASE '"&INI&"', '"&FIN&"','"&POS&"' "

rs.open cad,cnn

%>
<table align="center" cellpadding="2" cellspacing="1" bordercolor='<%=application("color1") %>' border="1" id="MSG" name="MSG" style="display:none">
  <tr><td colspan="3"  class="Estilo15">Nada que procesar dentro del rango de fechas/Tienda proporcionados....</td></tr>
</table>
<table align="center" cellpadding="2" cellspacing="1" bordercolor='<%=application("color1") %>' border="1" id="DATA" name="DATA" style="display:none">
  <tr><td colspan="6" align="center" class="Estilo24"><%=rs("nro") %> Documentos a Procesar</td></tr>
   <tr> <td align="center" class="Estilo21">1.-</td>
        <td class="Estilo21" style="text-align:left;">&nbsp;Revisando Rango de Fechas v&aacute;lido para transferencia&nbsp;</td>
	    <td align="center" class="Estilo21"><img src="images/wait.png"  id="espera1" name="espera1" /><img src="images/check.png" id="ok1" name="ok1" /></td>
        <td align="center" class="Estilo21">3.-</td>
        <td align="left" class="Estilo21">&nbsp;Transferencia de Documentos de Venta a Contabilidad&nbsp;</td>
	    <td align="center" class="Estilo21"><img src="images/wait.png"  id="espera3" name="espera3" /><img src="images/check.png" id="ok3" name="ok3" /></td>
   </tr>
  <tr><td colspan="6"  class="Estilo21"><hr /></td></tr>
  <tr> <td align="center" class="Estilo21">2.-</td>
        <td class="Estilo21" style="text-align:left;">&nbsp;Revisar/Actualizar datos Clientes nuevos&nbsp;</td>
	    <td align="center" class="Estilo21"><img src="images/wait.png"  id="espera2" name="espera2" /><img src="images/check.png" id="ok2" name="ok2" /></td>  
        <td align="center" class="Estilo21">4.-</td>
        <td style="text-align:left" class="Estilo21">&nbsp;Actualiza Stocks en SoftCom&nbsp;</td>
	    <td align="center" class="Estilo21"><img src="images/wait.png"  id="espera4" name="espera4" /><img src="images/check.png" id="ok4" name="ok4" /></td>
   </tr>
</table>

<table width="100%" align="center"><tr>
<td  valign="top" align="right" width="30%">
<iframe  width="100%" src="" id="Body10" name="Body10" scrolling="yes" frameborder="0" width="100%" height="10" align="right" style="display:none; vertical-align:top" onLoad="calcHeight(0);" ></iframe>
</td>
<td  valign="top" align="left" width="60%">
<iframe  width="100%" src="" id="Body11" name="Body11" scrolling="yes" frameborder="0" width="100%" height="10" align="left" style="display:none; vertical-align:top" onLoad="calcHeight(1);" ></iframe>
</td>
<td width="10%" >
<iframe  width="100%" src="" id="Body12" name="Body12" scrolling="yes" width="100%" frameborder="0" height="10" align="middle" style="display:none; vertical-align:top" onLoad="calcHeight(2);" ></iframe>
</td>
</tr>
</table>
</body>

<script type="text/jscript" language="jscript">
    NRO = parseInt('<%=rs("nro")%>', 10)
    mensa = ltrim('<%=mensa%>')
    pos = '<%= trim(request.QueryString("pos") ) %>'
    ini = '<%= trim(request.QueryString("ini") ) %>'
    fin = '<%= trim(request.QueryString("fin") ) %>'
 //   alert(mensa.length)
    if (mensa.length == 0) {
        document.all.DATA.style.display = 'none'
        document.all.MSG.style.display = 'none'
    }
    else {
        if (NRO == 0) {
            document.all.MSG.style.display = 'block'
            document.all.DATA.style.display = 'none'
        }
        else {
            document.all.DATA.style.display = 'block'
            document.all.MSG.style.display = 'none'
            document.all.ok1.style.display = 'none'
            document.all.ok2.style.display = 'none'
            document.all.ok3.style.display = 'none'
            document.all.ok4.style.display = 'none'
            clientes();
           // documentos();
           // stock();
        }
    }

function clientes() 
{   document.all.ok1.style.display = 'block'
    document.all.espera1.style.display = 'none'
    document.all.Body10.style.display='block'
    document.all.Body10.src = 'bake/bakecliente_real.asp?pos=' + pos + '&ini=' + ini + '&fin=' + fin
    return true
}

function documentos() {

    document.all.Body11.style.display = 'block'
    document.all.Body11.src = 'bake/bakeinterface_real.asp?pos=' + pos + '&ini=' + ini + '&fin=' + fin
    return true;
}

function stocks() {
   
    document.all.espera4.style.display = 'none'
    document.all.ok4.style.display = 'block'
    document.all.Body12.style.display = 'none'
    document.all.Body12.src = 'bake/bakecorreointerface_real.asp?pos=' + pos + '&ini=' + ini + '&fin=' + fin

    return true
}

</script>

</html>
