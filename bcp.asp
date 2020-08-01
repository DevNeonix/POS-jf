<%@ Language=VBScript %>
<%Session.LCID=2058%>
<% Response.Buffer = true %>
<%bcp = Request.Cookies("tienda")("bcp") 	%>
<script type="text/jscript" language="jscript">

</script>
<%' Definir el tamaño de la pagina
titulo = "Oferta BCP --> DNI"
%>
<script type="text/jscript" language="jscript">
// **************************************************************
//  Indicar el nombre de la página donde se realizan los cambios 
// **************************************************************
var funcionalidad = 'comun/INSERDNI.asp?'

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" lang="es-pe" />
<title><%=titulo%></title>
</head>
<body onkeyup="veri()" >

<form name="thisForm" id="thisForm" method="post" action="bcp.asp">

<table width="100%">
	<tr><td colspan="3"><hr /></td></tr>
    <tr><td align="center" class="Estilo6"><%=titulo%></td></tr>
	
</table>

<!--#include file="includes/Cnn.inc"-->
<!--#include file="COMUN/FUNCIONESCOMUNES.ASP"-->


<table	align="center" border="0" id="DATAENTRY"  >
   <tr>
   		<td>
            <table	align="center" cellpadding="1" cellspacing="5"  bgcolor="WHITE" border="1" >
              <tr valign="middle"> 
                <td bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  DNI : &nbsp;&nbsp;  </td>
                <td bgcolor="WHITE" valign="middle" width="150"> 
                  <input type="text" id="COD" name="COD"  class="Estilo2" onkeyup="Ingresar()" maxlength="8" />
                </td>
   
              </tr>
              <tr id="dn1" name="dn1" style="display:none">
            <td colspan="2"><input id="dni" name="dni" value="" class="Estilo2" readonly tabindex="-1" /></td>
         </tr> 
         <tr id="ok" name="ok" style="display:none">
            <td colspan="2"><input id="o1" name="o1" value="" class="Estilo2" readonly tabindex="-1" /></td>
         </tr> 

         <tr id="mal" name="mal" style="display:none">
            <td colspan="2"><input id="m1" name="m1" value="" class="Estilo2" readonly tabindex="-1" /></td>
         </tr> 
         <tr id="res" name="res" >
            <td colspan="2"><input type="reset" id="r1" name="11" value="NUEVA CONSULTA" style="background-color:#c82f8a; color:#fff; font-weight:bolder" /></td>
         </tr> 
</table>

<iframe frameborder="1" style="display:none" height="100" width="100%" id="ACTIV" name="ACTIV"></iframe>
<script type="text/jscript" language="jscript">
//limpia()
document.getElementById("COD").focus()
function veri() {

    dd = trim(document.all.COD.value)
    if (dd.length == 8) {
        Ingresar()
        return true;
    }
}
function Ingresar() {

    document.all.COD.value = toInt(document.all.COD.value)
    cad = "BAKE/BAKEBCP.asp?COD=" + document.all.COD.value
    dd = trim(document.all.COD.value)

    if (dd.length == 8) {
     //alert(cad)
        document.all.ACTIV.src = cad
    }
    else {
        return true
    }

}

function limpia() {
    document.all.o1.value = ''
    document.all.m1.value = ''
    document.all.ok.style.display = 'none'
    document.all.mal.style.display = 'none'
   
}
</script>
</form>
</body>
</html>
