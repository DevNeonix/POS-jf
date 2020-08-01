
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = true%>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos")%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<!--#include file="comun/funcionescomunes.asp"-->
<!--#include file="includes/funcionesVBscript.asp"-->
<!--#include file="includes/cnn.inc"-->


<script type="text/jscript" language="jscript">
	function calcHeight() {
		//find the height of the internal page
		var the_height = document.getElementById('mirada').contentWindow.document.body.scrollHeight;
		//change the height of the iframe
		document.getElementById('mirada').height = the_height + 20;
	}
</script>

<script type="text/jscript" language="jscript">
function strSQL() {

Doc = ''
for (i = 0; i < 5; i++) {
    if (document.all.tip[i].checked == true) {
        break;
    }

}

cad = "EDITdeta"+document.all.tip[i].value+".asp?TDA=" + document.all.tienda.value 
cad += '&oper=' + strzero(document.all.oper.value ,10)

    //alert(cad)
    parent.window.frames[1].window.location.replace(cad)
    return true
}
</script>


<% CAD =   " SELECT * FROM TIENDAS WHERE ESTADO ='A' order by descripcion "
          '  response.write(cad)
          '  response.write("<br>")
    RS.OPEN CAD,CNN
    IF rs.recordcount > 0 THEN rs.movefirst
%>

<body onload="top.parent.window.document.getElementById('body0').height = 480">

<form id ="thisForm" name= "thisForm" >

<table width="100%">
    <tr><td align="center" class="Estilo6">Edici&oacute;n/Modificaci&oacute;n de Operaciones :</td></tr>
</table>

<table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="2" cellspacing="1"  border="0">
	<tr valign="middle" >
    	<td class="Estilo11" valign="middle" align="right" rowspan="2">
            <label for="Radio">Tienda:&nbsp;</label></td> 
        <td class="Estilo12" align="left"  rowspan="2">
            <select  name="tienda" id="tienda">
                <option value = "TT" selected>TODAS</option>
                <%do while not rs.eof %>
                    <option value="<%=TRIM(RS("CODIGO"))%>"><%=TRIM(RS("DESCRIPCION")) %></option>
                    <%rs.movenext %>
                <%loop %>
                <%rs.close %>
            </select>
        </td>
        
        
        <td  class="Estilo12" align="left"  rowspan="2">
           <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0">
              <tr>
                   <td> <input type="radio" id="tip" name="tip" value="0">Borrar Operaci&oacute;n (sin actualizar stocks)</td>
                   <td> <input type="radio" id="tip" name="tip" value="1">Borrar Operaci&oacute;n (recupera stocks)</td>
              </tr>
              <tr>
                   <td> <input type="radio" id="tip" name="tip" value="2">Modifica Cabecera Documento</td>
                   <td> <input type="radio" id="tip" name="tip" value="3">Linea de Detalle</td>
              </tr>
              <tr>
                   <td> <input type="radio" id="tip" name="tip" value="4">Modifica Datos Pago</td>
                   <td><input type="radio" id="Radio1" name="tip" value="5">Cambia Flag Nota de Crédito</td>
              </tr>
              <tr>
              <td colspan="2">
                <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0">
              <tr>
                   <td class="Estilo11" valign="middle" align="right">Operación:&nbsp;&nbsp;</td>
                  <td> <input id="oper" name="oper" type="text" maxlength="10"  class="inputs" /></td>
               </tr>
               </table> 
                  </td>
              </tr>
           </table>
           
                
        </td>
 
        <td rowspan="2"><img src="images/ok.gif" onclick="strSQL()" style="cursor:pointer;"/></td>            
    </tr>
   
</table>

<iframe src="" id="mirada" name="mirada" style="display:none" width="100%"></iframe> 
</form>
</body>
<script language="jscript" type="text/jscript">

</script>
</html>
