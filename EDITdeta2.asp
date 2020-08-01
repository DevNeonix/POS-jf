<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = true %>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos") %>
<%'**************************************************************************
' CAMBIA DATOS DE CABECERAS COMO CLIENTE, NUMERO DE DOCUMENTO Y FECHA 
'**************************************************************************** %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<style>
.botonimagen{
  background-image:url('images/check.png');
  background-repeat: no-repeat;
  background-position: left center;
  cursor:pointer;
  border:none;
  height:35px;
  width:35px;
}
</style>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>


<LINK REL="stylesheet" TYPE="text/css" HREF="VENTAS.CSS">
<!--#include file="comun/funcionescomunes.asp"-->
<!--#include file="includes/funcionesVBscript.asp"-->
<!--#include file="includes/cnn.inc"-->
<!--#include file="comun/comunQRY.asp"-->
<script language="javascript" src="includes/cal.js"></script>
<script language="jscript" type="text/jscript">
function AGRANDA() {
    top.parent.window.document.getElementById('body0').height = 580
}
var oldrow = 1

function dd2(ff) {	// LLENA TEXTBOX ADICIONALES AL COMUN
    // LOS DEL COMUN SON CODIGO Y DESCRIPCION
    var t = document.all.TABLA;
    var pos = parseInt(ff);
    dd(ff, 0);
}
</script>
<%
TDA = request.QueryString("TDA")
ope = request.QueryString("oper")
if trim(tda) = "TT" then
    cad =   " select * from movimcab as mm LEFT OUTER join clientes as cc on cc.cliente = mm.cliente where operacion =  '"&ope&"' "
else
    cad =   " select * from movimcab as mm LEFT OUTER join clientes as cc on cc.cliente = mm.cliente  where tienda = '"&TDA&"' and operacion =  '"&ope&"' "
end if
'RESPONSE.WRITE(cAD)
'response.end
rs.open cad,cnn
%><center>
<span class="Estilo18">
<%
if rs.recordcount <=0 then 
    response.Write("<br><br><br>")
    response.write("Operacion no Existe en la tienda indicada ")
    RESPONSE.End
end if
RS.MOVEFIRsT
tda = rs("tienda")
Set RS2 = Server.CreateObject("ADODB.Recordset")
	RS2.ActiveConnection = Cnn
	RS2.CursorType       = 3 'CONST adOpenStatic = 3
	RS2.LockType         = 1 'CONST adReadOnly = 1
	RS2.CursorLocation   = 3 'CONST adUseClient = 3

	RS2.Open "select * from tiendas where codigo = '"&tda&"'", Cnn

%>
</span></center>
<body onload="AGRANDA()">
<center>
<form id ="thisForm" name= "thisForm" >
<table align="center" cellpadding="3" cellspacing="1" bordercolor='<%=application("color1") %>' border="1" id="TABLA" name="TABLA" 
bordercolorlight="<%=application("color2") %>"  bordercolordark="<%=application("color2") %>" >
    <tr> 
        <td class="Estilo12">Operacion :&nbsp; </td>
        <td class="Estilo1" align="left" colspan="2"> <%=OPE%></td>
           
    </tr>
    <tr> 
        <td class="Estilo12" align="right" >Tienda :&nbsp; </td>
        <td class="Estilo1" align="left" colspan="2"> <%=tda%>-<%=rs2("descripcion") %></td>
    </tr>
    <tr> 
        <td class="Estilo12">Documento : &nbsp;</td>
        <td class="Estilo1" align="left"> <%=RS("coddoc")%></td>
        <td class="Estilo1" align="left"><input id="ser" name="ser" class="estilo1" value='<%=rs("serie")%>' />-<input type="text" name="num" id="num" class="Estilo1" size="20" maxlength="7" 
        value="<%=TRIM(rs("numdoc")) %>"  style="background-color:#F9C1D9;"></td>    
    </tr>
    <tr> 
        <td class="Estilo12" valign="middle" align="right" >Cliente : &nbsp;</td> 
        <td align="left">
            <input type="text" name="CLI" id="CLI" class="Estilo1" onchange="cliente(this.value)" style="background-color:#F9C1D9;"
            size="20" maxlength="11" value="<%=rs("cliente") %>">
        <td><input type="text" name="DES" id="DES" value="<%=rs("nombre") %>" maxlength="50" size="50" style="border:none"
        class="Estilo12" readonly tabindex="-1"></td>
        <td style="display:none"><input type="text" name="DIR" id="DIR" value="" maxlength="100" size="80" class="Estilo12"  readonly tabindex="-1" style="display:none"></td></td>
    </tr>
     <tr> 
        <td class="Estilo12" align="right">Fecha : &nbsp; </td>
        <td class="Estilo1" align="left" colspan="2"><A href="javascript:showCal('Calendar1')"><img height=16 src="images/cal.gif" width=16 border=0></A>
        <input type="text" name="fec" id="fec" class="Estilo1" style="background-color:#F9C1D9;"
            readonly tabindex="-1" value="<%=formatdatetime(rs("fecdoc"),2) %>"></td>
    </tr>
</table>
<center>
<input type="button" class="botonimagen" value="&nbsp;" onclick="graba()" /> 
</center>
<iframe  width="100%" src="" id="mirada" name="mirada" scrolling="yes" frameborder="1" height="40" align="middle" style="display:none" ></iframe>
</form>

<script language="jscript" type="text/jscript">
    addCalendar("Calendar1", "Elija una fecha", "fec", "thisForm")

function cliente(dato) {

    ss = document.all.CLI.value
    cad = "bake/bakecliente.asp?pos=" + trim(dato)
    //   document.all.mirada.style.display='block'
    document.all.mirada.src = cad

    if (ss.length < 7) {
        document.all.CLI.value = strzero(trim(document.all.CLI.value), 11)
        dato = strzero(dato, 11)
    }

}


function graba() {
    ss = '<%=ope%>'
    cad  = "comun/modicab.asp?ope=" + trim(ss)
    cad += '&cli=' + trim(document.all.CLI.value)
    cad += '&fec=' + trim(document.all.fec.value)
    cad += '&num=' + strzero(trim(document.all.num.value), 7)
    cad += '&ser=' + strzero(trim(document.all.ser.value), 3)
    //   document.all.mirada.style.display='block'
    document.all.mirada.src = cad


}
</script>
</center>
</body>
</html>
