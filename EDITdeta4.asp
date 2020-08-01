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
    cad =   " select * from caja where operacion =  '"&ope&"' and estado ='a' order by lin "
else
    cad =   " select * from caja as mm where tienda = '"&TDA&"' and operacion =  '"&ope&"' and estado = 'a' order by lin"
end if
RESPONSE.WRITE(cAD)
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
<table align="center" cellpadding="3" cellspacing="1" bordercolor='<%=application("color2") %>' border="1" id="TABLA" name="TABLA" 
bordercolorlight="<%=application("color2") %>"  bordercolordark="<%=application("color1") %>" >
    <tr> 
        <td class="Estilo12">Operacion :&nbsp; </td>
        <td class="Estilo1" align="left" colspan="1"> <%=OPE%></td>
        <td class="Estilo12" width="5%">&nbsp; </td>
        <td class="Estilo12" align="right" >Tienda :&nbsp; </td>
        <td class="Estilo1" align="left" colspan="1"> <%=tda%>-<%=rs2("descripcion") %></td>
    </tr>
    <%rs.movefirst%>
    <%con =0%>
    <tr> 
        <td class="Estilo12" width="5%">Lin </td>
        <td class="Estilo12" width="10%">Moneda</td>
        <td class="Estilo12" width="20%">Tipo</td>
        <td class="Estilo12" width="10%">Monto</td>
        <td class="Estilo12" width="40%">Doc Ref</td>
    <%do while not rs.eof%>
    <tr> 
        <td class="Estilo1" align="center"> <input style="border:none" readonly="readonly" id="linea<%=RS("lin")%>" value="<%=RS("lin")%>" /></td> 
        <td class="Estilo1" align="left">
            <select  class="INPUTS" id="mon<%=con%>" name="mon<%=con%>"  value="<%=RS("moneda")%>" style="text-align:center;width:100%">
                <option value=""></option>
                <option value="MN">MN</option>
                <option value="US">US</option>
           </select> 
       </td>
       <td class="Estilo1" align="left">
            <select  class="INPUTS" id="tip<%=con%>" name="tip<%=con%>"  value="<%=RS("tipo")%>" style="text-align:center;width:100%">
                <option value=""></option>
                <option value="SOL">SOLES</option>
                <option value="DOL">DOLARES</option>
                <option value="MAS">MASTERCARD</option>
                <option value="VIS">VISA</option>
                <option value="NCR">NOTA DE CREDITO</option>
           </select> 
       </td> 
       <td class="Estilo1" align="right"><input id="sol<%=con%>" name="sol<%=con%>"  class="INPUTS" style="width:100%" value='<%=trim(RS("MONTO"))%>' onfocus="seleccionar(this)" /></td> 
       <td class="Estilo1" align="left" ><input id="obs<%=con%>" name="obs<%=con%>"  class="INPUTS" style="width:100%" value='<%=RS("nota")%>' onfocus="seleccionar(this)" /></td> 
        <td><input type="button" class="botonimagen" value="&nbsp;" onclick="graba('<% =RS("OPERACION") %>', document.getElementById('mon<%=con%>').value, document.getElementById('tip<%=con%>').value, document.getElementById('sol<%=con%>').value, document.getElementById('linea<%=RS("lin")%>').value,document.getElementById('obs<%=con%>').value)" /> </td>
    </tr>
        <script language=jscript type="text/jscript">
            cn = parseInt('<%=con%>',10)
            var subcadena = '<%=RS("MONeda")%>';
            var elemento = eval("document.all.mon"+cn);
            eval("document.all.mon"+cn+".selectedIndex = seleindice(subcadena, elemento)");
            var subcadena = '<%=RS("tipo")%>';
            var elemento = eval("document.all.tip"+cn);
            eval("document.all.tip"+cn+".selectedIndex = seleindice(subcadena, elemento)");
        </script>
        <%rs.movenext%>
        <%con = con + 1 %>
    <%loop%>   
</table>
<center>
<img src="images/NEW.GIF" onclick="nnuevalinea()"/>
</center>
<iframe  width="100%" src="" id="mirada" name="mirada" scrolling="yes" frameborder="1" height="200" align="middle" style="display:none" ></iframe>
</form>

<script language="jscript" type="text/jscript">
   
function nnuevalinea(){
	if(confirm("Nueva linea?")){
	document.all.mirada.style.display='block'
   document.all.mirada.src='./nuevalineacaja.asp?ope=<%=ope%>&tda=<%=tda%>'
	}
   
}

function graba(operacion,moneda,tipo,monto,linea,obs) {
    cad  = "comun/modicaja.asp?ope=" + operacion
    cad += '&mon=' + moneda
    cad += '&tipo=' + tipo
    cad += '&monto=' + monto
    cad += '&obs=' + obs
    cad += '&lin=' + linea
    //   document.all.mirada.style.display='block'
    alert(operacion+" mon:"+moneda+" tipo:"+tipo+" mto:"+monto+" linea:"+linea+" obs"+obs);
    document.all.mirada.src = cad
}

</script>
</center>
</body>
</html>
