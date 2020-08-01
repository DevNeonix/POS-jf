<%@Language=VBScript %>
<%Response.Buffer = true %>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos") %>
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="COMUN/FUNCIONESCOMUNES.ASP"-->
<!--#include file="COMUN/COMUNqry.ASP"-->
<!--#include file="includes/Cnn.inc"-->
<% CAD = "EXEC SP_TIPODECAMBIO"
RS.Open CAD, Cnn
cambio = rs("bajo")
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>CANCELACION</title>
<style type="text/css">
hr {
    border: none;
    /* top    */ border-top: 1px solid #ccc;
    /* middle */ background-color: #f09; color: #f09;
    /* bottom */ border-bottom: 1px solid #eee;
    height: 3px;
 
}
    .style1
    {  width: 100px;
       text-align:left; 
       font-family:arial;  
       font-size:30px;  
       color:#f09;    
       font-weight:800; 
       }
    .style2
    {  width: 150px;
       font-family:Tahoma;  
       font-size:30px;  
       color:Gray;    
       font-weight:800; 
       text-align:right; 
       border:2px;     
    }
   .style3
    {  width: 150px;
       font-family:Tahoma;  
       font-size:20px;  
       color:Gray;    
       font-weight:800; 
       text-align:right; 
       border:2px;   
    }
</style>
</head>
<%pago = request.QueryString("pago") %>
<body  style=" margin-left:10px; margin-right:10px  ">
<table id="Table1" align="center"  bordercolor="<%=Application("color2")%>"   bgcolor="<%=Application("color2")%>"  cellpadding="0"  cellspacing="1"  border="1" width="700px">
   <tr><td colspan="3" class="estilo18">T O T A L </td></tr>
   <tr>
        <td align="right" class="style3">TC = <%=round(cambio,2)%>&nbsp;</td>
        <td align="center" class="estilo18">SOLES</td>
        <td align="center" class="estilo18">U.S. $</td>
        <td align="center" class="style2" style="display:none;">&nbsp;</td>
   </tr>
   <tr>
        <td align="right" class="style1">&nbsp;</td>
        <td align="center" class="style2"><input id="totsol" name="totsol" value="<%=trim(pago)%>" class="style2" readonly tabindex="-1" /></td>
   <td align="center" class="style2"><input id="totdol" name="totdol" value="<%=trim(pago)%>" class="style2" readonly tabindex="-1" /></td>
   <td align="center" class="style2"  style="display:none;">&nbsp;</td></tr>


   <tr><td align="center" colspan="4"><hr  /></td></tr>
   <tr>
        <td class="style1">Visa :&nbsp;</td>
        <td align="center"  bordercolor="<%=Application("color1")%>"><input id="visa" name="visa" class="style2" onKeyUp="this.value=toInt(this.value)" onBlur="visa();" value='' /></td> 
        <td align="center"  bordercolor="<%=Application("color1")%>">&nbsp;</td>  
        <td align="center"  bordercolor="<%=Application("color1")%>" style="display:none;"><input id="v1" name="v1"  size="20"  /></td>  
   </tr>
   <tr>
        <td class="style1">Mastercard :&nbsp;</td>
        <td align="center"  bordercolor="<%=Application("color1")%>"><input id="master" name="master" class="style2" onKeyUp="this.value=toInt(this.value)" onBlur="master();" value='' /></td>
        <td align="center"  bordercolor="<%=Application("color1")%>">&nbsp;</td>  
        <td align="center"  bordercolor="<%=Application("color1")%>" style="display:none;"><input id="m1" name="m1" class="Estilo2"  /></td>  
   </tr>
   
   <tr>
        <td class="style1">Efectivo :&nbsp;</td>
        <td align="center"  bordercolor="<%=Application("color1")%>"><input id="sol" name="sol" class="style2" onKeyUp="this.value=toInt(this.value)" onBlur="efect();" value=''  /></td>  
        <td align="center"  bordercolor="<%=Application("color1")%>"><input id="dol" name="dol" class="style2" onKeyUp="this.value=toInt(this.value)" onBlur="efect();" value=''  /></td>  
        <td align="center"  bordercolor="<%=Application("color1")%>" style="display:none;"><input id="d1" name="d1" class="Estilo2"  /></td>  
   </tr>
   
   <tr>
        <td class="style1">N/Crédito :&nbsp;</td>
        <td align="center"  bordercolor="<%=Application("color1")%>"><input id="nc" name="nc" class="style2" onKeyUp="this.value=(this.value)" onBlur="credi();" /></td>  
        <td align="center"  bordercolor="<%=Application("color1")%>"><input id="sr1" name="sr1" size="4" maxlength="4"/>- <input id="ncr" name="ncr" onchange="nota()" size="15" maxlength="7" /></td>  
        <td align="center"  bordercolor="<%=Application("color1")%>" style="display:none;"><input id="n1" name="n1" class="Estilo2"  /></td>  
   </tr>
   <tr><td align="center" colspan="4"><hr  /></td></tr>
   <tr>
        <td align="right" class="style1">Vuelto :&nbsp;</td>
        <td align="center"  bordercolor="<%=Application("color1")%>"><input id="saldos" name="saldos" class="style2" readonly tabindex="-1" value='' /></td> 
        <td align="center"  bordercolor="<%=Application("color1")%>" colspan="2" >&nbsp;</td>  
   </tr>
   <tr>
    <td colspan="3" class="Estilo2"><input id="sonn" name="sonn" class="Estilo2" tabindex="-1" value='' style="display:none;"/></td>
   </tr>
   <td class="style1">VENDEDOR :&nbsp;</td>
        <td align="center"  bordercolor="<%=Application("color1")%>" ><input id="ven" name="ven" class="style2"  onchange="bake();" value=''  /></td>  
        <td align="center"  bordercolor="<%=Application("color1")%>" colspan="2" ><input id="des" name="des" class="style2" value='' readonly tabindex="-1" /></td>  
   </tr>
   <tr>
    <td colspan="4" align="center"><img src="images/print.jpg" border="0" id="prn" name="prn" onClick="printa();" style="display:block; cursor:pointer"/></td>
   </tr>
   </table>
   <iframe width="100%"  src="" id="cok" name="cok"  height="200" style="display:none;" align="middle">
</iframe>
</body>
<script language="jscript" type="text/jscript">
    var cambio = 0
    var soles = 0
    var dolares = 0
    var v1 = 0
    var m1 = 0
    var d1 = 0
    var ss = 0
	var n1 = 0
    var vuelto = 0
    cambio = '<%=round(cambio,2)%>'
    soles = toInt('<%=pago %>')
document.all.totdol.value = Math.round((soles/cambio)*100)/100
dolares = document.all.totdol.value


function bake() { 

document.all.cok.src='bake/bakevende.asp?pos='+ ltrim(rtrim(document.all.ven.value))
//document.getElementById('prn').style.display = 'block'
}
function largo() {
  // window.moveTo(100, 100)
  // window.resizeTo(800, 600)
   return true;
}
function visa() {
//    document.all.v1.value = redondea(Math.round(ltrim(rtrim(document.all.visa.value)) * 100) / 100, 2)
//    document.all.visa.value = redondea(Math.round(ltrim(rtrim(document.all.visa.value)) * 100) / 100, 2)
//    v1 = Math.round(document.all.v1.value *100)/100
	// FAC 20121231
	v1 = parseFloat(Math.round(document.all.visa.value*100)/100)

    salda();
}
function master() {
//    document.all.m1.value = Math.round(ltrim(rtrim(document.all.master.value)) * 100) / 100
//    document.all.master.value = redondea(Math.round(ltrim(rtrim(document.all.master.value)) * 100) / 100,2)
//    m1 = Math.round(document.all.m1.value*100)/100
	m1 = parseFloat(Math.round(document.all.master.value*100)/100)
	
    salda();
}
function credi() {
    document.all.n1.value = Math.round(ltrim(rtrim(document.all.nc.value)) * 100) / 100
    document.all.nc.value = redondea(Math.round(ltrim(rtrim(document.all.nc.value)) * 100) / 100, 2)
    n1 = Math.round(document.all.n1.value * 100) / 100
    salda();
}


function efect() {
    d1 = 0
    document.all.d1.value = ''

    if (trim(document.all.dol.value) != '') {
        d1 += parseFloat(Math.round(document.all.dol.value*100)/100)*cambio
        document.all.d1.value = d1
        document.all.dol.value = d1 //redondea(document.all.dol.value, 2)    
    }

    if (trim(document.all.sol.value) != '') {
        d1 += parseFloat(Math.round(document.all.sol.value * 100) / 100)
        document.all.d1.value = d1
        document.all.sol.value = d1 //redondea(document.all.sol.value ,2)
    }

    salda();
}

function salda() {
ss = soles - (parseFloat(v1) + parseFloat(m1) + parseFloat(d1)+ parseFloat(n1)) 
    

if (parseFloat(ss) < 0) {
    document.getElementById('saldos').className = 'estilo18'
   
}
else {
    document.getElementById('saldos').className = 'style2'

}

document.all.saldos.value = Math.abs(Math.round(ss * 100) / 100)
vuelto = document.getElementById('saldos').value
if (document.getElementById('saldos').className == 'estilo18') {
    document.getElementById('sonn').style.display = 'block'
    document.getElementById('sonn').value = 'Son : ' + FComson(document.getElementById('totsol').value)
}

return true;
}



function nota() {
    document.all.cok.src = 'bake/bakenota.asp?ser=' + ltrim(rtrim(document.all.sr1.value)) + '&ncr=' + ltrim(rtrim(document.all.ncr.value))
}


function printa() {

    if (n1 + v1 + m1 + d1 + ss == 0) {
        alert("Debe informar la forma y monto de pago ")
    return true}
    pagos = parseFloat(v1) + parseFloat(m1) + parseFloat(d1)+ parseFloat(n1) 
    totales = Math.round( pagos* 100)/100
   
    if ( Math.round(soles*100) - Math.round(totales*100) > 0 ) {
        alert("Monto a Pagar INCOMPLETO")
        return true;
    }

    if (ltrim(rtrim(document.all.nc.value)) != '') {
        if (parseInt(document.all.nc.value, 10) > 0) {
            if (ltrim(rtrim(document.all.sr1.value)) == '' || ltrim(rtrim(document.all.ncr.value)) == '') {
                alert("Debe Informar la Serie y número de la Nota de Credito con la que cancela")
                return true;
            }

            if (ltrim(rtrim(document.all.sr1.value)) == '' || ltrim(rtrim(document.all.ncr.value)) == '')
                return true;
        }
    }
   
    if (ltrim(rtrim(document.all.ven.value)) == '') {
        alert("Por favor informe su código de Vendedor")
        return false;
    }

/*  if (ltrim(rtrim(document.all.des.value)) == '') {
        alert("Vendedor no registrado")
        return false;}
*/

//alert(window.opener.parent.window.frames[0].window.document.all.miRadio[0].checked)




    if (window.opener.parent.window.frames[0].window.document.all.miRadio[0].checked == true) {
        docum = 'BOLETA # : ' + window.opener.parent.window.frames[0].window.document.all.BOL.value
        serie = Left(trim(window.opener.parent.window.frames[0].window.document.all.BOL.value), 3)
    }
    else {
        docum = 'FACTURA # : ' + window.opener.parent.window.frames[0].window.document.all.FAC.value
        serie = Left(trim(window.opener.parent.window.frames[0].window.document.all.FAC.value), 3)
    }

cad = '¿ Desea imprimir la ' + docum + '?'

if (confirm(cad) == false) {
        return true;
    }
    // descuento el vuelto para que la caja cuadre
    solez = parseFloat(document.all.sol.value)- parseFloat(vuelto)
// Primero pasa valores a cookies de caja, 
    setCookie('vis', document.all.visa.value)
    setCookie('mas', document.all.master.value)
    setCookie('sol', solez)
    setCookie('dol', document.all.dol.value)
    setCookie('ven', document.all.ven.value)
    setCookie('cre', document.all.nc.value)
    setCookie('ser', serie)
    setCookie('sr1', document.all.sr1.value)
    setCookie('ncr', document.all.ncr.value)
// ahora GRABA movimientos , 
    window.opener.parent.window.frames[1].window.graba()

// luego imprime.....
var opc = "directories=no,height=600,";
opc = opc + "hotkeys=no,location=no,";
opc = opc + "menubar=no,resizable=yes,";
opc = opc + "left=0,top=0,scrollbars=yes,";
opc = opc + "status=no,titlebar=yes,toolbar=yes,";
opc = opc + "width=900";
// window.open('boleta.asp', '', opc)
this.window.close();
}
</script>
</html>
