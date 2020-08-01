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
        top.parent.window.document.getElementById('body0').height = 640
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
    cad =   " select coddoc, serie, numdoc, cc.cliente, nombre, fecdoc,  tienda, isnull(pvp, 0) as pvp, isnull(descuento, 0) as descuento, isnull(subtot,0) as subtot, isnull(igv,0) as igv, isnull(total,0) as total, tipmov from movimcab as mm LEFT OUTER join clientes as cc on cc.cliente = mm.cliente where operacion =  '"&ope&"' "
else
    cad =   " select coddoc, serie, numdoc, cc.cliente, nombre, fecdoc, tienda, isnull(pvp, 0) as pvp, isnull(descuento, 0) as descuento, isnull(subtot,0) as subtot, isnull(igv,0) as igv, isnull(total,0) as total, tipmov  from movimcab as mm LEFT OUTER join clientes as cc on cc.cliente = mm.cliente  where tienda = '"&TDA&"' and operacion =  '"&ope&"' "
end if
'RESPONSE.WRITE(cAD)
'response.end
%><center>
<span class="Estilo18">
<%
rs.open cad,cnn
if rs.recordcount <=0 then 
    response.Write("<br><br><br>")
    response.write("Operacion no Existe en la tienda indicada ")
    RESPONSE.End
else
    rs.movefirst
    if rs("tipmov") = "E" then
        response.Write("<br><br><br>")
        response.Write("Operacion es un Ingreso")
        response.End
    End if
end if
RS.MOVEFIRsT
tda = rs("tienda")
TIP = RS("TIPMOV")
Set RS2 = Server.CreateObject("ADODB.Recordset")
	RS2.ActiveConnection = Cnn
	RS2.CursorType       = 3 'CONST adOpenStatic = 3
	RS2.LockType         = 1 'CONST adReadOnly = 1
	RS2.CursorLocation   = 3 'CONST adUseClient = 3
rs2.open "select igv from parametros",cnn
igvpor = cdbl(rs2("igv"))/100
rs2.close

	RS2.Open "select * from tiendas where codigo = '"&tda&"'", Cnn
%>
</span>
</center>
<body onload="AGRANDA()">
<center>
<form id ="thisForm" name= "thisForm" >
<table align="center" cellpadding="3" cellspacing="1" bordercolor='<%=application("color1") %>' border="0" id="TABLA" name="TABLA" 
bordercolorlight="<%=application("color2") %>"  bordercolordark="<%=application("color2") %>" >
    <tr> 
        <td class="Estilo12">Operacion :&nbsp; </td>
        <td class="Estilo1" align="left" colspan="2"><%=OPE%></td> 
        <td rowspan="5">&nbsp;&nbsp;</td>
        <td class="Estilo8" align="center"> P.Bruto </td> 
        <td class="Estilo8" align="center"> Descuento </td> 
        <td class="Estilo8" align="center"> Sut-Total </td> 
        <td class="Estilo8" align="center"> I.G.V. </td> 
        <td class="Estilo8" align="center"> Total </td> 
    </tr>
    <tr> 
        <td class="Estilo12" align="right">Tienda :&nbsp; </td>
        <td class="Estilo1" align="left" colspan="2"> <%=tda%>-<%=rs2("descripcion") %></td>
        <td class="Estilo1" align="right" colspan="1"><%=formatnumber(rs("pvp"),2,,true) %></td>
        <td class="Estilo1" align="right" colspan="1" style="background-color:#F8D3ED"><%=formatnumber(cdbl(rs("descuento"))*-1,2,,true) %></td>
        <td class="Estilo1" align="right" colspan="1"><%=formatnumber(rs("subtot"),2,,true) %></td>
        <td class="Estilo1" align="right" colspan="1"><%=formatnumber(rs("igv"),2,,true) %></td>
        <td class="Estilo1" align="right" colspan="1"><%=formatnumber(rs("total"),2,,true) %></td>
    </tr>
    <tr> 
        <td class="Estilo12">Documento : &nbsp;</td>
        <td class="Estilo1" align="left"> <%=trim(RS("coddoc"))%></td>
        <td class="Estilo1" align="left"><%=rs("serie")%>-<%=TRIM(rs("numdoc")) %></td>    

    </tr>
    <tr> 
        <td class="Estilo12" valign="middle" align="right" >Cliente : &nbsp;</td> 
        <td class="Estilo1" align="left"><%=rs("cliente") %></td>
        <td class="Estilo1" align="left"><%=rs("nombre") %></td>
    </tr>
     <tr> 
        <td class="Estilo12" align="right">Fecha : &nbsp; </td>
        <td class="Estilo1" align="left" colspan="2"><%=formatdatetime(rs("fecdoc"),2) %></td>
    </tr>
</table>
<script language="jscript" type="text/jscript">
    totdoc = parseInt(parseFloat('<%=rs("total")%>') * 100,10)
    brudoc = parseInt(parseFloat('<%=rs("pvp")%>') * 100,10)
    igvdoc = parseInt(parseFloat('<%=rs("igv")%>') * 100, 10)
    dctdoc = parseInt(parseFloat('<%=rs("descuento")%>') * 100, 10)
    cliente = '<%=rs("cliente") %>'
</script>



<table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="0"  cellspacing="1"  border="1" align="center" >
   <tr>
        <td align="center" class="Estilo8">IT</td>
	    <td align="center" class="Estilo8">Codigo</td>
        <td align="center" class="Estilo8">Descripcion</td>
        <td align="center" class="Estilo8">Cant.</td>
        <td align="center" class="Estilo8">PVP</td>
        <td align="center" class="Estilo8">%</td>
        <td align="center" class="Estilo8">Dcto.</td>
        <td align="center" class="Estilo8">IGV</td>
        <td align="center" class="Estilo8">Total</td>
        <td align="center" class="Estilo8" colspan="2" style="display:none">Total</td>
   </tr>
<%rs.close
pvt = 0
dct = 0
igv = 0
tot = 0
CAD=    " SELECT M2.CODART, M2.ITEM , isnull(M2.PORDES,0) as pordes, M2.SALE, M2.ENTRA, m2.igv,         " & _ 
        " V1.DESCRI, isnull(descuento,0)as descuento, ROUND(isnull(M2.PRECIO,0),2) AS PRECIO,           " & _
        " isnull(M2.PORDES,0) as pordes, PVT =case when isnull(m2.sale,0) = 0                           " & _
        " then isnull(M2.PRECIO,0) - isnull(M2.IGV,0) + isnull(M2.DESCUENTO,0)                          " & _
        " else (isnull(M2.PRECIO,0) - isnull(M2.IGV,0) + isnull(M2.DESCUENTO,0)) / isnull(M2.SALE,1) end  " & _ 
        " FROM MOVIMDET AS M2 inner join view_ARTICULOS_TIENDA V1       " & _
        " on LTRIM(RTRIM(M2.CODART)) = LTRIM(RTRIM(V1.CODIGO))                                     " & _
        " WHERE M2.OPERACION ='"&OPE&"'  AND V1.TIENDA='"&TDA&"'        " & _
        " ORDER BY M2.ITEM                                              " 
 '       RESPONSE.WRITE(CAD)
rs.open cad, cnn  
'response.end
if rs.recordcount > 0 then  RS.MOVEFIRST
I=0%>
    <%DO WHILE NOT RS.EOF%>
        <tr id="lin<%=i%>" name="lin<%=i%>" ondblclick="alert(this.id)">
            <td class="Estilo12" valign="middle" align="center"><%=RS("ITEM")%></td> 
            <td><input id="COD<%=i%>" name="COD<%=i%>" size="20" class="Estilo24" maxlength="25" value='<%=RS("CODART") %>'  onchange="carga('<%=i%>');stock('<%=i%>');descuento('<%=i%>')" ondblclick="hlp('<%=i%>', this)" /></td>
            <td><input id="DES<%=i%>" name="DES<%=i%>" size="60" class="Estilo13" readonly tabindex="-1" value='<%=RS("DESCRI") %>'/></td>
            <td><input id="CAN<%=i%>" name="CAN<%=i%>" size="8"  class="Estilo133" maxlength="2" 
            value='<%IF TIP ="S" THEN RESPONSE.WRITE(RS("SALE")) ELSE RESPONSE.WRITE("ENTRA") %>' onchange="stock('<%=i%>');descuento('<%=i%>')" /></td>
            <%if cdbl(rs("sale")) = 0 then sale = 1 else sale = cdbl(rs("sale")) %>
            <td><input id="PVP<%=I%>" name="PVP<%=i%>" size="10" class="Estilo133" maxlength="6" value='<%=cdbl(RS("PVT"))%>' /></td>
            <td><input id="DCT<%=i%>" name="DCT<%=i%>" size="12" class="Estilo133" value="<%=RS("pordes")%>"  onchange="descuento('<%=i%>');"/></td>
            <td><input id="POR<%=i%>" name="POR<%=i%>" size="10" class="Estilo133" value='<%=RS("descuento")%>'  /></td>
            <td><input id="IGV<%=i%>" name="IGV<%=i%>" size="10" class="Estilo133" value='<%=RS("IGV")%>'/></td>
            <td><input id="TOT<%=I%>" name="TOT<%=i%>" size="10" class="Estilo133" value='<%=FORMATNUMBER(RS("PRECIO"),2,,,TRUE)%>' /></td>
            <td><input id="STK<%=i%>" name="STK<%=i%>" size="10" class="Estilo133" value='' style="display:none" /></td>
            <td><input id="PDT<%=i%>" name="PDT<%=i%>" size="10" class="Estilo133" value='' style="display:none" /></td>

        </tr>
        <%I= I + 1 
        pvt = pvt + cdbl(rs("pvt"))
        dct = dct + cdbl(rs("descuento"))
        igv = igv + cdbl(rs("igv"))
        tot = tot + cdbl(rs("precio"))
        %>
        <%RS.MOVENEXT%>
    <%LOOP%>
    <% j = i
    for i=j to 15%>
    <tr id="lin<%=i%>" name="lin<%=i%>" <%if i > j then%> style="display:none" <%end if %> onclick="SHOW('<%=I+1%>')">
           <td class="Estilo12" valign="middle" align="center"><%=right("00"&i,2)%></td> 
           <td><input id="COD<%=i%>" name="COD<%=i%>" size="20" class="Estilo24" maxlength="25" value='' onchange="carga('<%=i%>');stock('<%=i%>');" ondblclick="hlp('<%=i%>', this)" /></td>
           <td><input id="DES<%=i%>" name="DES<%=i%>" size="60" class="Estilo13" readonly tabindex="-1" value=''/></td>
           <td><input id="CAN<%=i%>" name="CAN<%=i%>" size="8"  class="Estilo133" maxlength="2" value='' onchange="stock('<%=i%>');descuento('<%=i%>')" /></td>
           <td><input id="PVP<%=i%>" name="PVP<%=i%>" size="10" class="Estilo133" maxlength="6" value='' /></td>        
           <td><input id="DCT<%=i%>" name="DCT<%=i%>" size="12" class="Estilo133" value="" onchange="descuento('<%=i%>')"/></td>
           <td><input id="POR<%=i%>" name="POR<%=i%>" size="10" class="Estilo133" maxlength="6" value=''  /></td>
           <td><input id="IGV<%=i%>" name="IGV<%=i%>" size="10" class="Estilo133" maxlength="6" value=''/></td>
           <td><input id="TOT<%=i%>" name="TOT<%=i%>" size="10" class="Estilo133" maxlength="6" value='' /></td>
           <td><input id="STK<%=i%>" name="STK<%=i%>" size="10" class="Estilo133" maxlength="6" value='' style="display:none" /></td>
           <td><input id="PDT<%=i%>" name="PDT<%=i%>" size="10" class="Estilo133" maxlength="6" value='' style="display:none" /></td>
    </tr>
    <%next%>
    <tr class="Estilo8" align="right">
        <td colspan="4">Totales : </td>
        <td><input id="pvp" name="PVT" size="10"  maxlength="6" value='<%=formatnumber(PVT,2,,true)%>' class="Estilo8"  style="text-align:right;" readonly tabindex="-1" /></td>
        <td><input id="bruto" name="BRU" size="10" maxlength="6" value='' class="Estilo8"  style="text-align:right;" readonly tabindex="-1" /></td>
        <td><input id="dcto" name="DCT" size="12"  value="<%=formatnumber(dct,2,,true)%>" class="Estilo8"  style="text-align:right;" readonly tabindex="-1" /></td>
        <td><input id="igv" name="IGV" size="10"  maxlength="6" value='<%=formatnumber(igv,2,,true)%>' class="Estilo8"  style="text-align:right;" readonly tabindex="-1" /></td>
        <td><input id="tota" name="TOT" size="10" maxlength="6" value='<%=FORMATNUMBER(tot,2,,,TRUE)%>' class="Estilo8"  style="text-align:right;" readonly tabindex="-1" /></td>
    </tr>
</table>
<center>
<input type="button" class="botonimagen" value="&nbsp;" onclick="graba()" /> 
</center>
<iframe  width="100%" src="" id="mirada" name="mirada" scrolling="yes" frameborder="1" height="240" align="middle" style="display:none" ></iframe>
</form>
<script language="jscript" type="text/jscript">
var IGV = parseFloat('<%=cdbl(igv)/100%>')
// codigo de producto
var aCod = new Array('', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '')
// cantidad de venta
var aCan = new Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
// disponible para venta
var aDis = new Array()
// stock en almacen
var aStk = new Array()
// cantidad * precio de venta
var aTot = new Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
// cantidad * (precio de venta  - descuento)
var aNet = new Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
// % del descuento
var aDct = new Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
// % del descuento
var aPor = new Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
// Precio de venta publico
var aPvp = new Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
// precio con descuento aplicado
var aDes = new Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
// monto del descuento
var aMon = new Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
// IGV DE CADA LINEA
var aIgv = new Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

for (o = 0; o <= 15; o++) {
    if (trim(eval("document.all.CAN" + o + ".value"))!= '')
    { stock(o)}
}
TOTALES()

function hlp(op, obj) {
    vobj = trim(obj.value)
    // alert(vobj.length)
    window.open('help/hlpprendasVT.asp?pos=' + trim(vobj) + '&op=' + op + '&tda=' + '<%=tda%>')
    return true;
}

function graba() {

    //alert("en proceso de validacion....")  
    if (parseInt(totdoc/100,10) !=  parseInt(document.all.tota.value,10)) {
        alert("Los totales de Cabecera y Detalle no coinciden")
        alert(totdoc)
        alert(parseInt(document.all.tota.value * 100, 10))
        return false

    }
/*  if (brudoc != parseInt(document.all.pvp.value * 100, 10)) {
        alert("Los Precios sin IGV de Cabecera y Detalle no coinciden")
        return false
    }
    if (igvdoc != parseInt(document.all.igv.value * 100, 10)) {
        alert("Los IGV de Cabecera y Detalle no coinciden")
        return false
    }
*/
    ss = '<%=ope%>'
    // ARREGLA LOS DECIMALES PARA LA GRABACION
    for (K = 0; K < 15; K++) {
        aCan[K] = parseInt(aCan[K], 10)
        aTot[K] = parseInt(aTot[K] * 100, 10) / 100
        aMon[K] = parseInt(aMon[K] * 100, 10) / 100
        aDct[K] = parseInt(aDes[K] * 100, 10) / 100
        aIgv[K] = parseInt(aIgv[K] * 100, 10) / 100
        aPvp[K] = parseInt(aPvp[K] * 100, 10) / 100
    }


    // Datos para grabar cabecera y detalle
    cad = 'comun/modideta.asp?cli=' + trim(cliente)
    cad += '&cod=' + aCod
    cad += '&Can=' + aCan
    // neto
    cad += '&PVT=' + aTot
    // PRECIO INCLUIDO descuento
    cad += '&pdt=' + aDes
    // porcentaje del descuento
    cad += '&por=' + aPor
    // monto DEL DESCUENTO
    cad += '&des=' + aMon
    // valor del igv
    cad += '&igg=' + aIgv
    // % del IGV
    cad += '&porI=' + '<%=igv%>'
    cad += '&PVP=' + brudoc/100
    cad += '&dct=' + dctdoc/100
    cad += '&bru=' + brudoc/100
    cad += '&igv=' + igvdoc/100
    cad += '&net=' + totdoc / 100
    cad += '&tda=' + '<%=tda%>'
    cad += '&ope=' + '<%=ope%>'
   // document.all.mirada.style.display = 'block'
   // alert(cad)
    document.all.mirada.src = cad;
    return true
}

function carga(cn) { 
     cn = parseInt(cn, 10)
     tda = '<%=tda%>'
     dato = eval("window.document.all.COD" + cn + ".value");
     cad = "bake/bakeprendasVT.asp?pos=" + trim(dato) + "&op=" + cn + '&tda='+trim(tda);
     //document.all.mirada.style.display = 'block'
     document.all.mirada.src = cad;
     stock(cn);
     eval("window.document.all.DCT" + cn + ".value=0");
     descuento(cn);
     return true; 
}

function TOTALES() {  // suma prendas
    CAN = 0
    TOTAL = 0
    DCT = 0
    MON = 0
    PVP = 0
    IGV = 0
    for (m = 0; m < 15; m++) {
        CAN += parseInt(aCan[m], 10)
        TOTAL += parseInt(aTot[m] * 100, 10)
        MON += parseInt(aMon[m] * 100, 10)
        DCT += parseInt(aDes[m] * 100, 10)
        IGV += parseInt(aIgv[m] * 100, 10)
        PVP += parseInt(aCan[m] * aPvp[m] * 100, 10)
        // alert(aDct[m])
    }
    TOTAL = TOTAL / 100
    MON = MON / 100
    DCT = DCT / 100
    IGV = IGV / 100
    PVP = PVP / 100
    boleta = PVP

    document.all.pvp.value = cerea(PVP, 2)
    document.all.dcto.value = cerea(MON, 2)
    document.all.bruto.value = cerea(PVP - MON, 2)
    document.all.igv.value = cerea(IGV, 2)
    document.all.tota.value = cerea(TOTAL, 2)
}
function descuento(op) {
    var IGV = parseFloat('<%=cdbl(igvpor)%>')
    //alert(IGV)
    op = parseInt(op, 10)
    aDct[op] = parseFloat(eval("window.document.all.DCT" + op + ".value"))
    aPor[op] = parseFloat(eval("window.document.all.DCT" + op + ".value"))
   /* alert(aPor[op])
    alert(aDct[op])
    */
    // SI EXISTE VALOR DE DESCUENTO HAY QUE REDONDEARLO TAMBIEN

    if (parseFloat(aDct[op]) > 0) {
        // PRECIO CON DESCUENTO
        pdt = aPvp[op] - (aDct[op] * aPvp[op] / 100)
        // alert(pdt)
        pdt = Math.round(pdt * 100) / 100
        //  alert(pdt)

    }
    else {
        pdt = Math.round(aPvp[op] * 100) / 100
        aDct[op] = 0
    }
    // precio con el descuento
    aDes[op] = pdt

    eval("window.document.all.PDT" + op + ".value=cerea(pdt,2)");
    //alert(aDes[op])
   
    // importe DEL DESCUENTO
    aMon[op] = Math.round((aPvp[op] - aDes[op]) * aCan[op] * 100) / 100

    //   IGV ???
    PPP = aCan[op] * pdt
    //alert(PPP)
    aTot[op] = Math.round(PPP * 100) / 100

    aIgv[op] = aTot[op] * IGV
    aIgv[op] = Math.round(aIgv[op] * 100) / 100
    //alert(cerea(aIgv[op]), 2)
    //alert(aTot[op])
    aTot[op] += aIgv[op]
    //alert(aTot[op])
    eval("window.document.all.IGV" + op + ".value=aIgv[op]");
    eval("window.document.all.TOT" + op + ".value=cerea(aTot[op],2)");
    eval("window.document.all.POR" + op + ".value=aMon[op]")
    TOTALES();
    return true
}

function stock(op) {
//alert(op)
    aDisp = new Array();
    op = parseInt(op, 10);
    xx = parseInt(toInt(eval("window.document.all.CAN" + op + ".value")), 10)

    if (isNaN(xx)) {
        xx = 0
    }
    eval("window.document.all.CAN" + op + ".value=xx");
    aCod[op] = eval("window.document.all.COD" + op + ".value");
    aCan[op] = eval("window.document.all.CAN" + op + ".value");
    aStk[op] = parseInt(eval("window.document.all.STK" + op + ".value"), 10);

    // PRECIO PUBLICO
    aPvp[op] = parseFloat(eval("window.document.all.PVP" + op + ".value"))
    //  nuevoPVP = parseFloat(eval("window.document.all.PVP" + op + ".value") ) * parseFloat(IGV)
    // eval("window.document.all.PVP" + op + ".value=nuevoPVP")
    //alert(aPvp[op])
    // BUSCA SI EXISTE EL CODIGO EN LAS LINEAS DEL DOCUMENTO
    aDisp = Ascan(aCod, aCod[op]);

    if (aDisp.length > 0) {   // cuenta las prendas en las multiples lineas
        stk = aStk[op];
        for (k = 0; k < aDisp.length; k++)
        { stk -= aCan[aDisp[k]] }
        //verifica si queda Stock para poder descargar
        if (stk < 0) {
            alert("Esta venta genera un saldo negativo de : " + stk + " prendas");
            eval("window.document.all.CAN" + op + ".value=''");
            aCan[op] = 0;
            return true;
        }
    }
    // calcula SUB TOTAL DE LA LINEA, sumandole el IGV 
    // PVP YA TIENE EL REDONDEO A DOS DECIMALES Y CANTIDAD ES ENTERO
    aTot[op] = aCan[op] * aPvp[op]
    aTot[op] = Math.round(aTot[op] * 100) / 100
    aIgv[op] = aTot[op] * IGV
    aIgv[op] = Math.round(aIgv[op] * 100) / 100
    //alert(cerea(aIgv[op]), 2)
    aTot[op] += aIgv[op]
    //alert(aTot[op])
    eval("window.document.all.IGV" + op + ".value=aIgv[op]");
    eval("window.document.all.TOT" + op + ".value=cerea(aTot[op],2)");
   // eval("window.document.all.DCT" + op + ".value=0");
    descuento(op)
    TOTALES();

    return true;
}

function SHOW(op) {
    op = parseInt(op, 10)
if (op < 15)
    eval("window.document.all.lin" + op + ".style.display='block'");

}
</script>
</center>
</body>
</html>
