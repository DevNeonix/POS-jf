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
<title>Untitled Document</title>
</head>
<%rs.open "select ISC from parametros",cnn
ISC = cdbl(rs("isc"))

rs.close %>
<body>
<table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="0"  cellspacing="1"  border="1" align="center" >
   <tr>
        <td align="center" class="Estilo8">IT</td>
	    <td align="center" class="Estilo8">Codigo</td>
        <td align="center" class="Estilo8">Descripcion</td>
        <td align="center" class="Estilo8">Cant.</td>
        <td align="center" class="Estilo8">STK.</td>
        <td align="center" class="Estilo8">PVP</td>
        <td align="center" class="Estilo8">Dcto.</td>
        <td align="center" class="Estilo8">PDct</td>
        <td align="center" class="Estilo8">IGV</td>
        <td align="center" class="Estilo8">Total</td>
        <td align="center" class="Estilo8">ISC</td>
   </tr>
    <%for i=0 to 11 %>
    <tr id="lin<%=i%>" name="lin<%=i%>" <%IF i>0 THEN %>style="display:none"<%END IF %>>
        <td class="Estilo12" valign="middle" align="center"><%=i+1%></td> 
        <td><input id="COD<%=i%>" name="COD<%=i%>" size="20" class="Estilo24" maxlength="25" onchange="carga('<%=i%>');stock('<%=i%>')" value='' ondblclick="hlp('<%=i%>', this)" /></td>
        <td><input id="DES<%=i%>" name="DES<%=i%>" size="60" class="Estilo13" readonly tabindex="-1" value=''/></td>
        <td><input id="CAN<%=i%>" name="CAN<%=i%>" size="8"  class="Estilo133" maxlength="3" onchange="stock('<%=i%>')" value='' /></td>
        <td><input id="STK<%=i%>" name="STK<%=i%>" size="10" class="Estilo133" maxlength="6" readonly tabindex="-1" /></td>
        <td><input id="PVP<%=i%>" name="PVP<%=i%>" size="10" class="Estilo133" maxlength="6"  onchange="stock('<%=i%>')"  readonly tabindex="-1" value='' /></td>
       
        <td>
         <select type="text" id="DCT<%=i%>" name="DCT<%=i%>" class="Estilo13" onchange="descuento('<%=i%>')" disabled >
                       <option value="0"></option> 
                       <%CAD =	" select * from descuentos order by 2          " 
                       if rs.state > 0 then rs.close
                       rs.open CAD ,cnn
                       rs.movefirst
                       do while not rs.eof%>
                       <option value="<%=rs("valor") %>"><%=rs("descripcion") %></option>
                        <%rs.movenext
                       loop%>
                  </select>
                  <%RS.CLOSE %>
        
       </td>
        <td><input id="PDT<%=i%>" name="PDT<%=i%>" size="10" class="Estilo133" maxlength="6" readonly tabindex="-1" value=''/></td>
        <td><input id="IGV<%=i%>" name="IGV<%=i%>" size="10" class="Estilo133" maxlength="6" readonly tabindex="-1" value='' /></td>
        <td><input id="TOT<%=i%>" name="TOT<%=i%>" size="12" class="Estilo133" readonly tabindex="-1" value="" /></td>
        <td><input id="ISC<%=i%>" name="ISC<%=i%>" size="12" class="Estilo133" readonly tabindex="-1" value="" /></td>
        <td style="display:block"><input id="III<%=i%>" name="III<%=i%>" size="12" class="Estilo133" readonly tabindex="-1" value="" /></td>
    </tr>
    <%next %>
</table>
<iframe width="100%" src="" id="mirada" name="mirada"  scrolling="yes" frameborder="3" height="100" align="middle" style="display:none">
</iframe>
</body>

<%rs.open "select igv from parametros", cnn
rs.movefirst
igv = rs("igv")%>

<script type="text/jscript" language="jscript">
    var ISC = parseFloat('<%=ISC%>')
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
    // ISC DE CADA LINEA
    var aIsc = new Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    // Iii DE CADA LINEA (si esta afecto o no al isc)
    var aIii = new Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
function CLIENTE() {
    if (trim(parent.window.frames[0].window.document.all.CLI.value) == '') {
        alert("Por favor ingresar el cliente");
        parent.window.frames[0].window.document.all.CLI.focus();
        return false;
    }
    return true;
}

function carga(cn) {
    if (CLIENTE() == false) {
        eval("window.document.all.COD" + cn + ".value=''");
        return false;
    }
    cn = parseInt(cn, 10)
    //alert(cn)

    if ((ltrim(rtrim(eval("document.all.COD" + cn + ".value.toUpperCase()")))) == 'SALDO00000') {
        //alert()
        eval("document.all.PVP" + cn + ".readOnly=false")
       // alert(eval("document.all.COD" + cn + ".tabIndex"))
        eval("document.all.PVP" + cn + ".tabIndex=0")
        eval("document.all.DES" + cn + ".value='Saldos'")
        eval("document.all.COD" + cn + ".focus()")
       
    }
    if (cn > 0) {
        dto = trim(eval("window.document.all.CAN" + (cn - 1) + ".value"));
        if (isNaN(dto)) {
            alert("Favor informar la cantidad de prendas")
            eval("window.document.all.CAN" + (cn - 1) + ".focus()")
            eval("window.document.all.COD" + cn + ".value=''");
            eval("window.document.all.DES" + cn + ".value=''");
            
        }
    }

   // if ((ltrim(rtrim(eval("document.all.COD" + cn + ".value.toUpperCase()")))) != 'SALDO00000') {
        dato = eval("window.document.all.COD" + cn + ".value");
        cad = "bake/bakeprendasVT.asp?pos=" + trim(dato) + "&op=" + cn;
        //alert(cad)
        // document.all.mirada.style.display='block'
        document.all.mirada.src = cad;
    //}
    cn = parseInt(cn, 10);
    dd = cn + 1;
    eval("document.all.lin" + dd + ".style.display = 'block'");
 
    
}

function hlp(op, obj) {
    vobj= trim(obj.value)
   // alert(vobj.length)
    window.open('help/hlpprendasVT.asp?pos='+trim(vobj)+'&op='+op)
    return true;
}
function stock(op) {
    aDisp = new Array();
    op = parseInt(op, 10);
    xx = parseInt(toInt(eval("window.document.all.CAN" + op + ".value")), 10)

    if (isNaN(xx)) {
        xx = 0
    }
    
    eval("window.document.all.CAN" + op + ".value=xx");
    aCod[op] = eval("window.document.all.COD" + op + ".value");
    aCan[op] = eval("window.document.all.CAN" + op + ".value");
    aIii[op] = eval("window.document.all.III" + op + ".value");
    aStk[op] = parseInt(eval("window.document.all.STK" + op + ".value"), 10);

    if (parseInt(aIii[op], 10) == 1)
    { aIsc[op] = parseFloat(ISC) * parseInt(eval("window.document.all.CAN" + op + ".value"), 10); }
    else
    { aIsc[op] = 0}
    
    // PRECIO PUBLICO
    aPvp[op] = parseFloat(eval("window.document.all.PVP" + op + ".value") )


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
    aTot[op] = Math.round(aTot[op]*100)/100
    aIgv[op] =  aTot[op] * IGV
    aIgv[op] = Math.round(aIgv[op] * 100) / 100
    //alert(cerea(aIgv[op]), 2)
    aTot[op] += aIgv[op]
    //alert(aTot[op])
    eval("window.document.all.IGV" + op + ".value=aIgv[op]");
    eval("window.document.all.TOT" + op + ".value=cerea(aTot[op],2)");
    var xd = "ISC" + op.toString()
    var tt = parseInt(aCan[op],10) * parseFloat(ISC) * parseInt(aIii[op],10)
    document.getElementById(xd).value=  cerea(tt,2); 
    
    //eval("window.document.all.ISC" + op + ".value=cerea(parseFloat(aCan[op]) * parseFloat(ISC),2)"); 
    eval("window.document.all.DCT" + op + ".disabled=false");
    descuento(op)
    TOTALES();
    return true;
}

function descuento(op) {

    var IGV = parseFloat('<%=cdbl(igv)/100%>')

    op = parseInt(op, 10)
    aDct[op] = parseFloat(eval("window.document.all.DCT" + op + ".value"))
    aPor[op] = parseFloat(eval("window.document.all.DCT" + op + ".value"))
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
    //alert(IGV)
    aIgv[op] = aTot[op] * IGV
    aIgv[op] = Math.round(aIgv[op] * 100) / 100
    //alert(cerea(aIgv[op]), 2)
    //alert(aTot[op])
    aTot[op] += aIgv[op]
    //alert(aTot[op])
    eval("window.document.all.IGV" + op + ".value=aIgv[op]");
    eval("window.document.all.TOT" + op + ".value=cerea(aTot[op],2)");
       
    TOTALES();
    return true
}
 function TOTALES() {  // suma prendas
var  CAN = 0
var  TOTAL = 0
var  DCT = 0
var  MON = 0
var  PVP = 0
var  IGV = 0
var  isk = 0
     for (m = 0; m < 15; m++) {       
        CAN += parseInt(aCan[m], 10)
        TOTAL += parseInt(aTot[m]*100,10)
        MON += parseInt(aMon[m]*100,10)
        DCT += parseInt(aDes[m]*100,10)
        IGV += parseInt(aIgv[m] * 100, 10)
        isk += parseFloat(aIsc[m])
        PVP += parseInt(aCan[m] * aPvp[m] * 100, 10)
        //alert(isk)
     }
     TOTAL = TOTAL / 100
     MON = MON / 100
     DCT = DCT / 100
     IGV = IGV / 100
     PVP = PVP / 100
     isk = isk 
   //  alert(isk)
     parent.window.frames[2].window.document.all.canti.value = CAN

     if (parent.window.frames[0].window.document.all.miRadio[0].checked == true) 
         boleta = PVP - MON
     else
         boleta = PVP
     //FACTURA
     parent.window.frames[2].window.document.all.pvp.value      = cerea(PVP, 2)
     parent.window.frames[2].window.document.all.dcto.value     = cerea(MON, 2)
     parent.window.frames[2].window.document.all.bruto.value    = cerea(PVP-MON,2)
     parent.window.frames[2].window.document.all.igv.value      = cerea(IGV, 2)
     parent.window.frames[2].window.document.all.isc.value      = cerea(isk, 2)
     parent.window.frames[2].window.document.all.tota.value     = cerea(parseFloat(TOTAL)+parseFloat(isk), 2)
     // BOLETA
     // igv del descuento para la impresion de la boleta
     dt_igv =parseFloat('<%=rs("igv")%>')
     dt_igv = MON * (dt_igv/100)
     parent.window.frames[2].window.document.all.CAN.value      = CAN
     parent.window.frames[2].window.document.all.PVP.value      = cerea(TOTAL+MON+dt_igv, 2)
     parent.window.frames[2].window.document.all.DCT.value = cerea(MON + dt_igv, 2)
     parent.window.frames[2].window.document.all.ISC.value = cerea(parseFloat(isk), 2)
     parent.window.frames[2].window.document.all.TOT.value = cerea(parseFloat(TOTAL) + parseFloat(isk), 2)
     
     maximus = parseFloat(cerea(MON+dt_igv, 2))


     if ( parseInt(maximus,10) > 150) {
         parent.window.frames[2].window.document.all.dcto.className = 'Estilo8'
         parent.window.frames[2].window.document.all.DCT.className  = 'Estilo8';
     }
     else {
         parent.window.frames[2].window.document.all.dcto.className = 'Estilo14'
         parent.window.frames[2].window.document.all.DCT.className  = 'Estilo14';
     }
 }

 function graba() {
     cad = ''
// alert("entro")
     if (parent.window.frames[0].window.document.all.miRadio[0].checked == true) {
         doc = 'BL'
         ser = Left(parent.window.frames[0].window.document.all.BOL.value, 4)
         nro = strzero(parent.window.frames[0].window.document.all.BOLDOC.value,7)
         //nro = Right(parent.window.frames[0].window.document.all.BOL.value,(nro.length-5))
     }
     else {
         doc = 'FC'
         ser = Left(parent.window.frames[0].window.document.all.FAC.value, 4)
         nro = strzero(parent.window.frames[0].window.document.all.FACDOC.value,7)
         //nro = Right(parent.window.frames[0].window.document.all.FAC.value, (nro.length - 5))
     }
     //   alert()
     // ARREGLA LOS DECIMALES PARA LA GRABACION
     qtyISC = 0
     for (K = 0; K < 15; K++) {
         aCan[K] = parseInt(aCan[K], 10)
         aTot[K] = parseInt(aTot[K] * 100, 10)/100
         aMon[K] = parseInt(aMon[K] * 100, 10)/100
         aDct[K] = parseInt(aDes[K] * 100, 10)/100
         aIgv[K] = parseInt(aIgv[K] * 100, 10)/100
         aPvp[K] = parseInt(aPvp[K] * 100, 10) / 100
         aIsc[K] = parseInt(aIsc[K] * 100, 10) / 100
         if (parseInt(aIii[K], 10) == 1)
         { qtyISC += parseInt(aCan[K], 10) }
     }


     // Datos para grabar cabecera y detalle

     cad = 'comun/inserconti.asp?cli=' + ltrim(rtrim(parent.window.frames[0].window.document.all.CLI.value))
     cad += '&cod=' + aCod
     cad += '&Can=' + aCan
     // neto
     cad += '&PVT=' + aTot
     cad += '&doc=' + doc
     cad += '&mov=S'
     cad += '&ser=' + ser
     cad += '&nro=' + nro
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
     cad += '&PVP=' + parent.window.frames[2].window.document.all.pvp.value
     cad += '&dct=' + parent.window.frames[2].window.document.all.dcto.value
     cad += '&bru=' + parent.window.frames[2].window.document.all.bruto.value
     cad += '&igv=' + parent.window.frames[2].window.document.all.igv.value
     cad += '&net=' + parent.window.frames[2].window.document.all.tota.value
     cad += '&isc=' + parent.window.frames[2].window.document.all.isc.value
 //    cad += '&is2=' + parent.window.frames[2].window.document.all.ISC.value
    // igv = '<%=request.QueryString("igv")%>'
     cad += '&qISC=' + qtyISC
     
     var opc = "directories=no,height=600,";
     opc = opc + "hotkeys=no,location=no,";
     opc = opc + "menubar=no,resizable=no,";
     opc = opc + "left=0,top=0,scrollbars=yes,";
     opc = opc + "status=no,titlebar=no,toolbar=no,";
     opc = opc + "width=800";

     //  alert('<%=Request.Cookies("sr1")%>')

     alert(cad)
     window.open(cad, '', opc) 
     
     

 }
</script>
</html>
