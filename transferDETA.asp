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
<body onload="agranda()">
<table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="0"  cellspacing="1"  border="1" align="center" >
   <tr>
        <td align="center" class="Estilo8">IT</td>
	    <td align="center" class="Estilo8">Codigo</td>
        <td align="center" class="Estilo8">Descripcion</td>
        <td align="center" class="Estilo8">Cant.</td>
        <td align="center" class="Estilo8">STK.</td>
   </tr>
    <%for i=0 to 99 %>
    <tr id="lin<%=i%>" name="lin<%=i%>" <%IF i>0 THEN %>style="display:none"<%END IF %>>
        <td class="Estilo12" valign="middle" align="center"><%=i+1%></td> 
        <td><input id="COD<%=i%>" name="COD<%=i%>" size="20" class="Estilo24" maxlength="25" onchange="carga('<%=i%>');stock('<%=i%>')" value='' ondblclick="hlp('<%=i%>', this)" /></td>
        <td><input id="DES<%=i%>" name="DES<%=i%>" size="60" class="Estilo13" readonly tabindex="-1" value=''/></td>
        <td><input id="CAN<%=i%>" name="CAN<%=i%>" size="8"  class="Estilo133" maxlength="3" onchange="stock('<%=i%>')" value='' /></td>
        <td><input id="STK<%=i%>" name="STK<%=i%>" size="10" class="Estilo133"  readonly tabindex="-1" /></td>
    </tr>
    <%next %>
</table>


<center><img src="images/print.jpg" onclick="graba()" style="cursor:pointer;" id="pr" name="pr"/></center>  

<iframe width="100%" src="" id="mirada" name="mirada"  scrolling="yes" frameborder="3" height="100" align="middle" style="display:none">
</iframe>
</body>


<script type="text/jscript" language="jscript">
    // codigo de producto
    var aCod = new Array()
    // cantidad de venta
    var aCan = new Array()
    // disponible para venta
    var aDis = new Array()
    // stock en almacen
    var aStk = new Array()
    for (t = 0; t < 100; t++) 
    {   aCod[t] = ''
        aCan[t] = 0
    }
    function carga(cn) {

        if (cn > 0) {
            dto = trim(eval("window.document.all.CAN" + (cn - 1) + ".value"));
            if (isNaN(dto)) {
                alert("Favor informar la cantidad de prendas")
                eval("window.document.all.CAN" + (cn - 1) + ".focus()")
                eval("window.document.all.COD" + cn + ".value=''");
                eval("window.document.all.DES" + cn + ".value=''");
                return true;
            }
        }

        dato = eval("window.document.all.COD" + cn + ".value");
        cad = "bake/bakeprendasTran.asp?pos=" + trim(dato) + "&op=" + trim(cn);
        document.all.mirada.src = cad;
        cn = parseInt(cn, 10);
        dd = cn + 1;
        eval("document.all.lin" + dd + ".style.display = 'block'");

    }

    function hlp(op, obj) {
        vobj = trim(obj.value)
        // alert(vobj.length)
        window.open('help/hlpprendasVT.asp?pos=' + trim(vobj) + '&op=' + op)
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
        aStk[op] = parseInt(eval("window.document.all.STK" + op + ".value"), 10);
        // BUSCA SI EXISTE EL CODIGO EN LAS LINEAS DEL DOCUMENTO
        aDisp = Ascan(aCod, aCod[op]);

        if (aDisp.length > 0) {   // cuenta las prendas en las multiples lineas
            stk = aStk[op];
            for (k = 0; k < aDisp.length; k++)
            { stk -= aCan[aDisp[k]] }
            //verifica si queda Stock para poder descargar
            if (stk < 0) {
                alert("Esta operacion genera un saldo negativo de : " + stk + " prendas");
                eval("window.document.all.CAN" + op + ".value=''");
                aCan[op] = 0;
                return true;
            }
        }
        return true;
}

    function graba() {
        // Datos para grabar cabecera y detalle
        cad = 'comun/insertrans.asp?cli=00000000000'  
        cad += '&cod=' + aCod
        cad += '&Can=' + aCan
        cad += '&doc=TR'
        cad += '&mov=S'
        cad += '&ser=' +  Left(ltrim(rtrim(parent.window.frames[0].window.document.all.DOC.value)), 3)
        cad += '&nro=' + Right(ltrim(rtrim(parent.window.frames[0].window.document.all.DOC.value)), 7)
        cad += '&OBS=' + ltrim(rtrim(parent.window.frames[0].window.document.all.NRO.value))
        cad += '&TIP=' + ltrim(rtrim(parent.window.frames[0].window.document.all.TIP.value))
        var opc = "directories=no,height=600,";
        opc = opc + "hotkeys=no,location=no,";
        opc = opc + "menubar=no,resizable=no,";
        opc = opc + "left=0,top=0,scrollbars=yes,";
        opc = opc + "status=no,titlebar=no,toolbar=no,";
        opc = opc + "width=800";
        document.all.pr.style.display='none'
        //alert(cad)
        window.open(cad, '', opc)



    }
</script>
</html>
