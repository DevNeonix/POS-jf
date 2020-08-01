<%@Language=VBScript%>
<%Response.Buffer = true%>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos")%>

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

<%pos = trim(request.QueryString("pos") )
cAso = trim(request.QueryString("caso") )
ope = trim(request.QueryString("ope") )
pdv = trim(request.QueryString("pdv") )%>

<script language="jscript" type="text/jscript">
function agranda(pos) {
pos = 'trim(request.QueryString("pos") )'
if (trim(pos) == '')
    top.parent.window.document.getElementById('body0').height = 400
   
}
</script>
<body onload="agranda()">
  <%cad = "select serie, correl from documento where cia = '"&tienda&"' and codigo = 'NC'" 
rs.open cad,cnn
if rs.recordcount <= 0 then
    response.write("<center><font color =red size=3 face=arial>No existe correlativo de Notas de cr&eacute;dito para esta tienda </font></center>")
    response.End
end if
rs.movefirst
'response.Write(cad)
'response.end
corr = cdbl(rs("correl"))+ 1

nc = trim(rs("serie")) + "-" + right("0000000"+trim(cstr(corr)),7)
rs.close


' la primera vez que se muestra la pantalla no tiene querystring y se deb mostratr en blanco la pantalla
if len(trim(pos)) <=0 then  response.End 

'response.write(caso)

 cad =  " select distinct m1.ITEM AS IT, V1.CODIGO, V1.DESCRI, M1.PRECIO, M1.SALE AS CANT   " & _
        " , v1.tienda  as TDA, operacion as OPE                                             " & _
        " from movimdet  m1 inner join view_articulos_tienda v1 on v1.codigo = m1.codart    " & _
        " and m1.tienda = v1.tienda                                                         " & _
        " where operacion ='"&ope&"'                                                        " & _
        " and m1.flag = ''                                                                  " & _
        " order by 1                                                                        "
' FAC 20130222 mostrar todos los item del documento      " where operacion ='"&ope&"' and m1.codart = '"&pos&"'                              " & _
' response.Write(cad)
 
    RS.OPEN CAD,CNN

 If rs.recordcount<=0 then   
    RESPONSE.Write("SIN REGISTROS....")
    response.End  
 END IF      
 rs.movefirst
 %>
<table align="center" cellpadding="2" cellspacing="1" bordercolor='<%=application("color2") %>' border="0" >
   <tr> 
        <%for j=0 to rs.fields.count-1 %>
            <td class="Estilo8">&nbsp;<%=RS.FIELDS(j).NAME%>&nbsp;</td>
        <%next %>
         <td class="Estilo8">&nbsp;Chk&nbsp;</td>
   </tr>
    <%I=1 %>
    <%do while not rs.eof %>
        <tr bgcolor="<%IF I MOD 2 = 0 THEN RESPONSE.WRITE(application("color1")) ELSE RESPONSE.WRITE(application("color2"))%>" class="Estilo19">
            <%for j=0 to rs.fields.count-3 %>
            <td id='COL<%=i%><%=j%>' name='COL<%=i%><%=j%>' onclick="alert(this.innerText)" ><%=RS.FIELDs.item(j)%></td>
            <%next %>
            <%for h=j to rs.fields.count-1 %>
            <td align="right" id='COL<%=i%><%=h%>' name='COL<%=i%><%=h%>'><%=trim(RS.FIELDs.item(h))%>&nbsp;&nbsp;</td>
            <%next %>
            <td align="center"><input id='chk<%=i%>' name='chk<%=i%>' type="checkbox" value=0 /></td>
       </tr> 
       <%I= I + 1%>
       <%rs.movenext %>
    <%loop%>

</table>
<iframe width="100%"  src="" id="cok" name="cok"  height="100" style="display:none;" align="middle">
</iframe>

<table align="center" cellpadding="0" cellspacing="1" bordercolor='<%=application("color2") %>' border="0" >
    </tr>
       <td class="Estilo8">VENDEDOR :&nbsp;</td>
       <td align="left"  bordercolor="<%=Application("color1")%>" ><input id="ven" name="ven" class="style2" value='' /></td>  
       <td align="left"  bordercolor="<%=Application("color1")%>" colspan="2" ><input id="des" name="des" class="style2" value='' readonly tabindex="-1" /></td>  
    </tr>
    </tr>
       <td class="Estilo8">CLIENTE :&nbsp;</td>
       <td align="left"><input type="text" name="CLI" id="CLI" value="" class="style2" onchange="cliente(this.value)"  size="20" maxlength="11"></td>
       <td colspan="2" >
        <input type="text" name="DES" id="DES" value="" maxlength="50" class="style2" readonly tabindex="-1">
        <input type="text" name="DIR" id="DIR" value="" class="style2" style="display:none" readonly tabindex="-1"></td>
    </tr>
    <tr><td align="center" class="style1" colspan="3"><img src="images/print.jpg"  id="prn" name="prn" width="30" style="cursor:pointer" onclick="imprime()"/></td>
    </tr>
</table>
<iframe src="" allowScriptAccess='always'  id="mirada" name="mirada" style="display:none"></iframe>
</body>
<%rs.close
set rs = nothing
set cnn = nothing
%>

<script language=jscript type="text/jscript">
    var aCns = Array()
    var aCds = Array()
    var aPrs = Array()
    var aLns = Array()

function bake() {
    document.all.cok.src = 'bake/bakevende.asp?pos=' + ltrim(rtrim(document.all.ven.value))
    //document.getElementById('prn').style.display = 'block'
}
function cliente(dato) {

    ss = document.all.CLI.value
    cad = "bake/bakecliente.asp?pos=" + trim(dato)
    //   document.all.mirada.style.display='block'
    document.all.mirada.src = cad

    if (ss.length < 8) {
        document.all.CLI.value = strzero(trim(document.all.CLI.value), 11)
        dato = strzero(dato, 11)
    }

}
function imprime() {
    
    bake()
    if (ltrim(rtrim(document.all.ven.value)) == '')
        return true;
  
    i = parseInt('<%=cint(i) %>', 10)   
    m=0

    for (z = 1; z < i; z++)
    {  
        mm = eval("document.all.chk" + z + ".checked")

        if (mm == true) 
        {   // cantidad
            aCns[m] = eval("document.all.COL" + z + "4.innerText")
            // codigo del articulo
            aCds[m] = eval("document.all.COL" + z + "1.innerText")
            // precio del articulo
            aPrs[m] = eval("document.all.COL" + z + "3.innerText")
            // numero de lines
            aLns[m++] = eval("document.all.COL" + z + "0.innerText")
        }
       
    }

    if (m == 0) 
    {   alert("No ha seleccionado ninguna linea para aplicar la Nota de Credito")
        return true;
    }


    caso = parseInt('<%=trim(caso)%>', 10)

    docum = '<%=nc%>'

    if (caso > 2) {
        cad = 'El producto no fue vendido en esta tienda\n¿ Desea imprimir un Vale de Compra por el valor?'
        if (confirm(cad) == false) {
            alert("Impresion Cancelada ");
            return true;
        }
        else {
            ken = 'emitevale.asp?NC=' + docum + '&cns=' + aCns
        }

    }
    else {
        cad = '¿ Desea imprimir la Nota de Credito # ' + docum + '?'
        if (confirm(cad) == false) {
            alert("Impresion Cancelada ");
            return true;
        }
        else
            ken = 'emitenota.asp?NC=' + docum + '&cns=' + aCns
    }

    var opc = "directories=no,height=500,";
    opc = opc + "hotkeys=no,location=no,";
    opc = opc + "menubar=no,resizable=no,";
    opc = opc + "left=100,top=100,scrollbars=yes,";
    opc = opc + "status=no,titlebar=no,toolbar=no,";
    opc = opc + "width=700";
    OPE = "<%=TRIM(OPE) %>"

    ven1 = ken

    ven1 += '&cli=' + document.all.CLI.value
    ven1 += '&cds=' + aCds
    ven1 += '&prs=' + aPrs
    ven1 += '&lns=' + aLns
    ven1 += '&ope=' + OPE
    ven1 += '&VND=' + trim(document.all.ven.value)

    window.open(ven1, "NOTA_DE_CREDITO", opc)
}
 
 </script>
</html>
