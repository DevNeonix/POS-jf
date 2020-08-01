<%@ Language=VBScript %>
<% Response.Buffer = true %>
<% Session.LCID=2058%>
<% tienda = Request.Cookies("tienda")("pos") %>
<%Usuario = Request.Cookies("tienda")("usr")%>
<%Usuario = Request.Cookies("tienda")("usr")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<!--#include file="COMUN/FUNCIONESCOMUNES.ASP"-->
<!--#include file="COMUN/COMUNqry.ASP"-->
<!--#include file="includes/Cnn.inc"-->
<link rel="Stylesheet" type="text/css" href="ventas.css" />
<% 
dim meses
meses = "Enero, Febrero, Marzo, Abril, Mayo, Junio, Julio, Agosto, Setiembre, Octubre, Noviembre, Diciembre"
ames= split(meses,",")
 ' NUMERO DE VALE A EMITIR
cad = "select correl as docum, serie as sr from documento where cia = '"&tienda&"' and codigo = 'VL' and estado = 'a'"
'response.write(cad)
rs.open cad, cnn
if rs.recordcount <=0 then
    response.write("<center>")
    response.write("<font color=red size =2 face =arial>No existe Tipo de Documento/correlativo registrado para esta tienda ")
    
    response.write("<br>")
    response.write("<br>")
    response.write("<br>")
    response.write("<br>")
    response.write("<br>")
    response.write("<br>")
    response.write("<br>")
    response.write("<b><p onclick='window.close()' style='cursor:pointer'> (x) Cerrar</p>")
    response.write("</center>")
    response.end
else

'RESPONSE.END    
    VL =  cdbl(rs("docum")) + 1
    VL =  right("0000000" + trim(cstr(VL)),7)
    sr = rs("sr")
   ' RESPONSE.WRITE(VL)
end if
rs.close
    
    ' OPERACION QUE GENERA EL VALE
    OPE = REQUEST.QueryString("OPE")
    ' ARRAY DE CANTIDADES
    CNS = REQUEST.QueryString("CNS")
    aCan = split(cns,",")
    ' ARRAY DE CODIGOS
    CDS = REQUEST.QueryString("CDS")
    aCod = split(cds,",")
    ' ARRAY DE PRECIOS
    PRS = REQUEST.QueryString("PRS")
    aPre = split(prs,",")
     ' ARRAY DE numeros de linea
    LNS = REQUEST.QueryString("LNS")
    aLin = split(lns,",")
    ' vendedora que emite el Vale de Compra
    ven = REQUEST.QueryString("VND")
    'response.write(lns)
    'response.write("<br>")
    aa = "'" + replace(lns, ",", "','") + "'"
    'response.write(aa)

cad = "select coddoc, serie+numdoc as docum, cliente, porigv, tienda from movimcab where operacion = '"&ope&"'"
rs.open cad,cnn
'response.write(cad)
cli = rs("cliente")
tori = rs("tienda")
'*****************************************
    'response.end
'*****************************************

rs.movefirst
tipori = rs("coddoc")
numori = rs("docum")
serori = left(rs("docum"),3)
iig  = rs("porigv")
' necesito sumar los parciales del detalle para la cabecera de la Nota de Crédito
cad =   " select sum(precio) as total, sum(descuento) as descuento,     " & _
        " sum(igv) as igv, sum(precio-igv) as subtotal                  " & _
        " from movimdet                                                 " & _
        " where operacion ='"&ope&"' and item in ("&aa&")               "
rs.close
rs.open cad, cnn
rs.movefirst

pvp = cdbl(rs("total"))-cdbl(rs("descuento"))- cdbl(rs("igv"))
dct = rs("descuento")
bru = rs("subtotal")
igv = rs("igv")
net = rs("total")



' Inicia transacción , para que los datos no queden a medias
Cnn.BeginTrans	
rs.close
ll = "select top 1 operacion as ope from movimcab order by 1 desc"
rs.open ll, cnn
' opr  es el nuevo núemro de operación a general
opr = righT("0000000000" + trim(cstr(cdbl(rs("ope"))+1)),10)
RS.CLOSE
rs.open "select nombre from clientes where cliente = '"&cli&"'", cnn
rs.movefirst
nomcli = rs("nombre")
rs.close

rs.open "select descripcion from tiendas where codigo = '"&tori&"'", cnn
rs.movefirst
tienori = rs("descripcion")
rs.close

FEC ="getdate()"


' hacer el insert en la cabecera como operación nueva
	CAD = 	" insert into movimcab  (operacion, tienda,     " & _
			" coddoc,   TIPMOV, SERIE, NUMDOC, FECDOC,      " & _
            " MONEDA, PVP, DESCUENTO, SUBTOT, IGV, TOTAL,   " & _
            " CLIENTE, VENDEDOR, usuario, fecha, ESTADO ,   " & _
            " porigv, DOCORI, NUMORI, SERORI)  values( '"&OPr&"',   " & _
            " '"&tienda&"', 'VL', ' ', '"&SR&"','"&vl&"',   " & _
            " "&FEC&", 'MN', "&PVP&", "&DCT&", "&BRU&",     " & _
            " "&IGV&",  "&NET&", '"&CLI&"', '"&VEN&"',      " & _
            " '"&USUARIO&"',  "&fec&" ,'A', "&IIG&",        " & _
            " '"&tipori&"', '"&numori&"', '"&serori&"');    " 



for t=lbound(aLin) to ubound(aLin)
    ' hacer el insert en el detalle como operación nueva
    LIN = aLin(T)
    ln = right("00"+ trim(cStr(t+1)),2)
    cn = aCan(t)
    cd =  acod(t)
    CAD = CAD + " INSERT INTO MOVIMDET (OPERACION, TIENDA, ITEM,    " & _
                " CODART, VALE, PRECIO, DESCUENTO, IGV, PORDES )   " & _
                "  (SELECT '"&OPR&"', '"&TIENDA&"', '"&ln&"',  " & _
                " codart, sale, precio, descuento, igv, pordes      " & _
                " from movimdet where item = '"&LIN&"' and          " & _
                " operacion = '"&ope&"') ;                          "

    ' UN  VALE NO ACTUALIZA stock
    ' UN  VALE NO marca el FLAG  * al artículo 
next

' ACTUALIZAR EL CORRELATIVO DEL DOCUMENTO EMITIDO

CAD = CAD + " UPDATE DOCUMENTO SET CORREL = '"&vl&"'       " & _
            " WHERE CIA = '"&tienda&"' AND CODIGO = 'VL'; "


' *********************************************************
' HAY QUE ENVIAR UN MAIL A LA TIENDA DE ORIGEN PARA QUE 
' EMITAN LA n/c CUANDO LLEGUE LA PRENDA, INDICANDO 
' DOCUMENTOS DE ORIGEN
' *********************************************************
nume = sr + "-" + vl

doc = "Documento de Origen = " + tipori + " " + serori + "-" + numori
tie = Request.Cookies("tienda")("POS")+" - " +  Request.Cookies("tienda")("TDA") 
rs.open "select correo from tiendas where codigo = '"& tori&"'", cnn
rs.movefirst
recibe = rs("correo")
rs.close
cad = cad + " EXEC SP_correo_vale '"&opr&"', '"&CLI&"', '"&nume&"', '"&TIE&"', '"&doc&"', '"&recibe&"' " 
'************************************************************
cnn.execute cad
'************************************************************
'response.write (cad)

if  err.number <> 0 then
	Response.Write ("No se han podido actualizar los datos soliciatados,  Reintente en unos minutos")
	Cnn.RollbackTrans
	Cnn.Abort
else
	Cnn.CommitTrans	
end if
'response.end


%>
<table width="100%" align="center" >
<tr style="height:90px; vertical-align:top"><td align="right" class="Estilo7">Lima, <%=day(date())%>, de <%=ames(month(date())-1) %> de <%=formatnumber(year(date()),0,,true) %></td></tr>
<tr>
<td align="center" height="95px" valign="top" class="Estilo17">VALE DE COMPRA POR CANJE # <%=VL %></td>
</tr>
<tr>
    <td class="Estilo7">El presente documento se extiende por la Devolucion de prendas adquiridas del cliente <%=ucase(nomcli)%> 
    y servira para canjearlo con la  Nota de Credito que emitira la tienda <%=tienori %> de acuerdo al siguiente detalle :
    </td>
</tr>
</table>
<%
cad = " SELECT DISTINCT m1.ITEM, V1.CODIGO, V1.DESCRI, M1.PRECIO, M1.VALE  " & _
      " FROM movimdet  m1 inner join view_articulos_tienda v1 on v1.codigo " & _
	  "  = m1.codart and m1.tienda = v1.tienda WHERE operacion ='"&OPR&"'  " & _
      " ORDER BY 1                                                         "
RS.OPEN CAD
RS.MOVEFIRST

 %>
<CENTER>
<%=String(130,"-") %>
</CENTER>
<table width="100%" align="center" >
<tr class="Estilo5">
<td>IT</td>
<td>CODIGO</td>
<td>DESCRIPCION</td>
<td>CANT</td>
<td>PRECIO</td>
</tr>
<%DO WHILE NOT RS.EOF %>    
    <tr class="Estilo7">
        <td><%=rs("item")%></td>
        <td><%=rs("codigo")%></td>
        <td><%=rs("descri")%></td>
        <td><%=rs("vale")%></td>
        <td><%=rs("precio")%></td>
       
    </tr>
    <%rs.movenext %>
<%LOOP %>
</table>

<p>
<center>
<img src="images/print.jpg" border="0" id="prn" name="prn" onClick="printa();" style="display:block; cursor:pointer"/>
</center>
</p>

<script language="jscript" type="text/jscript">
    
function printa() {
    document.all.prn.style.display = 'none'
    window.print()
    window.opener.parent.window.parent.window.document.all.body0.src = 'notacredito.asp'
    this.window.close()
}


</script> 

</body>




</html>
