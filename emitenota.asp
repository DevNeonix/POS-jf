<%@ Language=VBScript %>
<% Response.Buffer = true %>
<% Session.LCID=2058%>
<% tienda = Request.Cookies("tienda")("pos") %>
<%Usuario = Request.Cookies("tienda")("usr")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<!--#include file="COMUN/FUNCIONESCOMUNES.ASP"-->
<!--#include file="COMUN/COMUNqry.ASP"-->
<!--#include file="includes/Cnn.inc"-->

<%  ' NUMERO DE NOTA DE CREDITO A EMITIR
    NC = REQUEST.QueryString("NC")
    ' OPERACION QUE GENERA LA NOTA DE CREDITO
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
    ' vendedora que emite la NC
    ven = REQUEST.QueryString("VND")
    'response.write(lns)
    'response.write("<br>")
    aa = "'" + replace(lns, ",", "','") + "'"
    'response.write(aa)


'validacion para no generar documentos repetidas'
if rs.state > 0 then
    rs.close
end if

RS.Open "select * from movimcab where coddoc='NC' and SERIE = '"&trim(left(nc,3))&"' AND NUMDOC='"&trim(right(nc,7))&"' and FECDOC >=CONVERT(datetime,'2019-01-10')", Cnn
if rs.recordcount > 0 then
    response.Write("Este documento ya ha sido registrada anteriormente, por favor verifique en reimpreciones.")
    response.end
end if

if rs.state > 0 then
    rs.close
end if


cad = "select coddoc, serie COLLATE Modern_Spanish_CI_AI+numdoc as docum, cliente, porigv from movimcab where operacion  = '"&ope&"'"
rs.open cad,cnn
'response.write(cad)
cli = rs("cliente")
cli = request.QueryString("cli")

rs.movefirst
tipori = rs("coddoc")
numori = rs("docum")
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


FEC ="getdate()"
ser= left(nc,3)
nro = right(nc,7)

' hacer el insert en la cabecera como operación nueva
	CAD = 	" insert into movimcab  (operacion, tienda,             " & _
			" coddoc,   TIPMOV, SERIE, NUMDOC, FECDOC,              " & _
            " MONEDA, PVP, DESCUENTO, SUBTOT, IGV, TOTAL,           " & _
            " CLIENTE, VENDEDOR, usuario, fecha, ESTADO ,           " & _
            " porigv, DOCORI, NUMORI, serori)  values( '"&OPr&"',   " & _
            " '"&tienda&"', 'NC', 'E', '"&SER&"','"&NRO&"',         " & _
            " "&FEC&", 'MN', "&PVP&", "&DCT&", "&BRU&",             " & _
            " "&IGV&",  "&NET&", '"&CLI&"', '"&VEN&"',              " & _
            " '"&USUARIO&"',  "&fec&" ,'A', "&IIG&",                " & _
            " '"&tipori&"', '"&numori&"','"&serori&"');             " 



for t=lbound(aLin) to ubound(aLin)
    ' hacer el insert en el detalle como operación nueva
    LIN = aLin(T)
    ln = right("00"+ trim(cStr(t+1)),2)
    cn = aCan(t)
    cd =  acod(t)
    CAD = CAD + " INSERT INTO MOVIMDET (OPERACION, TIENDA, ITEM,    " & _
                " CODART, ENTRA, PRECIO, DESCUENTO, IGV, PORDES )   " & _
                "  (SELECT '"&OPR&"', '"&TIENDA&"', "&ln&",  " & _
                " codart, sale, precio, descuento, igv, pordes      " & _
                " from movimdet where item = '"&LIN&"' and          " & _
                " operacion = '"&ope&"') ;                          "

    ' hacer ingreso de stock
    cad = cad + " update articulos set stock = stock + "&cn&" where " & _
                " tienda = '"&tienda&"' AND CODIGO = '"&cd&"' ;     "      
    
    ' marcar con el FLAG  * al artículo al que se le ha aplicado una NC
    cad = cad + " update movimdet set flag = '*'        " & _
                " where operacion = '"&ope&"'    and    " &_
                "  item = '"&LIN&"' ;                   "


next
response.write (cad)
' ACTUALIZAR EL CORRELATIVO DEL DOCUMENTO EMITIDO

CAD = CAD + " UPDATE DOCUMENTO SET CORREL = '"&NRO&"'       " & _
            " WHERE CIA = '"&tienda&"' AND CODIGO = 'NC'; "

cnn.execute cad
if  err.number <> 0 then
	Response.Write ("No se han podido actualizar los datos soliciatados,  Reintente en unos minutos")
	Cnn.RollbackTrans
	Cnn.Abort
else
	Cnn.CommitTrans	
end if
'response.end
%>
<script language="jscript" type="text/jscript">
    //cad = 'nota.asp?op='+ '<%=trim(opr)%>'
    cad = 'ppnc.asp?ope='+ '<%=trim(opr)%>'
   window.location.replace(cad)
</script> 
</body>




</html>
