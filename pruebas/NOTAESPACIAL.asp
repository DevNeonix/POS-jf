<%@ Language=VBScript %>
<% Response.Buffer = true %>
<% Session.LCID=2058%>
<% tienda = Request.Cookies("tienda")("pos") %>
<%Usuario = Request.Cookies("tienda")("usr")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<!--#include file="FUNCIONESCOMUNES.ASP"-->
<!--#include file="COMUNqry.ASP"-->
<!--#include file="../includes/Cnn.inc"-->

<%  
    ' NOTA DE CREDITO DE ORIGEN
    ser = REQUEST.QueryString("SER")
    NRO = REQUEST.QueryString("NRO")
    DOC = REQUEST.QueryString("DOC")
    
    CAN = REQUEST.QueryString("CAN")    
    COD = REQUEST.QueryString("COD")
    PRE = cdbl(REQUEST.QueryString("SOL"))
    txt = REQUEST.QueryString("des")
    CLI = REQUEST.QueryString("CLI")
    rs.open "select igv/100 AS IGV from parametros "
    iig = cdbl(rs("IGV")) * 100
    igv = pre * cdbl(rs("IGV"))
    rs.close
    tot = pre + igv
    ' vendedora que emite la NC
    ven = USUARIO
    FEC ="getdate()"
      
   cad = "select serie, correl from documento where cia = '"&tienda&"' and codigo = 'NC'"
        rs.open  cad, cnn
        rs.movefirst    
        ser= rs("serie")
        nro = right("0000000"+cstr(cdbl(rs("correl"))+1),7)
rs.close

ll = "select top 1 operacion as ope from movimcab order by 1 desc"
rs.open ll, cnn
' opr  es el nuevo núemro de operación a general
opr = righT("0000000000" + trim(cstr(cdbl(rs("ope"))+1)),10)
RS.CLOSE




' hacer el insert en la cabecera como operación nueva
	CAD = 	" insert into movimcab  (operacion, tienda,            " & _
			" coddoc,   TIPMOV, SERIE, NUMDOC, FECDOC,             " & _
            " MONEDA, PVP, DESCUENTO, SUBTOT, IGV, TOTAL,          " & _
            " CLIENTE, VENDEDOR, usuario, fecha, ESTADO ,          " & _
            " porigv, DOCORI, NUMORI, serori)                      " & _
            " values('"&OPR&"', '"&tienda&"', 'NC', 'E',           " & _
            " '"&SER&"','"&NRO&"', "&FEC&", 'MN',                  " & _
            " "&PRE&", 0, "&pre&", "&IGV&",  "&tot&",              " & _
            " '"&CLI&"', '"&VEN&"', '"&USUARIO&"',                 " & _
            " "&fec&" ,'A', "&IIG&", '"&DOC&"', '"&NRO&"',         " & _
            " '"&SER&"');                                          " 
    CAD = CAD + " INSERT INTO MOVIMDET (OPERACION, TIENDA, ITEM,   " & _
                " CODART, VALE, PRECIO, DESCUENTO, IGV, PORDES,    " & _
                " texto )   (SELECT '"&OPR&"', '"&TIENDA&"',  1,   " & _
                " 'ESPECIAL', 1, "&TOT&", 0, "&igv&", 0, '"&txt&"') ;         "

  

' ACTUALIZAR EL CORRELATIVO DEL DOCUMENTO EMITIDO

CAD = CAD + " UPDATE DOCUMENTO SET CORREL = '"&NRO&"'       " & _
            " WHERE CIA = '"&tienda&"' AND CODIGO = 'NC';   "

RESPONSE.Write(CAD)

'RESPONSE.END
' Inicia transacción , para que los datos no queden a medias
Cnn.BeginTrans	

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
    cad = '../nota.asp?op='+ '<%=trim(opr)%>'
   window.location.replace(cad)
</script> 
</body>




</html>
