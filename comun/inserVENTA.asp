<%@LANGUAGE="VBSCRIPT"  CODEPAGE="65001"%>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<%Usuario = Request.Cookies("tienda")("usr")%>
<%tda     = Request.Cookies("tienda")("pos")%>
<%sol    = Request.Cookies("sol")%>
<%vis    = Request.Cookies("vis")%>
<%mas    = Request.Cookies("mas")%>
<%dol    = Request.Cookies("dol")%>
<%ven    = Request.Cookies("ven")%>
<%cre    = Request.Cookies("cre")%>
<%sr1    = Request.Cookies("sr1")%>
<%ncr    = Request.Cookies("ncr")%>
<!--#include file="../includes/Cnn.inc"-->
<!--#include file="../comun/funcionescomunes.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<body>
<% 
'*********************** OJO **********************
'REVISAR EL LARGO DE LAS COLUMNAS de la tabla
'*********************** OJO **********************
' codigo de articulos
cod = Request.QueryString("cod")
aCod = split(cod,",")
' cantidad vendida
can = Request.QueryString("can")
aCan = split(can,",")
' precio neto = PVP - dct 
pvt = Request.QueryString("pvt")
aPvt = split(pvt,",")
' descuento en soles
des = Request.QueryString("des")
aDes = split(des,",")
'porcentaje del decuento
por = Request.QueryString("por")
aPor = split(por,",")
'monto del Igv
igg = Request.QueryString("igg")
aIgv = split(igg,",")
'% Igv
IIG = Request.QueryString("PorI")
doc = request.QueryString("doc")
tip = request.QueryString("mov")
ser = request.QueryString("ser")

nro = request.QueryString("nro")
FEC ="getdate()"
mon = "MN"
pvp = request.QueryString("pvp")
dct = request.QueryString("dct")
bru = request.QueryString("bru")
igv = request.QueryString("igv")
isc = request.QueryString("isc")
is2 = request.QueryString("is2")
net = request.QueryString("net")
cli = request.QueryString("cli")
qIsc = request.QueryString("qIsc")




'validacion para no generar boletas repetidas'


if rs.state > 0 then
    rs.close
end if

RS.Open "select * from movimcab where SERIE = '"&trim(ser)&"' AND NUMDOC='"&trim(nro)&"'", Cnn
if rs.recordcount > 0 then
    response.Write("Esta Boleta ya ha sido registrada anteriormente, por favor verifique en reimpreciones.")
    response.end
end if




CAD = "EXEC SP_TIPODECAMBIO"
if rs.state > 0 then
    rs.close
end if

RS.Open CAD, Cnn
TCAMBIO = RS("BAJO")
RS.CLOSE
' Inicia transacci�n , para que los datos no queden a medias
Cnn.BeginTrans	
rs.open "select top 1 operacion as ope from movimcab order by 1 desc;", cnn
if rs.recordcount <=0 then 
    ope= "0000000000"
else
    ope = righT("0000000000" + trim(cstr(cdbl(rs("ope"))+1)),10)
end if
RS.CLOSE
	CAD = 	" insert into movimcab  (operacion, tienda, " & _
			" coddoc,   TIPMOV,   " & _
            " SERIE, NUMDOC, FECDOC, MONEDA, PVP,       " & _
			" DESCUENTO, SUBTOT, IGV, TOTAL, CLIENTE,   " & _
            " VENDEDOR, usuario, fecha, ESTADO , porigv," & _
            " isc, iscqty)                              " & _
            " values( '"&OPE&"','"&TDA&"', '"&DOC&"',   " & _
            " '"&TIP&"', '"&SER&"', '"&NRO&"', "&FEC&", " & _
            " '"&MON&"', "&PVP&", "&DCT&", "&BRU&",     " & _
            " "&IGV&",  "&NET&", '"&CLI&"', '"&VEN&"',  " & _
            " '"&USUARIO&"',    "&fec&" ,'A', "&IIG&",  " & _
            " "&isc&", "&qIsc&" ); " 

RESPONSE.WRITE("<br>")
RESPONSE.WRITE(CAD)
RESPONSE.WRITE("<br>")
RESPONSE.WRITE("<br>")
RESPONSE.WRITE(aCod(0))
RESPONSE.WRITE(aCod(1))
rs.open "select igv from parametros", cnn
rs.movefirst
GGG = cdbl(rs("igv"))/100
RS.CLOSE
FOR I=0 TO 14
    IF (LEN(TRIM(ACOD(I))) > 0  ) THEN
        LIN = RIGHT("00" + TRIM(CSTR(I+1)),2)
        pvt = round(aPVT(i),2)
        DTO = ROUND(AdES(I),2)
        IGG = ROUND(aIgv(i),2)
        POR = ROUND(aPor(i),2)
        ' LINEA DE DETALLE DE DOCUMENTO DE VENTA
        CAD = CAD + " INSERT INTO MOVIMDET (OPERACION, TIENDA, ITEM,    " & _
                    " CODART, SALE, PRECIO, DESCUENTO, IGV, PORDES )    " & _
                    " VALUES ('"&OPE&"', '"&TDA&"', '"&LIN&"',          " & _
                    " ltrim(rtrim('"&aCod(I)&"')),  "&aCan(i)&", "&PVT&", "&DTO&",    " & _
                    " "&IGG&" , "&POR&") ;        "
        ' ACTUALIZA EL STOCK DEL ARTICULO
        CAD = CAD + " UPDATE ARTICULOS SET STOCK = STOCK - "&aCan(i)&" ,    " & _
                    " usuario = '"&ven&"', fecha = "&fec&", estado = 'A'    " & _
                    " where TIENDA = '"&tda&"' and codigo = ltrim(rtrim('"&aCod(i)&"')) ; " 
    ELSE
        EXIT FOR
    END IF
NEXT    
' ACTUALIZAR EL CORRELATIVO DEL DOCUMENTO EMITIDO

CAD = CAD + " UPDATE DOCUMENTO SET CORREL = '"&NRO&"'       " & _
            " WHERE CIA = '"&TDA&"' AND CODIGO = '"&DOC&"'; "

' INGRESAR LOS VALORES DEL MOVIMIENTO DE CAJA
LN = 1
IF isnumeric(sol) THEN
    if cdbl(sol) > 0  then
        CAD = CAD + " INSERT INTO CAJA (TIENDA, OPERACION, LIN, MONEDA, TIPO, MONTO,    " & _
                    " TCAMBIO, USUARIO, FECHA, ESTADO ) VALUES ('"&TDA&"', '"&OPE&"',   " & _
                    " "&LN&", 'MN', 'SOL', "&SOL&", "&TCAMBIO&", '"&VEN&"', "&FEC&", 'A'); "
        LN =  LN + 1
    end if
END IF

IF isnumeric(dol) THEN
    if  cdbl(dol) > 0  then
        CAD = CAD + " INSERT INTO CAJA (TIENDA, OPERACION, LIN, MONEDA, TIPO, MONTO,    " & _
                    " TCAMBIO, USUARIO, FECHA, ESTADO ) VALUES ('"&TDA&"', '"&OPE&"',   " & _
                    " "&LN&", 'US', 'DOL', "&DOL&", "&TCAMBIO&", '"&VEN&"', "&FEC&", 'A'); "
        LN =  LN + 1
    end if
END IF
'response.Write( isnumeric(sol) )
IF  isnumeric(vis)  THEN
    if cdbl(vis) > 0  then
        CAD = CAD + " INSERT INTO CAJA (TIENDA, OPERACION, LIN, MONEDA, TIPO, MONTO,    " & _
                    " TCAMBIO, USUARIO, FECHA, ESTADO ) VALUES ('"&TDA&"', '"&OPE&"',   " & _
                    " "&LN&", 'MN', 'VIS', "&vis&", "&TCAMBIO&", '"&VEN&"', "&FEC&", 'A'); "
        LN =  LN + 1
    end if
END IF

IF  isnumeric(mas) THEN
    if cdbl(mas) > 0 then
        CAD = CAD + " INSERT INTO CAJA (TIENDA, OPERACION, LIN, MONEDA, TIPO, MONTO,    " & _
                    " TCAMBIO, USUARIO, FECHA, ESTADO ) VALUES ('"&TDA&"', '"&OPE&"',   " & _
                    " "&LN&", 'MN', 'MAS', "&MAs&", "&TCAMBIO&", '"&VEN&"', "&FEC&", 'A'); "
        LN =  LN + 1
    end if
END IF


IF  isnumeric(CRE) THEN
    if cdbl(CRE) > 0 then
        NOTA = TRIM(SR1) + "-" + TRIM(NCR)
        CAD = CAD + " INSERT INTO CAJA (TIENDA, OPERACION, LIN, MONEDA, TIPO, MONTO,    " & _
                    " TCAMBIO, USUARIO, FECHA, ESTADO, NOTA ) VALUES ('"&TDA&"', '"&OPE&"',   " & _
                    " "&LN&", 'MN', 'NCR', "&CRE&", "&TCAMBIO&", '"&VEN&"', "&FEC&", 'A','"&NOTA&"'); "
        LN =  LN + 1
    end if
END IF




RESPONSE.WRITE("GUARDANDO Y ENVIANDO A SUNAT TU VENTA POR FAVOR ESPERA...")
RESPONSE.WRITE("<br>")
RESPONSE.WRITE("<br>")


'RESPONSE.END




Cnn.Execute(CAD)
if  err.number <> 0 then
	Response.Write ("No se han podido actualizar los datos soliciatados,  Reintente en unos minutos")
	Cnn.RollbackTrans
	Cnn.Abort%>
    <script language="javascript" type="text/jscript">
        delCookie("dol")
        delCookie("sol")
        delCookie("vis")
        delCookie("mas")
       // window.close()
    </script>
<%else
	Cnn.CommitTrans	
end if

Cnn.Close	
set Cnn = Nothing
SET RS = Nothing	

%>

<script language="javascript" type="text/jscript">


doc = '<%=trim(doc) %>'
   
delCookie("dol")
delCookie("sol")
delCookie("vis")
delCookie("mas")
this.window.focus()
var opc = "directories=no,height=600,";
opc = opc + "hotkeys=no,location=no,";
opc = opc + "menubar=no,resizable=yes,";
opc = opc + "left=0,top=0,scrollbars=yes,";
opc = opc + "status=no,titlebar=yes,toolbar=yes,";
opc = opc + "width=900";

if (doc == 'BL') {
    //this.window.focus()
    //window.location.replace('../boleta.asp?ope=' + '<%=OPE %>','BOLETA', opc)
    window.location.replace('../ppboleta.asp?ope=' + '<%=OPE %>','BOLETA', opc)
}
else {
    //this.window.focus()
    //window.location.replace('../factura.asp?ope=' + '<%=OPE %>', 'Factura', opc)
    window.location.replace('../ppfactura.asp?ope=' + '<%=OPE %>', 'Factura', opc)
}


setTimeout(function(){
window.opener.top.document.getElementById("body0").src=window.opener.top.document.getElementById("body0").src
},1500);
                                                                                    
</script>
</BODY>
</HTML>
