<%@LANGUAGE="VBSCRIPT"  CODEPAGE="65001"%>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<%Usuario = Request.Cookies("tienda")("usr")%>


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
'*********************** OJO **********************+
' codigo de la tienda a la que pertenece la operacion
tda = request.QueryString("tda")
ope = request.QueryString("ope")
' codigo de articulos
cod = Request.QueryString("cod")
response.Write(cod)
RESPONSE.WRITE("<br>")
RESPONSE.WRITE("<br>")

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
net = request.QueryString("net")
cli = request.QueryString("cli")
CAD = "EXEC SP_TIPODECAMBIO"
RS.Open CAD, Cnn
TCAMBIO = RS("BAJO")
RS.CLOSE
' Inicia transacción , para que los datos no queden a medias
Cnn.BeginTrans	
CAD = "SELECT *  FROM MOVIMDET  where operacion = '"&ope&"' ; "
RS.OPEN CAD, CNN
if rs.recordcount >0 then rs.movefirst
cad= ""

' RESTAURA EL STOCK AL ALMACEN DE LA TIENDA (CANTIDAD ANTERIOR)
do while not rs.eof
    cod = rs("codart")
    can = cdbl(rs("sale"))
    cad = cad + " update articulos set stock = stock + "&can&" where tienda = '"&tda&"' and codigo = '"&cod&"' ; "
    rs.movenext
loop

cad =  cad +  " delete movimdet where operacion = '"&ope&"'; "
LIN=1
FOR I=LBOUND(ACOD) TO UBOUND(ACOD)
    IF CINT(ACAN(I)) > 0   THEN
        LIN = RIGHT("00" + TRIM(CSTR(LIN)),2)
        pvt = round(aPVT(i),2)
        DTO = ROUND(AdES(I),2)
        IGG = ROUND(aIgv(i),2)
        POR = ROUND(aPor(i),2)
        ' LINEA DE DETALLE DE DOCUMENTO DE VENTA
        CAD = CAD + " INSERT INTO MOVIMDET (OPERACION, TIENDA, ITEM,    " & _
                    " CODART, SALE, PRECIO, DESCUENTO, IGV, PORDES )    " & _
                    " VALUES ('"&OPE&"', '"&TDA&"', '"&LIN&"',          " & _
                    " '"&aCod(I)&"',  "&aCan(i)&", "&PVT&", "&DTO&",    " & _
                    " "&IGG&" , "&POR&") ;        "
        ' ACTUALIZA EL STOCK DEL ARTICULO
        CAD = CAD + " UPDATE ARTICULOS SET STOCK = STOCK - "&aCan(i)&" ,    " & _
                    " usuario = '"&ven&"', fecha = "&fec&", estado = 'A'    " & _
                    " where TIENDA = '"&tda&"' and codigo = '"&aCod(i)&"' ; " 

    END IF
    LIN = LIN + 1
NEXT    

RESPONSE.WRITE(CAD)
RESPONSE.WRITE("<br>")
RESPONSE.WRITE("<br>")

'RESPONSE.END




Cnn.Execute(CAD)
if  err.number <> 0 then
	Response.Write ("No se han podido actualizar los datos soliciatados,  Reintente en unos minutos")
	Cnn.RollbackTrans
	Cnn.Abort    
else
	Cnn.CommitTrans	
end if

Cnn.Close	
set Cnn = Nothing
SET RS = Nothing	

%>

<script language="javascript" type="text/jscript">
    alert("Datos Actualizados")
</script>
</BODY>
</HTML>
