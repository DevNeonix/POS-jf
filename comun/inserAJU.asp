<%@LANGUAGE="VBSCRIPT"  CODEPAGE="65001"%>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<%Usuario = Request.Cookies("tienda")("usr")%>
<%tda     = Request.Cookies("tienda")("pos")%>
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
RESPONSE.Write(COD)
aCod = split(cod,",")
' cantidad vendida
can = Request.QueryString("can")
response.Write(cAn)
aCan = split(can,",")
doc = request.QueryString("doc")
dES = left(request.QueryString("obs"),50)
tip = request.QueryString("tip")
mov = request.QueryString("mov")
ser = request.QueryString("ser")
nro = request.QueryString("nro")
FEC ="getdate()"
mon = "MN"
cli = request.QueryString("cli")

' Inicia transacción , para que los datos no queden a medias
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
            " VENDEDOR, usuario, fecha, ESTADO , porigv)" & _
            " values( '"&OPE&"','"&TDA&"', '"&DOC&"',   " & _
            " '"&mov&"', '"&SER&"', '"&NRO&"', "&FEC&", " & _
            " '"&MON&"', null, null, null, null, null,  " & _
            " '"&CLI&"', '"&USUARIO&"',                 " & _
            " '"&USUARIO&"',    "&fec&" ,'A', null);    " 

CAD = CAD + " INSERT INTO VARIOS (OPERACION, DESCRIPCION, moti)  " & _
            " VALUES('"&OPE&"', '"&DES&"', '"&tip&"') ;          "

FOR I=0 TO 99
    IF (LEN(TRIM(ACOD(I))) > 0  ) THEN
        LIN = RIGHT("00" + TRIM(CSTR(I+1)),2)
        
        ' LINEA DE DETALLE DE DOCUMENTO DE VENTA
        CAD = CAD + " INSERT INTO MOVIMDET (OPERACION, TIENDA, ITEM,    " & _
                    " CODART, entra, PRECIO, DESCUENTO, IGV, PORDES )    " & _
                    " VALUES ('"&OPE&"', '"&TDA&"', '"&LIN&"',          " & _
                    " '"&aCod(I)&"',  "&aCan(i)&", null, null, null,    " & _
                    " null) ;                                           "
        ' ACTUALIZA EL STOCK DEL ARTICULO
        CAD = CAD + " UPDATE ARTICULOS SET STOCK = STOCK + "&aCan(i)&" ,    " & _
                    " usuario = '"&usuario&"', fecha = "&fec&", estado = 'A'" & _
                    " where TIENDA = '"&tda&"' and codigo = '"&aCod(i)&"' ; " 
    ELSE
        EXIT FOR
    END IF
NEXT    
' ACTUALIZAR EL CORRELATIVO DEL DOCUMENTO EMITIDO

CAD = CAD + " UPDATE DOCUMENTO SET CORREL = '"&NRO&"'       " & _
            " WHERE CIA = '"&TDA&"' AND CODIGO = '"&DOC&"'; "

' INGRESAR LOS VALORES DEL MOVIMIENTO DE CAJA
LN = 1

RESPONSE.WRITE(CAD)
RESPONSE.WRITE("<br>")
RESPONSE.WRITE("<br>")


'RESPONSE.END




Cnn.Execute(CAD)
if  err.number <> 0 then
	Response.Write ("No se han podido actualizar los datos soliciatados,  Reintente en unos minutos")
	Cnn.RollbackTrans
	Cnn.Abort%>
<%else
	Cnn.CommitTrans	
end if

Cnn.Close	
set Cnn = Nothing
SET RS = Nothing	

%>

<script language="javascript" type="text/jscript">
doc = '<%=trim(doc) %>'
   
this.window.focus()
var opc = "directories=no,height=600,";
opc = opc + "hotkeys=no,location=no,";
opc = opc + "menubar=no,resizable=yes,";
opc = opc + "left=0,top=0,scrollbars=yes,";
opc = opc + "status=no,titlebar=yes,toolbar=yes,";
opc = opc + "width=900";


    window.location.replace('../notaAJU.asp?ope=' + '<%=OPE %>','AJUSTE', opc)
</script>
</BODY>
</HTML>
