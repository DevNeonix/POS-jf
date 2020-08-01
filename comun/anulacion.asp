<%@ Language=VBScript %>
<%TDA = Request.Cookies("TIENDA")("POS")%>
<%Response.Buffer = TRUE %>
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
pos = ucase(TRIM(Request.QueryString("pos")))


cad = "update caja set monto=0, tipo='ANU' where operacion = '"&pos&"'; "
'*************************************************************
' BORRA  LAS LINEAS DE TIPOS DE PAGO Y DEJA SOLO UNA
' PARA LA ANULACION EN CERO
'************************************************************
cad =  cad + " DELETE CAJA  where operacion = '"&pos&"' AND LIN > 1 ; "
'************************************************
' ACTUALIZA CABECERA PONE EN CERO LOS MONTOS
' COLOCA CLIENTE = ANULADO
'************************************************
CLI = RIGHT("00000000000" + TDA,11)
cad = cad + " update movimcab set cliente = '"&CLI&"', PVP=0, DESCUENTO=0,  " & _
            " SUBTOT=0, IGV = 0, TOTAL=0, DOCORI = NULL, SERORI = NULL, NUMORI = NULL  where operacion = '"&pos&"';       "

'************************************************
' aumenta o resta el stock segun sea entrada o
' VENTA, SI ES VALE... NO SE MUEVE EL STOCK
'************************************************
DAR = "SELECT * FROM MOVIMDET WHERE OPERACION = '"&POS&"' "
RS.OPEN DAR, CNN
RS.MOVEFIRST
    DO WHILE NOT RS.EOF
        canti=0
        IF cdbl(RS("ENTRA")) > 0 THEN CANTI = cdbl(RS("ENTRA"))*-1
        IF cdbl(RS("SALE")) > 0 THEN CANTI = cdbl(RS("SALE"))
        IF cdbl(RS("VALE")) > 0 THEN CANTI = 0
        cod = rs("codart")
        ' primero actualizamos stock
        cad = cad  + "update articulos set stock = stock + "&canti&"  where codigo = '"&cod&"' and tienda = '"&tda&"'; "
        ln = rs("item")
        ' ahora actualizamos a 0 el detalle linea a linea
        cad= cad +  " update MOVIMDET set entra=0, sale=0, vale=0, precio=0, descuento = 0, igv=0,   " & _
                    " pordes =null, flag = '' WHERE OPERACION = '"&POS&"' and codart = '"&cod&"' and item = '"&ln&"'; "

        RS.MOVENEXT
    LOOP


'response.end
Cnn.BeginTrans	
Cnn.Execute(CAD)
if  err.number <> 0 then
	Response.Write ("No se han podido actualizar los datos soliciatados,  Reintente en unos minutos")
	Cnn.RollbackTrans
	Cnn.Abort
else
	Cnn.CommitTrans	
end if

rs.close
Cnn.Close	
set Cnn = Nothing
SET RS = Nothing	

%>

<script language="javascript" type="text/jscript">
    alert("STOCKS REINTEGRADOS. LISTO");
    window.location.replace('../anuladeta.asp?pos=')  
</script>

</BODY>
</HTML>
