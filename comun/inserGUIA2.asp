<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<%Usuario = Request.Cookies("tienda")("usr")%>
<%tda     = Request.Cookies("tienda")("pos")%>
<%destda  = Request.Cookies("tienda")("tda") %>
<%Response.Buffer = TRUE %>
<!--#include file="../includes/Cnn.inc"-->
<!--#include file="../comun/funcionescomunes.asp"-->
<!--#include file="../includes/funcionesVBscript.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<body>

<%
pos = request.form("POS")
' Filtrando articulos a insertar
existen=" SELECT c6_ccodigo                             " & _
		" FROM   RSFACCAR.dbo.al0012movd as RS          " & _
        " where c6_ctd = 'GS' and c6_cnumdoc ='"&pos&"' " & _ 
        " AND C6_CCODIGO <> 'TXT'                       " & _ 
        " except                                        " & _
        " SELECT codigo COLLATE Modern_Spanish_CI_AI    " & _
        " FROM articulos where tienda = '"&tda&"';      " 

rs.open existen,cnn

if rs.recordcount >0 then
	Response.Write rs.recordcount
	RS.MOVEFIRST
	CAD = ""
	DO WHILE NOT RS.EOF
		COD = RS("c6_ccodigo")
		CAD = CAD + " insert into articulos  (TIENDA,   " & _
		    		" CODIGO, MINIMO, STOCK, LISTA1,    " & _
                    " LISTA2, LISTA3, ESTADO, planilla, " & _
	    	    	" usuario, fecha )    values(       " & _
	    	    	" '"&tda&"','"&COD&"', '2',      " & _
                    " 0,  0, 0,   0, 'A', '1',          " & _
                    " 'AUTOM',  "&date()&" );         " 
   		rs.movenext
	Loop
	CNN.BEGINTRANS
	CNN.EXECUTE CAD
	if err.number <> 0 then
		Response.Write ("No se han podido actualizar los datos soliciatados,  Reintente en unos minutos")
		Cnn.RollbackTrans
		Cnn.Abort
	else
		Cnn.CommitTrans
	end if
	MAILS = "EXEC SP_correo_precios '"&TDA&"', '"&POS&"', '"&destda&"' "
	CNN.EXECUTE MAILS
end if
rs.close


'*********************** OJO **********************
'REVISAR EL LARGO DE LAS COLUMNAS de la tabla
'*********************** OJO **********************
lin = Request.QueryString("lin")
doc = request.form("POS")
FECHA ="{ fn NOW() }"

'PRIMERO SE HACE EL INSERT EN LA CABECERA
RS.OPEN "SELECT top 1 cast(operacion as float) AS TOPE FROM movimcab order by 1 desc",CNN
' fac 20121228 no se pq no lo hace adicione if
if rs.eof then
	OPERACION = RIGHT("0000000000"+ TRIM(CSTR(1)),10)
else
	OPERACION = RIGHT("0000000000"+ TRIM(CSTR(RS("TOPE")+1)),10)
end if
RS.CLOSE

KAD =   " SELECT SERIE, CORREL, TIPMOV FROM DOCUMENTO " & _
        " WHERE CIA = '"&TDA&"' AND CODIGO = 'NI'     "
RS.OPEN kAD,CNN
CC = CDBL(RS("CORREL")) + 1
COR = RIGHT("00000000000"+TRIM(CC),7)
SER = RS("SERIE")
TIP = RS("TIPMOV")
RS.CLOSE

' PRIMERO ACTUALIZO EL CORRELATIVO
CAD =   " UPDATE DOCUMENTO SET CORREL = '"&COR&"'           " & _
        " WHERE CIA = '"&TDA&"' AND CODIGO = 'NI';          "

'SEGUIMOS CON LA CABECERA DEL DOCUMENTO
cad = CAD +   " insert into movimcab (                      " & _
              " OPERACION, TIENDA, CODDOC, TIPMOV, SERIE,   " & _
              " NUMDOC, FECDOC, DOCORI, NUMORI, MONEDA,     " & _
              " USUARIO, FECHA, ESTADO) VALUES (  " & _   
              " '"&OPERACION&"', '"&TDA&"', 'NI', '"&TIP&"'," & _ 
              " '"&SER&"',  '"&COR&"', "&FECHA&", 'GS',     " & _
              " '"&DOC&"', NULL,  '"&USUARIO&"',       " & _
              " "&FECHA&", 'A' ) ;                          " 
             
' AHORA HAY QUE LLENAR LAS LINEAS DE DETALLE DE MOVIMIENTO.....
' Y ACTUALIZAR STOCKS
for i = 1 to lin

	CODI= (Request.Form("COD"&I))
	canI= (Request.Form("CAN"&I))
	PreI= (Request.Form("PRE"&I))
    IT  = RIGHT("00"+ CSTR(I),2) 
	cad = cad + " INSERT INTO MOVIMDET	                    " & _
		        " ( OPERACION, ITEM, CODART, ENTRA,TIENDA ) " & _
		        " VALUES('"&OPERACION&"', '"&IT&"',		    " & _
		        " '"&CODI&"', '"&CANI&"', '"&TDA&"' ) ;		"
                	
    CAD = CAD + " UPDATE ARTICULOS SET PLANILLA = '1',      " & _
                " STOCK = STOCK + "&CANI&", Lista1 = "&PREI    & _
				" WHERE TIENDA =   " & _
                " '"&TDA&"' AND CODIGO = '"&CODI&"' ;       "
NEXT


' Inicia transacción , para que los datos no queden a medias
'RESPONSE.WRITE("<br>")
'RESPONSE.WRITE(CAD)
'RESPONSE.END
Cnn.BeginTrans	
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
'regresa a la página de donde fué llamado, para que vea que agregó el registro
%>

<script language="javascript" type="text/jscript">
	cad = '../INGRESOS.ASP?POS='+'<%=doc%>'
	window.location.replace = cad
</script>

</BODY>
</HTML>
