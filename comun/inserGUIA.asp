<%@ Language=VBScript%>
<%Response.Buffer = true %>
<%Session.LCID=2058%>
<%Usuario = Request.Cookies("tienda")("usr")%>
<%tda     = Request.Cookies("tienda")("pos")%>
<%destda  = Request.Cookies("tienda")("tda") %>

<!--#include file="../includes/Cnn.inc"-->
<!--#include file="../comun/funcionescomunes.asp"-->
<!--#include file="../includes/funcionesVBscript.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>

<body>
<%
cpos = request.form("POS")	' numero de documento origen
cradio = request.form("radio1")	'FAC 20130320	opcion seleccionada
cdocori= "GS"	'FAC 20130325
cdocIS = "NI"	'FAC 20130325

if cradio="Rop3" then	'FAC 20130325
	cdocIS="NS"
end if
if cradio="Rop2" then 
	cdocori="PE"
end if
if cradio="Rop3" then 
	cdocori="PS"
end if
FECHA ="{ fn NOW() }"
cwhere = request.Form("fcwhere")	' condicion pra el select

'response.Write("<br>"+cpos)
'response.Write("<br>"+cradio)
'response.Write("<br>"&cdocIS)
'response.Write("<br>"+cwhere)

' Filtrando articulos a insertar
existen=" SELECT c6_ccodigo                             " & _
		" FROM   RSFACCAR..al0012movd as RS             " & _
 		cwhere & _       
        " except                                        " & _
        " SELECT codigo COLLATE Modern_Spanish_CI_AI    " & _
        " FROM articulos where tienda = '"&tda&"';      " 
'" where c6_ctd = 'GS' and c6_cnumdoc ='"&cpos&"'" & _ 
'        " AND C6_CCODIGO <> 'TXT'                       " & _ 

'response.Write("<br>"+existen)

rs.open existen,cnn

if rs.recordcount >0 then
	Response.Write rs.recordcount	
	RS.MOVEFIRST
	response.Write(" Nuevos articulos")

	CAD = ""
	DO WHILE NOT RS.EOF
		COD = RS("c6_ccodigo")
		CAD = CAD + " insert into articulos  (TIENDA,   " & _
		    		" CODIGO, STOCK, MINIMO, LISTA1,    " & _
                    " LISTA2, LISTA3, planilla, usuario," & _
	    	    	"  fecha,ESTADO )    values(       " & _
	    	    	" '"&tda&"','"&COD&"', 0,      " & _
                    " 2,  0, 0,   0,  '1',          " & _
                    " 'AUTOM', "&fecha&", 'A');         " 
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
	'MAILS = "EXEC SP_correo_precios '"&TDA&"', '"&cpos&"', '"&destda&"' "
	'CNN.EXECUTE MAILS
end if
rs.close

'*********************** OJO **********************
'REVISAR EL LARGO DE LAS COLUMNAS de la tabla
'*********************** OJO **********************
lin = Request.QueryString("lin")
doc = request.form("pos")
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
        " WHERE CIA = '"&TDA&"' AND CODIGO = '"&cdocIS&"'"

'response.Write("<br>"&kad)

RS.OPEN kAD,CNN

CC = CDBL(RS("CORREL")) + 1
COR = RIGHT("00000000000"+TRIM(CC),7)
SER = RS("SERIE")
TIP = RS("TIPMOV")
RS.CLOSE

' PRIMERO ACTUALIZO EL CORRELATIVO
CAD =   " UPDATE DOCUMENTO SET CORREL = '"&COR&"'           " & _
        " WHERE CIA = '"&TDA&"' AND CODIGO = '"&cdocIS&"';  "

'SEGUIMOS CON LA CABECERA DEL DOCUMENTO
cad = CAD +   " insert into movimcab (                      " & _
              " OPERACION, TIENDA, CODDOC, TIPMOV, SERIE,   " & _
              " NUMDOC, FECDOC, DOCORI, NUMORI, MONEDA,     " & _
              " USUARIO, FECHA, ESTADO) VALUES (  " & _   
              " '"&OPERACION&"', '"&TDA&"', '"&cdocIS&"', '"&TIP&"'," & _ 
              " '"&SER&"',  '"&COR&"', "&FECHA&", '"&cdocori&"',     " & _
              " '"&DOC&"', NULL,  '"&USUARIO&"',       " & _
              " "&FECHA&", 'A' ) ;                          " 
             
' AHORA HAY QUE LLENAR LAS LINEAS DE DETALLE DE MOVIMIENTO.....
' Y ACTUALIZAR STOCKS
for i = 1 to lin

	CODI= (Request.Form("COD"&I))
	canI= (Request.Form("CAN"&I))
	canS= (Request.Form("CAN"&I))
	PreI= (Request.Form("PRE"&I))
    if len(trim(prei)) = 0 then prei = 0
    IT  = RIGHT("00"+ CSTR(I),2) 

	if tip="E" then	'FAC 20130326
		cans=0
	else
		cani=0
	end if	

	cad = cad + " INSERT INTO MOVIMDET	                            " & _
		        " ( OPERACION, ITEM, CODART, ENTRA, sale, TIENDA )  " & _
		        " VALUES('"&OPERACION&"', '"&IT&"',		            " & _
		        " '"&CODI&"', '"&CANI&"', '"&canS&"', '"&TDA&"' ) ;	"
	if tip="E" then
    CAD = CAD + " UPDATE ARTICULOS SET PLANILLA = '1',          " & _
                " STOCK = STOCK + "&CANI&", Lista1 =  (SELECT AR_NPRECI6 FROM RSFACCAR..AL0012ARTI WHERE AR_CCODIGO = '"&CODI&"')  " & _
				" WHERE TIENDA =                                " & _
                " '"&TDA&"' AND CODIGO = '"&CODI&"' ;           "
	else
	CAD = CAD + " UPDATE ARTICULOS SET PLANILLA = '1',          " & _
                " STOCK = STOCK - "&CANS&", Lista1 = (SELECT AR_NPRECI6 FROM RSFACCAR..AL0012ARTI WHERE AR_CCODIGO = '"&CODI&"')  " & _
				" WHERE TIENDA =                                " & _
                " '"&TDA&"' AND CODIGO = '"&CODI&"' ;           "
	end if
NEXT


' Inicia transacción , para que los datos no queden a medias
' RESPONSE.WRITE("<br>"&CAD)
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
	cad = '../INGRESOS.ASP?POS='+'<%=doc%>'+'&crad='+'<%=cradio%>'
	window.location.replace = cad
</script>

</BODY>
</HTML>
