<%@ Language=VBScript %>
<!--#include file="includes/Cnn.inc"-->
<!--#include file="includes/funcionesVBscript.asp"-->
<!--#include file="funcionesefact.asp"-->
<%  

Response.CharSet = "UTF-8"
CTD = request.querystring("CTD")
SER = request.querystring("SER")
DOC = request.querystring("DOC")

IF UCASE(TRIM(CTD)) <> "NC" THEN 
    response.write("esto no es una Nota de Crédito")
	response.end
END IF 

'CABECERA
cad =   " select * from view_guia_real_notas   " & _
        " where F5_CTD = '"&CTD&"' AND         " & _
        " ltrim(rtrim(F5_CsUNEST)) = '' AND    " & _   
        " F5_CNUMSER = '"&SER&"' AND           " & _
        " F5_CNUMDOC = '"&DOC&"'               "
	'response.write(cad)
    if rs.state > 0 then 
	    rs.close
    end if
    rs.open cad,cnn
    if rs.recordcount <= 0 then 
        RESPONSE.WRITE("Documento invalido")
        response.end
    END IF

	rs.movefirst
    CODCLI = RS("F5_CCODCLI")

    cadDet = " SELECT * FROM (select * from RSFACCAR..FT0012FACD  where F6_CTD = '"&CTD&"' AND F6_CNUMSER = '"&SER&"' AND F6_CNUMDOC = '"&DOC&"' union  " & _
             " select * from RSFACCAR..FT0012ACUD  where F6_CTD = '"&CTD&"' AND F6_CNUMSER = '"&SER&"' AND F6_CNUMDOC = '"&DOC&"') XX WHERE F6_CCODIGO NOT IN( 'TXT' ,'.') ORDER BY F6_CITEM "
	
    set rsDet = RsNuevo
	rsDet.open cadDet,cnn


    cadDet = " SELECT * FROM (select * from RSFACCAR..FT0012FACD  where F6_CTD = '"&CTD&"' AND F6_CNUMSER = '"&SER&"' AND F6_CNUMDOC = '"&DOC&"' union  " & _
             " select * from RSFACCAR..FT0012ACUD  where F6_CTD = '"&CTD&"' AND F6_CNUMSER = '"&SER&"' AND F6_CNUMDOC = '"&DOC&"') XX WHERE F6_CCODIGO NOT IN( 'TXT' ,'.') ORDER BY F6_CITEM "
	'RESPONSE.WRITE(cadDet)
    set rsDet = RsNuevo
	rsDet.open cadDet,cnn

cadCli = "SELECT * FROM RSFACCAR..FT0012CLIE WHERE CL_CCODCLI = '"&CODCLI&"' "
	'response.write(cadCli)
	set rsCLI = RsNuevo
	rsCLI.open cadCli,cnn

	if rsCLI.recordcount <= 0 then
		response.write("Verifique y corrija el cliente.")
		response.end
	end if
    rsCLI.movefirst



'globales'
'=============================================================================================================='
'=============================================================================================================='
'=============================================================================================================='
coddoc = "07"

'DATOS DE LA EMPRESA
'=============================================================================================================='
'=============================================================================================================='
'=============================================================================================================='
miRuc = "20600689101"
miRS = "FAJAS Y ELASTICOS TEXTILES S.A.C"
miNombreComercial = "Jacinta Fernandez"
miDireccion = "AV. LA MAR NRO. 523"




'NOMBRE DE ARCHIVO'
'=============================================================================================================='
'=============================================================================================================='
'=============================================================================================================='

docori = left(ucase(trim(rs("f5_crftd"))),1)
ndoc = ucase(docori&right(trim(rs("f5_crfnser")),3))&"-"&ucase(right(trim(rs("f5_crfndoc")),7))
nombre = miRuc&"-"&coddoc&"-"&trim(rs("f5_cnumser"))&"-"&trim(rs("f5_cnumdoc"))
'FIN NOMBRE DE ARCHIVO'
'=============================================================================================================='
 
	
    ' el Real no acumula el IGV para Boletas!!!!!!!!!!!!!!!!!!!
    cadIGV = "select  sum(igv) as igv, sum(DCTO) AS DCTO from (Select sum(F6_NIGV) as igv, sum(F6_NDESCTO) AS DCTO From RSFACCAR..FT0012ACUD Where F6_CCODAGE= '0001' AND F6_CTD ='"&ctd&"' AND F6_CNUMSER='"&ser&" ' AND F6_CNUMDOC='"&doc&"' " & _
             " UNION Select sum(F6_NIGV) as igv, sum(F6_NDESCTO) AS DCTO From RSFACCAR..FT0012FACD Where F6_CCODAGE= '0001' AND F6_CTD ='"&ctd&"' AND F6_CNUMSER='"&ser&" ' AND F6_CNUMDOC='"&doc&"' ) as igvdes"
	'response.write(cadIGV)
    set rsIGV = RsNuevo
	rsIGV.open cadIGV,cnn


'   pvp 		= FORMATNUMBER(cdbl(RS("pvp")),2,,,false)
'   SUBTOTAL 	= FORMATNUMBER(cdbl(RS("SUBTOT")),2,,,false)
'   IGV 		= FORMATNUMBER(cdbl(RS("IGV")),2,,,false)
'   ICBPER 		= formatnumber(cdbl(reemplazaNull(RS("isc"),0)),2,,,false)
'   Total 		= FORMATNUMBER(cdbl(RS("TOTAL")),2,,,false)
'   MON 		= "PEN"
'   TIPOPERACION = "0101"
'   cntLinDet = rsDet.recordcount
'   codUBIGUEO="150103"

IGV 		= FORMATNUMBER(cdbl(rsIGV("IGV")),2,,,false)
    ' SIN IGV
    SUBTOTAL 	= FORMATNUMBER(cdbl(RS("F5_NIMPORT"))-CDBL(IGV),2,,,false)
	pvp 	    = FORMATNUMBER(cdbl(RS("F5_NIMPORT"))-CDBL(IGV),2,,,false)
	ICBPER 		= 0
	Total 		= FORMATNUMBER(cdbl(RS("f5_nimport")),2,,,false)
    IF UCASE(TRIM(RS("F5_CCODMON"))) = "MN" THEN 	MON = "PEN" ELSE MON = "USD"
	TIPOPERACION = "0101"
	cntLinDet = rsDet.recordcount
	codUBIGUEO="150103"




'DATOS DEL CLIENTE


'0 - DOC.TRIB.NO.DOM.SIN.RUC
'1 - DNI
'4 - Carnet de extranjería
'6 - Registro Único de Contribuyentes
'7 - PasaporteA - Cédula Diplomática de identidad
'B - DOC.IDENT.PAÍS.RESIDENCIA-NO.D
'C - TaxIdentificationNumber-TIN–DocTrib PP.NN
'D - Identification Number - IN – Doc Trib PP. JJ
'E - TAM- Tarjeta Andina de Migración'
clienteDOC 			= trim(rsCLI("cl_ctipper"))

''	clienteDOC = "72218501"

	if clienteDOC = "N" then
		clienteTipDoc		= "1"
        NumDocCLI = trim(rsCLI("cl_cnumide"))
	else

		clienteTipDoc		= "6"
		NumDocCLI = trim(rsCLI("cl_cnumruc"))
		


	end if
	clienteDireccion	= trim(rsCLI("cl_cdircli"))
	clienteRazon		= AlphaNumericOnly(trim(rsCLI("cl_cnomcli")))
	clienteMail			= trim(rsCLI("cl_CEMAIL"))


'***********CLIENTE VARIOS*******'
'RESPONSE.WRITE(NumDocCLI)
if NumDocCLI =  "000000000" then
	clienteTipDoc		= "0"
	NumDocCLI		= "000000000"
end if
if LEN(TRIM(NumDocCLI)) =  8 then
	clienteTipDoc		= "1"
	NumDocCLI		= NumDocCLI
end if

'FILA1
'=============================================================================================================='
'=============================================================================================================='
'=============================================================================================================='


'FILA 1 A: * Fecha Emisión'
f1Fecdoc = trim(AlphaNumericOnly(CStr(year(RS("F5_DFECDOC"))))   &"-"&   AlphaNumericOnly(right("00"&month(RS("F5_DFECDOC")),2))  &"-"&    AlphaNumericOnly(right("00"&day(RS("F5_DFECDOC")),2)))
'f1Fecdoc = "2019-09-28"
'FILA 1 B: * Número de documento'
F1Numdoc 	= trim(rs("f5_cnumser"))&"-"&trim(rs("f5_cnumdoc"))

'FILA 1 C :   Tipo de moneda 
F1C 		=MON

'FILA 1 D :   Sumatoria Monto baseIGV o IVAP
F1D 		=abs(pvp		)

'FILA 1 E :   Total IGV o IVAP 
F1E 		=abs(igv		)

'FILA 1 F :   Tipo monedaIGV o IVAP
F1F 		=MON

'FILA 1 G :   Sumatorio monto base ISC 
F1G 		=""		

'FILA 1 H :   Sumatoria monto total ISC 
F1H 		=""		

'FILA 1 I :   Tipo moneda ISC 
F1I 		=""		

'FILA 1 J :   Sumatoria monto baseOtros tributos
F1J 		=""		

'FILA 1 K :   Sumatoria monto totalOtro tributos
F1K 		=""		

'FILA 1 L :   Tipo moneda Otro tributos 
F1L 		=""		

'FILA 1 M :   Importe total delcomprobante
F1M 		=abs(total)		

'FILA 1 N :   Monto otros cargos 
F1N 		=""		

'FILA 1 O :   Total operacionesexportación
F1O 		=""		

'FILA 1 P :   Total operaciones gravadasIGV o IVAP
F1P 		=abs(SUBTOTAL)		

'FILA 1 Q :   Total operaciones Inafectas 
F1Q 		=""		

'FILA 1 R :   Total operacionesexoneradas
F1R 		=""		

'FILA 1 S :   Total operacionesGratuitas
F1S 		=""		

'FILA 1 T :   Sumatoria de impuestosde operaciones gratuitas
F1T 		=""		

'FILA 1 U :	  Sumatoria monto totalICBPER
F1U 		=abs(ICBPER)		

'FILA 1 V :
F1V 		=""		

'FILA 1 W :
F1W 		=""		

'FILA 1 X :
F1X 		=""		

'FILA 1 Y :
F1Y 		=""		

'FILA 1 Z :
F1Z 		=""		

'FILA 1 AA:
F1AA		=""		

'FILA 1 AB:
F1AB		=""		

'FILA 1 AC:
F1AC		=""		

'FILA 1 AD:
F1AD		=""		

'FILA 1 AE:   Cantidad de Líneas deldocumento
F1AE		=cntLinDet	

'FILA 1 AF:   Cantidad documentosasociados
F1AF		="1"		

'FILA 1 AG:   Cantidad guías asociadas yotros documentosasociados
F1AG		=""		

'FILA 1 AH:   Monto para Redondeo 
F1AH		=""		

'FILA 1 AI:   Monto total deimpuestos
F1AI		=abs(formatnumber(cdbl(IGV)+cdbl(ICBPER),2,,,false)		)

'FILA 1 AJ:   Total valor de venta
F1AJ		=abs(SUBTOTAL)		



fila1 = f1Fecdoc 	
fila1 = fila1&","&F1Numdoc 	
fila1 = fila1&","&F1C 		
fila1 = fila1&","&F1D 		
fila1 = fila1&","&F1E 		
fila1 = fila1&","&F1F 		
fila1 = fila1&","&F1G 		
fila1 = fila1&","&F1H 		
fila1 = fila1&","&F1I 		
fila1 = fila1&","&F1J 		
fila1 = fila1&","&F1K 		
fila1 = fila1&","&F1L 		
fila1 = fila1&","&F1M 		
fila1 = fila1&","&F1N 		
fila1 = fila1&","&F1O 		
fila1 = fila1&","&F1P 		
fila1 = fila1&","&F1Q 		
fila1 = fila1&","&F1R 		
fila1 = fila1&","&F1S 		
fila1 = fila1&","&F1T 		
fila1 = fila1&","&F1U 		
fila1 = fila1&","&F1V 		
fila1 = fila1&","&F1W 		
fila1 = fila1&","&F1X 		
fila1 = fila1&","&F1Y 		
fila1 = fila1&","&F1Z 		
fila1 = fila1&","&F1AA		
fila1 = fila1&","&F1AB		
fila1 = fila1&","&F1AC		
fila1 = fila1&","&F1AD		
fila1 = fila1&","&F1AE		
fila1 = fila1&","&F1AF		
fila1 = fila1&","&F1AG		
fila1 = fila1&","&F1AH		
fila1 = fila1&","&F1AI		
fila1 = fila1&","&F1AJ		

response.write(utf8_simbom(fila1))

'FILA2
'=============================================================================================================='
'=============================================================================================================='
'=============================================================================================================='


'FILA 2 A:  Numero de guia 
F2A		=	""
'FILA 2 B:  Código de la guia 
F2B		=	""
'FILA 2 C:  Número otro documento 
F2C		=	""
'FILA 2 D:  Código del tipo otrodocumento
F2D		=	""
'FILA 2 E:  ATTACH_DOC
F2E		=	""

fila2=F2A
fila2=fila2&","&F2B
fila2=fila2&","&F2C
fila2=fila2&","&F2D
fila2=fila2&","&F2E


response.write("<br/>")
response.write(utf8_simbom(fila2))


'FILA3
'=============================================================================================================='
'=============================================================================================================='
'=============================================================================================================='

'FILA 5 A:Apellidos y nombres,denominación o razónsocial
f3RazonSocial 							=miRS
'FILA 5 B:Nombre comercial 
f3NomComercial 							=miNombreComercial
'FILA 5 C:Número de RUC 
f3NumRuc	 							=miRuc
'FILA 5 D:Código Ubigeo 
f3CodUbigeo 							=codUBIGUEO
'FILA 5 E:Dirección 
f3Direccion 							=miDireccion
'FILA 5 F:Urbanización 
f3Urbanizacion 							=""
'FILA 5 G:Departamento 
f3Departamento 							="LIMA"
'FILA 5 H:Provincia 
f3Provincia 							="LIMA"
'FILA 5 I:Distrito 
f3Distrito 								="ATE"
'FILA 5 J:Codigo del pais 
f3CodPais	 							="PE"
'FILA 5 K:Código delestablecimiento
f3CodEst	 							="0000"




fila3 = f3RazonSocial
fila3 = fila3 &","&f3NomComercial
fila3 = fila3 &","&f3NumRuc	
fila3 = fila3 &","&f3CodUbigeo
fila3 = fila3 &","&f3Direccion
fila3 = fila3 &","&f3Urbanizacion
fila3 = fila3 &","&f3Departamento
fila3 = fila3 &","&f3Provincia
fila3 = fila3 &","&f3Distrito
fila3 = fila3 &","&f3CodPais	
fila3 = fila3 &","&f3CodEst
response.write("<br/>")
response.write(utf8_simbom(fila3))


'FILA4
'=============================================================================================================='
'=============================================================================================================='
'=============================================================================================================='


'FILA 4  A:Número de documento
f4NroDoc 								=NumDocCLI
'FILA 4  B:Tipo de documento 
f4TipDoc 								=clienteTipDoc
'FILA 4  C:Razón social 
f4RazonSocial 							=clienteRazon
'FILA 4  D:Nombre comercial 
f4NomComercial 							=""
'FILA 4  E:Código ubigeo 
f4CodUbigeo 							=""
'FILA 4  F:Dirección 
f4Direccion 							=clienteDireccion
'FILA 4  G:Urbanización 
f4Urbanizacion 							=""
'FILA 4  H:Departamento 
f4Departamento 							=""
'FILA 4  I:Provincia 
f4Provincia 							=""
'FILA 4  J:Distrito 
f4Distrito 								=""
'FILA 4  K:Código de país 
f4CodPais 								="PE"
'FILA 4  L:Correo
f4Correo 								=clienteMail

fila4 = f4NroDoc
fila4 = fila4 &","&f4TipDoc 					
fila4 = fila4 &","&f4RazonSocial 
fila4 = fila4 &","&f4NomComercial
fila4 = fila4 &","&f4CodUbigeo 		
fila4 = fila4 &","&f4Direccion 		
fila4 = fila4 &","&f4Urbanizacion
fila4 = fila4 &","&f4Departamento
fila4 = fila4 &","&f4Provincia 		
fila4 = fila4 &","&f4Distrito 			
fila4 = fila4 &","&f4CodPais 				
fila4 = fila4 &","&f4Correo					

response.write("<br/>")
response.write(utf8_simbom(fila4))


'FILA7
'=============================================================================================================='
'=============================================================================================================='
'=============================================================================================================='

f51000 = Numlet(ABS(cdbl(total)))
f51002 = ""
f52000 = ""
f52001 = ""
f52002 = ""
f52003 = ""
f52004 = ""
f52005 = ""
f52006 = ""
f52007 = ""
f52008 = ""
f52009 = ""
f52010 = ""

fila5 = f51000
fila5 = fila5&","&f51002
fila5 = fila5&","&f52000
fila5 = fila5&","&f52001
fila5 = fila5&","&f52002
fila5 = fila5&","&f52003
fila5 = fila5&","&f52004
fila5 = fila5&","&f52005
fila5 = fila5&","&f52006
fila5 = fila5&","&f52007
fila5 = fila5&","&f52008
fila5 = fila5&","&f52009
fila5 = fila5&","&f52010
response.write("<br/>")
response.write(utf8_simbom(fila5))



'FILA6
'=============================================================================================================='
'=============================================================================================================='
'=============================================================================================================='

'FILA 6 A:  Observaciones 
FILA6A	=	"hemos debitado su cuentas"

'FILA 6 B:  Orden de compra 
FILA6B	=	""

'FILA 6 C:  Fecha de vencimiento 
FILA6C	=	""

'FILA 6 D:  Motivo de nota 
FILA6D	=	""

'FILA 6 E:  Pedido de nota 
FILA6E	=	""

'FILA 6 F:  Código cliente 
FILA6F	=	""

'FILA 6 G:  Código vendedor 
FILA6G	=	""

'FILA 6 H:  Código venta 
FILA6H	=	""

'FILA 6 I:  Orden de venta 
FILA6I	=	""

'FILA 6 J:  Número Interno 
FILA6J	=	""

'FILA 6 K:  Número pedido 
FILA6K	=	""

'FILA 6 L:  Condición de pago 
FILA6L	=	""

'FILA 6 M:  Condición general 
FILA6M	=	""

'FILA 6 N:  Tipo de pago 
FILA6N	=	""

'FILA 6 O:  Forma de pago 
FILA6O	=	""

'FILA 6 P:  Fecha de pago 
FILA6P	=	""

'FILA 6 Q:  Fecha de orden 
FILA6Q	=	""

'FILA 6 R:  Teléfono / Fax 
FILA6R	=	""

'FILA 6 S:  Emitido por
FILA6S	=	""

'FILA 6 T:  Entrega Factura 
FILA6T	=	""

'FILA 6 U:  Tipo de cambio 
FILA6U	=	""

'FILA 6 V:  Código SAP 
FILA6V	=	""

'FILA 6 W:  Sede 
FILA6W	=	""

'FILA 6 X:  Usuario 
FILA6X	=	""

'FILA 6 Y:  Solicitud 
FILA6Y	=	""

'FILA 6 Z:  Oficina venta 
FILA6Z	=	""

'FILA 6 AA:  Firma 
FILA6AA	=	""

'FILA 6 AB:  Contrato 
FILA6AB	=	""

'FILA 6 AC:  Proyecto 
FILA6AC	=	""

'FILA 6 AD:  Fecha de salida 
FILA6AD	=	""

'FILA 6 AE:  Dirección de entrega 
FILA6AE	=	""

'FILA 6 AF:  Lote 
FILA6AF	=	""

'FILA 6 AG:  Producto 
FILA6AG	=	""

'FILA 6 AH:  Flete 
FILA6AH	=	""

'FILA 6 AI:  Seguro 
FILA6AI	=	""

'FILA 6 AJ:  Total CFR/CPT 
FILA6AJ	=	""

'FILA 6 AK:  Total FOB/FCA 
FILA6AK	=	""

'FILA 6 AL:  Intereses 
FILA6AL	=	""

'FILA 6 AM:  Comisiones 
FILA6AM	=	""

'FILA 6 AN:  Nro DUA 
FILA6AN	=	""

'FILA 6 AO:  Numero contenedor 
FILA6AO	=	""

'FILA 6 AP:  Total bultos 
FILA6AP	=	""

'FILA 6 AQ:  Total Artículos 
FILA6AQ	=	""

'FILA 6 AR:  Total Bruto 
FILA6AR	=	""

'FILA 6 AS:  Almacén 
FILA6AS	=	""

'FILA 6 AT:  O 
FILA6AT	=	""

'FILA 6 AU:  C 
FILA6AU	=	""

'FILA 6 AV:  Z - OF 
FILA6AV	=	""

'FILA 6 AW:  G 
FILA6AW	=	""

'FILA 6 AX:  Número de documento otroparticipante
FILA6AX	=	""

'FILA 6 AY:  Tipo de documento otrosparticipante
FILA6AY	=	""

'FILA 6 AZ:  Apellidos y nombres de otroparticipante
FILA6AZ	=	""

'FILA 6 BA:  ID de orden de compra 
FILA6BA	=	""

'FILA 6 BB:  Referencia de cliente
FILA6BB	=	""



FILA6 = FILA6A
FILA6 = FILA6&","&FILA6B
FILA6 = FILA6&","&FILA6C
FILA6 = FILA6&","&FILA6D
FILA6 = FILA6&","&FILA6E
FILA6 = FILA6&","&FILA6F
FILA6 = FILA6&","&FILA6G
FILA6 = FILA6&","&FILA6H
FILA6 = FILA6&","&FILA6I
FILA6 = FILA6&","&FILA6J
FILA6 = FILA6&","&FILA6K
FILA6 = FILA6&","&FILA6L
FILA6 = FILA6&","&FILA6M
FILA6 = FILA6&","&FILA6N
FILA6 = FILA6&","&FILA6O
FILA6 = FILA6&","&FILA6P
FILA6 = FILA6&","&FILA6Q
FILA6 = FILA6&","&FILA6R
FILA6 = FILA6&","&FILA6S
FILA6 = FILA6&","&FILA6T
FILA6 = FILA6&","&FILA6U
FILA6 = FILA6&","&FILA6V
FILA6 = FILA6&","&FILA6W
FILA6 = FILA6&","&FILA6X
FILA6 = FILA6&","&FILA6Y
FILA6 = FILA6&","&FILA6Z
FILA6 = FILA6&","&FILA6AA	
FILA6 = FILA6&","&FILA6AB	
FILA6 = FILA6&","&FILA6AC	
FILA6 = FILA6&","&FILA6AD	
FILA6 = FILA6&","&FILA6AE	
FILA6 = FILA6&","&FILA6AF	
FILA6 = FILA6&","&FILA6AG	
FILA6 = FILA6&","&FILA6AH	
FILA6 = FILA6&","&FILA6AI	
FILA6 = FILA6&","&FILA6AJ	
FILA6 = FILA6&","&FILA6AK	
FILA6 = FILA6&","&FILA6AL	
FILA6 = FILA6&","&FILA6AM	
FILA6 = FILA6&","&FILA6AN	
FILA6 = FILA6&","&FILA6AO	
FILA6 = FILA6&","&FILA6AP	
FILA6 = FILA6&","&FILA6AQ	
FILA6 = FILA6&","&FILA6AR	
FILA6 = FILA6&","&FILA6AS	
FILA6 = FILA6&","&FILA6AT	
FILA6 = FILA6&","&FILA6AU	
FILA6 = FILA6&","&FILA6AV	
FILA6 = FILA6&","&FILA6AW	
FILA6 = FILA6&","&FILA6AX	
FILA6 = FILA6&","&FILA6AY	
FILA6 = FILA6&","&FILA6AZ	
FILA6 = FILA6&","&FILA6BA	
FILA6 = FILA6&","&FILA6BB	


response.write("<br/>")
response.write(utf8_simbom(FILA6))

'********************************************************'
'********************************************************'
'********FILA 7 DOCUMENTO AL QUE SE MODIFICA*************'
'********************************************************'
'********************************************************'
oriSer		= left(trim(rs("F5_CRFNSER")),4)
	oriNumDoc	= right(trim(rs("F5_CRFNDOC")),7)
if trim(rs("F5_CRFTD")) = "BV" then	
	oriTipDoc	= "03"
elseif trim(rs("F5_CRFTD")) = "FT" then	
	oriTipDoc	= "01"
else
	response.write("No se reconoce este documento")
	response.end
end if

CADORI = " SELECT * FROM (SELECT * FROM RSFACCAR..FT0012ACUC                       " & _
             " WHERE F5_CNUMSER = '"&ORISER&"' AND F5_CNUMDOC = '"&oriNumDoc &"'    " & _
             " UNION SELECT * FROM RSFACCAR..FT0012FACC                           " & _
             " WHERE F5_CNUMSER = '"&ORISER&"' AND F5_CNUMDOC = '"&oriNumDoc &"')NN "	
'response.write(CADORI)
set rsOrigen = RsNuevo
rsOrigen.open cadOri,cnn


if rsOrigen.recordcount = 0 then
	response.write("No se reconoce el documento de Origen")
	response.end
else
	rsOrigen.movefirst		
end if



'FILA 7  A:Número de documento
FILA7A = oriSer&"-"&oriNumDoc
'FILA 7  B:Tipo de documento
FILA7B = oriTipDoc
'FILA 7  C:Codigo de tipo de nota de credito
FILA7C = "07"
'FILA 7  D:Motivo de tipo de nota de credito
FILA7D = "Devolucion por item"



'FILA 7  E:Fecha de emisión
FILA7E = year(rsOrigen("F5_DFECDOC"))&"-"&right("00"&month(rsOrigen("F5_DFECDOC")),2)&"-"&right("00"&day(rsOrigen("F5_DFECDOC")),2)
'FILA7E = "2019-09-15"









'FILA 7  F:RELATED_DOC
FILA7F = "RELATED_DOC"

FILA7 	= FILA7A
FILA7  	= FILA7 &","&FILA7B
FILA7  	= FILA7 &","&FILA7C
FILA7  	= FILA7 &","&FILA7D
FILA7  	= FILA7 &","&FILA7E
FILA7  	= FILA7 &","&FILA7F

response.write("<br/>")
response.write(utf8_simbom(FILA7))






'FILA ULTIMA (Datos de la lina)
fultima = "FF00FF"

dim fs,f
set fs=Server.CreateObject("Scripting.FileSystemObject")
'set f=fs.CreateTextFile("d:\VENTAS_NEW\efact\"&nombre&".csv",true)
'set f=fs.CreateTextFile("d:\efact\daemon\documents\"&nombre&".csv",true)
set f=fs.CreateTextFile("d:\EFACT_MODULO\daemon\documents\in\creditnote\"&nombre&".csv",true)

'LINEA 1'
f.WriteLine(utf8_simbom(fila1))
'LINEA 2'
f.WriteLine(utf8_simbom(fila2))
'LINEA 3'
f.WriteLine(utf8_simbom(fila3))
'LINEA 4'
f.WriteLine(utf8_simbom(fila4))
'LINEA 5'
f.WriteLine(utf8_simbom(fila5))
'LINEA 6'
f.WriteLine(utf8_simbom(fila6))
'LINEA 7'
f.WriteLine(utf8_simbom(FILA7))







'************************************************************************************'
'***************************************DETALLE**************************************'
'************************************************************************************'
'FILA DETALLE   A :  Número de orden 
FILADETALLEA 	= 	""
'FILA DETALLE   B :  Unidad de medida 
FILADETALLEB 	= 	""
'FILA DETALLE   C :  Cantidad 
FILADETALLEC 	= 	""
'FILA DETALLE   D :  Descripción detallada 
FILADETALLED 	= 	""
'FILA DETALLE   E :  Precio venta unitario 
FILADETALLEE 	= 	""
'FILA DETALLE   F :  Código de precio de ventaunitario
FILADETALLEF 	= 	"01"
'FILA DETALLE   G :  Valor referencial unitario 
FILADETALLEG 	= 	""
'FILA DETALLE   H :  Código del valor referencialunitario
FILADETALLEH 	= 	""
'FILA DETALLE   I :  Monto base IGV o IVAP 
FILADETALLEI 	= 	""
'FILA DETALLE   J :  Monto total IGV o IVAP
FILADETALLEJ 	= 	""
'FILA DETALLE   K :  Afectación IGV o IVAP 
FILADETALLEK 	= 	""
'FILA DETALLE   L :  Código de tributo 
FILADETALLEL 	= 	""
'FILA DETALLE   M :  Porcentaje IGV o IVAP 
FILADETALLEM 	= 	""
'FILA DETALLE   N :  Monto base ISC 
FILADETALLEN 	= 	""
'FILA DETALLE   O :  Monto total ISC 
FILADETALLEO 	= 	""
'FILA DETALLE   P :  Código de tipos de sistema decálculo ISC
FILADETALLEP 	= 	""
'FILA DETALLE   Q :  Código tributo ISC 
FILADETALLEQ 	= 	""
'FILA DETALLE   R :  Código de Producto SUNAT 
FILADETALLER 	= 	""
'FILA DETALLE   S :  Código de producto 
FILADETALLES 	= 	""
'FILA DETALLE   T :  Valor unitario 
FILADETALLET 	= 	""
'FILA DETALLE   U :  Valor de venta 
FILADETALLEU 	= 	""
'FILA DETALLE   V :  Monto base Otros tributos 
FILADETALLEV 	= 	""
'FILA DETALLE   W :  Porcentaje Otros tributos 
FILADETALLEW 	= 	""
'FILA DETALLE   X :  Monto total Otros tributos 
FILADETALLEX 	= 	""
'FILA DETALLE   Y :  Cantidad de bolsas plastico 
FILADETALLEY 	= 	""
'FILA DETALLE   Z :  Monto unitario de la bolsade plástico
FILADETALLEZ 	= 	""
'FILA DETALLE   AA:  Monto total ICBPER 
FILADETALLEAA	= 	""
'FILA DETALLE   AB:  
FILADETALLEAB	= 	""
'FILA DETALLE   AC:  
FILADETALLEAC	= 	""
'FILA DETALLE   AD:  
FILADETALLEAD	= 	""
'FILA DETALLE   AE:  
FILADETALLEAE	= 	""
'FILA DETALLE   AF:  
FILADETALLEAF	= 	""
'FILA DETALLE   AG:  Monto total impuestos 
FILADETALLEAG	= 	""
'FILA DETALLE   AH:  Total de la Línea 
FILADETALLEAH	= 	""
'FILA DETALLE   AI:  Descuento procentaje
FILADETALLEAI	= 	""
'FILA DETALLE   AJ:  Descuento Importe 
FILADETALLEAJ	= 	""
'FILA DETALLE   AK:  Descuento 1 
FILADETALLEAK	= 	""
'FILA DETALLE   AL:  Descuento 2 
FILADETALLEAL	= 	""
'FILA DETALLE   AM:  Descuento 3 
FILADETALLEAM	= 	""
'FILA DETALLE   AN:  Código cliente 
FILADETALLEAN	= 	""
'FILA DETALLE   AO:  Lotes 
FILADETALLEAO	= 	""
'FILA DETALLE   AP:  Número guía 
FILADETALLEAP	= 	""
'FILA DETALLE   AQ:  Peso total 
FILADETALLEAQ	= 	""
'FILA DETALLE   AR:  Fecha vencimiento 
FILADETALLEAR	= 	""
'FILA DETALLE   AS:  Totales 
FILADETALLEAS	= 	""
'FILA DETALLE   AT:  Cantidades 
FILADETALLEAT	= 	""
'FILA DETALLE   AU:  N° de contrato: Ventas sectorpúblico
FILADETALLEAU	= 	""
'FILA DETALLE   AV:  Fecha de otorgamiento delcrédito
FILADETALLEAV	= 	""
'FILA DETALLE   AW:  Código del tipo de préstamo 
FILADETALLEAW	= 	""
'FILA DETALLE   AX:  Número de la partidaregistral
FILADETALLEAX	= 	""
'FILA DETALLE   AY:  Código de indicador deprimera vivienda
FILADETALLEAY	= 	""
'FILA DETALLE   AZ:  Predio: Código de ubigeo 
FILADETALLEAZ	= 	""
'FILA DETALLE   BA:  Predio: Dirección completa ydetallada
FILADETALLEBA	= 	""
'FILA DETALLE   BB:  Predio: Urbanización 
FILADETALLEBB	= 	""
'FILA DETALLE   BC:  Predio: Provincia 
FILADETALLEBC	= 	""
'FILA DETALLE   BD:  Predio: Distrito 
FILADETALLEBD	= 	""
'FILA DETALLE   BE:  Predio: Departamento 
FILADETALLEBE	= 	""
'FILA DETALLE   BF:  Código de producto GS1 
FILADETALLEBF	= 	""
'FILA DETALLE   BG:  Tipo de estructura GTIN delcódigo de producto GS1
FILADETALLEBG	= 	""
'FILA DETALLE   BH:  Porcentaje del ISC
FILADETALLEBH	= 	""


for DET = 0 TO rsDet.recordcount-1

	'INICIALIZO VARIABLES'
	'FILA DETALLE   A :  Número de orden 
	FILADETALLEA 	= 	""
	'FILA DETALLE   B :  Unidad de medida 
	FILADETALLEB 	= 	""
	'FILA DETALLE   C :  Cantidad 
	FILADETALLEC 	= 	""
	'FILA DETALLE   D :  Descripción detallada 
	FILADETALLED 	= 	""
	'FILA DETALLE   E :  Precio venta unitario 
	FILADETALLEE 	= 	""



	'FILA DETALLE   I :  Monto base IGV o IVAP 
	FILADETALLEI 	= 	""
	'FILA DETALLE   J :  Monto total IGV o IVAP
	FILADETALLEJ 	= 	""
	'FILA DETALLE   K :  Afectación IGV o IVAP 
	FILADETALLEK 	= 	""
	'FILA DETALLE   L :  Código de tributo 
	FILADETALLEL 	= 	""
	'FILA DETALLE   M :  Porcentaje IGV o IVAP 
	FILADETALLEM 	= 	""
	

	'FILA DETALLE   T :  Valor unitario 
	FILADETALLET 	= 	""
	'FILA DETALLE   U :  Valor de venta 
	FILADETALLEU 	= 	""


	'FILA DETALLE   Y :  Cantidad de bolsas plastico 
	FILADETALLEY 	= 	""
	'FILA DETALLE   Z :  Monto unitario de la bolsade plástico
	FILADETALLEZ 	= 	""
	'FILA DETALLE   AA:  Monto total ICBPER 
	FILADETALLEAA	= 	""



	'FILA DETALLE   AG:  Monto total impuestos 
	FILADETALLEAG	= 	""
	'FILA DETALLE   AH:  Total de la Línea 
	FILADETALLEAH	= 	""

	'FILA DETALLE   BH:  Porcentaje del ISC
	FILADETALLEBH	= 	""



	'ASIGNO VALOR REAL A LAS VARIABLES'



		DESCRIP = ""
		DESCRIP = AlphaNumericOnly(ucase(replace(replace(trim(rsDet("f6_cdescri")),",",""),"/"," ")))

		cadTXT = " SELECT  F6_CITEM,'  %5D  '+F6_CDESCRI as F6_CDESCRI FROM (select * from RSFACCAR..FT0012FACD  where F6_CTD = '"&CTD&"' AND F6_CNUMSER = '"&SER&"' AND F6_CNUMDOC = '"&DOC&"' union  " & _
             " select * from RSFACCAR..FT0012ACUD  where F6_CTD = '"&CTD&"' AND F6_CNUMSER = '"&SER&"' AND F6_CNUMDOC = '"&DOC&"') XX where  upper(F6_CCODIGO) = upper('txt')  and convert(int,F6_CITEM) >  "&cint(rsDet("f6_citem"))&"  and F6_CITEM <  isnull((select top 1 F6_CITEM FROM (     SELECT *    FROM RSFACCAR..FT0012FACD    WHERE F6_CTD = '"&CTD&"' AND F6_CNUMSER = '"&SER&"' AND F6_CNUMDOC = '"&DOC&"'    UNION    SELECT *    FROM RSFACCAR..FT0012ACUD    WHERE F6_CTD = '"&CTD&"' AND F6_CNUMSER = '"&SER&"' AND F6_CNUMDOC = '"&DOC&"') XX WHERE UPPER(F6_CCODIGO)  <> UPPER('txt') and convert(int,F6_CITEM) > "&cint(rsDet("f6_citem"))&" ORDER BY F6_CITEM),999) ORDER BY F6_CITEM; "
			

		    set rsTXT = RsNuevo
		    if rsTXT.state > 0 then
		    	rsTXT.close
		    end if
			rsTXT.open cadTXT,cnn


			if rsTXT.recordcount > 0 then
				'response.write(cadTXT)
				for txtcont = 0 to rsTXT.recordcount-1 
					DESCRIP = DESCRIP & rsTXT("F6_CDESCRI")
					rsTXT.movenext
				next
			end if

			if I = rsDet.recordcount-1  then
				cadDetCant = " SELECT case isnull(sum(f6_ncantid),1) WHEN 0 then 1 else isnull(sum(f6_ncantid),1) end as sumCanTid FROM (select * from RSFACCAR..FT0012FACD  where F6_CTD = '"&CTD&"' AND F6_CNUMSER = '"&SER&"' AND F6_CNUMDOC = '"&DOC&"' union  " & _
	             " select * from RSFACCAR..FT0012ACUD  where F6_CTD = '"&CTD&"' AND F6_CNUMSER = '"&SER&"' AND F6_CNUMDOC = '"&DOC&"') XX  where upper(F6_CCODIGO) <> upper('txt') "
				response.write(cadDetCant)
			    set rsDetCant = RsNuevo
				rsDetCant.open cadDetCant,cnn
				if rsDetCant.recordcount > 0 then
					'DESCRIP = DESCRIP&"  %5D  Total articulos: "&(rsDetCant("sumCanTid"))
				end if
			end if
	'if abs(cdbl(rsDet("F6_NCANTID"))) < 1 then
''		response.write("Las unidades vendidas no pueden ser 0 !!")
''		response.end
''	end if


CANTIDADDET = abs(cdbl(rsDet("f6_ncantid")))

	'FILA DETALLE   A :  Número de orden 
	FILADETALLEA 	= 	DET+1
	'FILA DETALLE   B :  Unidad de medida 
	FILADETALLEB 	= 	"C62"		'APLICANDO CODIGO INTERNACIONAL'	
	IF CDBL(CANTIDADDET) = 0 THEN
		CANTIDADDET=1
	END IF
	'FILA DETALLE   C :  Cantidad 
	FILADETALLEC 	= CANTIDADDET	
	


	'FILA DETALLE   D :  Descripción detallada 
	FILADETALLED 	= 	DESCRIP
	'FILA DETALLE   E :  Precio venta unitario 
	FILADETALLEE 	= 	abs(formatnumber( ( CDBL(rsDet("f6_nimpmn"))   ) / CANTIDADDET,2,,,false ))



	'FILA DETALLE   I :  Monto base IGV o IVAP 
	FILADETALLEI 	= 	abs(formatnumber(CDBL(rsDet("f6_nimpmn")) - CDBL(rsDet("f6_nigv")),2,,,false))
	'FILA DETALLE   J :  Monto total IGV o IVAP
	FILADETALLEJ 	= 	abs(formatnumber(CDBL(rsDet("f6_nigv")),2,,,false) )
	'FILA DETALLE   K :  Afectación IGV o IVAP 
	FILADETALLEK 	= 	"10"
	'FILA DETALLE   L :  Código de tributo 
	FILADETALLEL 	= 	"1000"
	'FILA DETALLE   M :  Porcentaje IGV o IVAP 
	FILADETALLEM 	= 	abs(formatnumber(cdbl(rsDet("f6_nigvpor")),0,,,false))
	
	'FILA DETALLE   S :  Código de producto 
	FILADETALLES 	= 	ucase(trim(rsDet("f6_ccodigo")))


	'FILA DETALLE   T :  Valor unitario 
	FILADETALLET 	= 	abs(formatnumber( ( abs(CDBL(rsDet("f6_nimpmn")))  - abs(CDBL(rsDet("f6_nigv")))  ) / CANTIDADDET,2,,,false ))
	'FILA DETALLE   U :  Valor de venta 
	FILADETALLEU 	= 	abs(formatnumber(CDBL(FILADETALLET)*cdbl(FILADETALLEC) - abs(cdbl(rsDet("f6_ndescto"))),2,,,false))


	'FILA DETALLE   Y :  Cantidad de bolsas plastico 
	FILADETALLEY 	= 	0
	'FILA DETALLE   Z :  Monto unitario de la bolsade plástico
	FILADETALLEZ 	= 	0
	'FILA DETALLE   AA:  Monto total ICBPER 
	FILADETALLEAA	= 	0







	'FILA DETALLE   AG:  Monto total impuestos 
	FILADETALLEAG	= 	abs(formatnumber(CDBL(rsDet("f6_nigv")),2,,,false))
	'FILA DETALLE   AH:  Total de la Línea 
	FILADETALLEAH	= 	abs(formatnumber(CDBL(rsDet("f6_nimpmn")) - CDBL(rsDet("f6_nigv")),2,,,false))
	
	'FILA DETALLE   BH:  Porcentaje del ISC
	FILADETALLEBH	= 	"0"



	'SOLO APARECERA SI HAY DESCUENTO
	if cdbl(rsDet("f6_npordes")) > 0 then
		FILADETALLEAI					= abs(formatnumber(cdbl(rsDet("f6_npordes"))/100,2,,,false))
		FILADETALLEAJ					= abs(formatnumber(cdbl(rsDet("f6_ndescto")),2,,,false))
	end if
	

	

	FILADETALLE = FILADETALLEA 
	FILADETALLE = FILADETALLE&","&FILADETALLEB 
	FILADETALLE = FILADETALLE&","&FILADETALLEC 
	FILADETALLE = FILADETALLE&","&FILADETALLED 
	FILADETALLE = FILADETALLE&","&FILADETALLEE 
	FILADETALLE = FILADETALLE&","&FILADETALLEF 
	FILADETALLE = FILADETALLE&","&FILADETALLEG 
	FILADETALLE = FILADETALLE&","&FILADETALLEH 
	FILADETALLE = FILADETALLE&","&FILADETALLEI 
	FILADETALLE = FILADETALLE&","&FILADETALLEJ 
	FILADETALLE = FILADETALLE&","&FILADETALLEK 
	FILADETALLE = FILADETALLE&","&FILADETALLEL 
	FILADETALLE = FILADETALLE&","&FILADETALLEM 
	FILADETALLE = FILADETALLE&","&FILADETALLEN 
	FILADETALLE = FILADETALLE&","&FILADETALLEO 
	FILADETALLE = FILADETALLE&","&FILADETALLEP 
	FILADETALLE = FILADETALLE&","&FILADETALLEQ 
	FILADETALLE = FILADETALLE&","&FILADETALLER 
	FILADETALLE = FILADETALLE&","&FILADETALLES 
	FILADETALLE = FILADETALLE&","&FILADETALLET 
	FILADETALLE = FILADETALLE&","&FILADETALLEU 
	FILADETALLE = FILADETALLE&","&FILADETALLEV 
	FILADETALLE = FILADETALLE&","&FILADETALLEW 
	FILADETALLE = FILADETALLE&","&FILADETALLEX 
	FILADETALLE = FILADETALLE&","&FILADETALLEY 
	FILADETALLE = FILADETALLE&","&FILADETALLEZ 
	FILADETALLE = FILADETALLE&","&FILADETALLEAA
	FILADETALLE = FILADETALLE&","&FILADETALLEAB
	FILADETALLE = FILADETALLE&","&FILADETALLEAC
	FILADETALLE = FILADETALLE&","&FILADETALLEAD
	FILADETALLE = FILADETALLE&","&FILADETALLEAE
	FILADETALLE = FILADETALLE&","&FILADETALLEAF
	FILADETALLE = FILADETALLE&","&FILADETALLEAG
	FILADETALLE = FILADETALLE&","&FILADETALLEAH
	FILADETALLE = FILADETALLE&","&FILADETALLEAI
	FILADETALLE = FILADETALLE&","&FILADETALLEAJ
	FILADETALLE = FILADETALLE&","&FILADETALLEAK
	FILADETALLE = FILADETALLE&","&FILADETALLEAL
	FILADETALLE = FILADETALLE&","&FILADETALLEAM
	FILADETALLE = FILADETALLE&","&FILADETALLEAN
	FILADETALLE = FILADETALLE&","&FILADETALLEAO
	FILADETALLE = FILADETALLE&","&FILADETALLEAP
	FILADETALLE = FILADETALLE&","&FILADETALLEAQ
	FILADETALLE = FILADETALLE&","&FILADETALLEAR
	FILADETALLE = FILADETALLE&","&FILADETALLEAS
	FILADETALLE = FILADETALLE&","&FILADETALLEAT
	FILADETALLE = FILADETALLE&","&FILADETALLEAU
	FILADETALLE = FILADETALLE&","&FILADETALLEAV
	FILADETALLE = FILADETALLE&","&FILADETALLEAW
	FILADETALLE = FILADETALLE&","&FILADETALLEAX
	FILADETALLE = FILADETALLE&","&FILADETALLEAY
	FILADETALLE = FILADETALLE&","&FILADETALLEAZ
	FILADETALLE = FILADETALLE&","&FILADETALLEBA
	FILADETALLE = FILADETALLE&","&FILADETALLEBB
	FILADETALLE = FILADETALLE&","&FILADETALLEBC
	FILADETALLE = FILADETALLE&","&FILADETALLEBD
	FILADETALLE = FILADETALLE&","&FILADETALLEBE
	FILADETALLE = FILADETALLE&","&FILADETALLEBF
	FILADETALLE = FILADETALLE&","&FILADETALLEBG
	FILADETALLE = FILADETALLE&","&FILADETALLEBH
	response.write("<br/>")
	response.write(FILADETALLE)
	f.WriteLine(FILADETALLE)
	rsDet.movenext
next




'response.end
'LINEA ULTIMA'

response.write("<br/>")
response.write(fultima)
f.WriteLine(fultima)

f.close
set f=nothing
set fs=nothing



	'Response.Redirect "/apijf/public/index.php/auth?tipdoc=invoice&ope="&ope&"&file="&nombre&".xml"

	'response.end
	cadCntOpeReal = "SELECT operacion FROM REAL_POS order by 1 desc"
	set rsCntOpeReal = RsNuevo
	rsCntOpeReal.open cadCntOpeReal,cnn
	if rsCntOpeReal.recordcount=0 then
		ope = 1
	else
		ope = cint(right(rsCntOpeReal("operacion"),9)) + 1
	end if


	operacion = "R"&right("000000000"&ope,9)

	response.write(operacion)
	cnn.execute("insert into REAL_POS values('"&operacion&"','NC','"&SER&"','"&DOC&"');")



Sleep(2)

Response.Redirect "/apijf/public/index.php/auth?tipdoc=creditnote&ope="&ope&"&file="&nombre&".xml"




%>  
