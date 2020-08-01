<%@ Language=VBScript %>
<!--#include file="includes/Cnn.inc"-->
<!--#include file="includes/funcionesVBscript.asp"-->
<!--#include file="funcionesefact.asp"-->
<%  

Response.CharSet = "UTF-8"
ope = request.querystring("ope")

cad = "select * from movimcab where OPERACION='"&ope&"'"
if rs.state > 0 then 
	rs.close
end if

rs.open cad,cnn

if rs.recordcount > 0 then 
	rs.movefirst

	'VALIDANDO SI ES UNA BOLETA'
	if rs("coddoc") <> "FC" then
		response.write("esto no es una boleta")
		response.end
	end if

	cadDet = "select d.*,l.AR_CDESCRI as descri from movimdet d full outer join rsfaccar..AL0012ARTI l on d.CODART = l.AR_CCODIGO where OPERACION='"&ope&"'"

	set rsDet = RsNuevo
	rsDet.open cadDet,cnn

	cadCli = "select * from CLIENTES where CLIENTE='"&trim(RS("CLIENTE"))&"' and estado ='A' "
	'response.write(cadCli)
	set rsCLI = RsNuevo
	rsCLI.open cadCli,cnn


	if rsCLI.recordcount = 0 then
		response.write("Verifique y corrija el cliente.")
		response.end
	else
		rsCLI.movefirst
	end if


	'globales'
	'=============================================================================================================='
	'=============================================================================================================='
	'=============================================================================================================='
	coddoc = "01"

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
	nombre = miRuc&"-"&coddoc&"-"&ucase(""&right(trim(rs("serie")),4))&"-"&ucase(right(trim(rs("numdoc")),7))



	pvp 	= FORMATNUMBER(cdbl(RS("pvp")),2,,,false)
	SUBTOTAL 	= FORMATNUMBER(cdbl(RS("SUBTOT")),2,,,false)
	IGV 		= FORMATNUMBER(cdbl(RS("IGV")),2,,,false)
	ICBPER 		= formatnumber(cdbl(RS("isc")),2,,,false)
	Total 		= FORMATNUMBER(cdbl(RS("TOTAL")),2,,,false)
	MON 		= "PEN"
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
	clienteDOC 			= trim(rsCLI("cliente"))

''	clienteDOC = "72218501"

	if len(clienteDOC) = 8 then
		clienteTipDoc		= "1"
		response.end("TIENE QUE SER RUC!!")
	elseif len(clienteDOC) = 11 then
		clienteTipDoc		= "6"
		
	else
		clienteTipDoc		= "0"
	end if
	clienteDireccion	= trim(rsCLI("direccion"))
	clienteRazon		= trim(rsCLI("nombre"))
	clienteMail			= trim(rsCLI("mail"))





	'FILA1
	'=============================================================================================================='
	'=============================================================================================================='
	'=============================================================================================================='

	
	'FILA 1 A: * Fecha Emisión'
	f1Fecdoc = CStr(year(RS("FECDOC"))&"-"&right("00"&month(RS("FECDOC")),2)&"-"&right("00"&day(RS("FECDOC")),2))
	
	'FILA 1 B: * Número de documento'
	f1Numdoc = ucase(""&right(trim(rs("serie")),4))&"-"&ucase(trim(rs("numdoc")))
	
	'FILA 1 C: * Tipo de documento'
	f1Coddoc 								=coddoc
	
	'FILA 1 D: * Tipo de MONEDA'
	f1TipMon 								=MON
	
	'FILA 1 E: * Sumatoria Monto base IGV'
	'f1SumMontoBaseIGV						=SUBTOTAL
	f1SumMontoBaseIGV						=pvp
	
	'FILA 1 F: * Importe total IGV o IVAP 
	f1ImporteIGV							=IGV
	
	'FILA 1 G: * Tipo moneda IGV 
	f1TipMonIgv								=MON
	
	'FILA 1 H: Sumatorio monto base ISC 
	f1SumMontoBaseISC						=""
	
	'FILA 1 I: Sumatoria monto total ISC 
	f1SumMontoTotalISC						=""
	
	'FILA 1 J: Tipo moneda ISC 
	f1TipMonISC								=""
	
	'FILA 1 K: Sumatoria monto base OTROS
	f1SumMontoBaseOtros						=""
	
	'FILA 1 L: Sumatoria monto total OTROS
	f1SumMontoTotalOtros					=""
	
	'FILA 1 M: Tipo moneda OTROS
	f1TipMonOtros							=""
	
	'FILA 1 N: * Importe total comprobante
	f1ImporteTotal							=TOTAL
	
	'FILA 1 O: O
	f1O										=""
	
	'FILA 1 P: p
	f1P										=""
	
	'FILA 1 Q: * Tipo Operacion (VENTA INTERNA)
	f1TipOperacion							="0101"
	
	'FILA 1 R: R
	f1R										=""
	
	'FILA 1 S: Sumatoria monto total ICBPER
	f1S										=ICBPER
	
	'FILA 1 T: Sumatoria de impuestos deoperaciones gratuitas :(
	f1T										=""
	
	'FILA 1 U: total Operaciones exportacion
	f1TotalOpeExpor							=""
	
	'FILA 1 V: Total de operaciones gravadas
	f1TotalOpeGrav 							=SUBTOTAL
	
	'FILA 1 W: Total Operacion Inafectas
	f1TotalOpeInaf							=""
	
	'FILA 1 X: Total operaciones Exoneradas
	f1TotalOpeExon							=""
	
	'FILA 1 Y: Total de operaciones Gratuitas
	f1TotalOpeGrat							=""
	
	'FILA 1 Z: Z
	f1Z										=""
	
	'FILA 1 AA: Monto Base percepcion
	f1MontoBasPercep						=""
	
	'FILA 1 AB: Monto total Percepcion
	f1MontoTotPercep						=""
	
	'FILA 1 AC: Importe total Incluido Percepcion
	f1ImporteTotalPercep					=""
	
	'FILA 1 AD: Codig de bien o Servicio de detraccion
	f1CodDetraccion							=""
	
	'FILA 1 AE: Monto detraido de la detraccion
	f1MontoDetraDetraccion					=""
	
	'FILA 1 AF: Tasa o porcentaje de la detraccion
	f1ProcenDetra							=""
	
	'FILA 1 AG: Numero de banco de la detraccion
	f1NroBancoDetracc						=""
	
	'FILA 1 AH: Importe total incluido la detraccion
	f1ImporteTotalDetracc						=""
	
	'FILA 1 AI: Ai
	f1AI										=""
	
	'FILA 1 AJ: aj
	f1AJ										=""
	
	'FILA 1 AK: ak
	f1AK										=""
	
	'FILA 1 AL: cantidad de lineas del documento
	f1CantLineasDoc								=cntLinDet
	
	'FILA 1 AM: codigo del regimen de la percepcion
	f1CodRegPerc								=""
	
	'FILA 1 AN: cantidad guias o otros documentos asociadas  al documentos
	f1CantGuias									=""
	
	'FILA 1 AO: cantidad de anticipos asociadas
	f1CantAnticipos								=""
	
	'FILA 1 AP: Totalanticipos
	f1TotalAnticipos							=""
	
	'FILA 1 AQ: cantidadpuntodepartida y llegada
	f1CantPuntPartiLlegada						=""
	
	'FILA 1 AR: monto descuento global ab
	f1MontoDescGlobalAB							=""
	
	'FILA 1 AS: monto descuento global no ab
	f1MontoDescGlobalNOAB						=""
	
	'FILA 1 AT: monto anticipo gravado igv o ivap
	f1MontoAnticipoGravadoIGV					=""
	
	'FILA 1 AU: monto anticipo exonerado
	f1MontoAnticipoExon							=""
	
	'FILA 1 AV: monto anticipo inafecto
	f1MontoAnticipoInaf							=""
	
	'FILA 1 AW: monto base FISE
	f1MontoBaseFise								=""
	
	'FILA 1 AX: monto total FISE
	f1MontoTotalFise							=""
	
	'FILA 1 AY: Recargo al consumo y/opropinas
	f1RecargoPropina							=""
	
	'FILA 1 AZ: Monto cargo global AB 
	f1MontoCargoGlobalAB						=""
	
	'FILA 1 BA: Monto cargo global no AB 
	f1MontoCargoGlobalNOAB						=""
	
	'FILA 1 BB: Monto total de impuestos 
	f1MontoTotalImpuestos						=formatnumber(cdbl(IGV)+cdbl(ICBPER),2,,,false)
	
	'FILA 1 BC: Total Valor de Venta 
	f1TotalValorVenta							=SUBTOTAL
	
	'FILA 1 BD: Total precio de venta 
	f1TotalPercioVenta							=TOTAL
	
	'FILA 1 BE: Total descuentos(Que no afectan la base)
	f1TotalDescuentos							=""
	
	'FILA 1 BF: Total cargos(Que no afectan la base)
	f1TotalCargos								=""
	
	'FILA 1 BG: Monto para Redondeo del Importe Total
	f1MontoParaRedondeoImporteTotal				=""

	'FILA 1 BH: Total descuentos AB
	f1TotalDescuentosAB			=FORMATNUMBER(RS("DESCUENTO"),2,,,FALSE)


	fila1=""&f1Fecdoc &","
	fila1=fila1 & ""&f1Numdoc&","
	fila1=fila1 & ""&f1Coddoc&","
	fila1=fila1 & ""&f1TipMon&","
	fila1=fila1 & ""&f1SumMontoBaseIGV&","
	fila1=fila1 & ""&f1ImporteIGV&","
	fila1=fila1 & ""&f1TipMonIgv&","
	fila1=fila1 & ""&f1SumMontoBaseISC&","
	fila1=fila1 & ""&f1SumMontoTotalISC&","
	fila1=fila1 & ""&f1TipMonISC&","
	fila1=fila1 & ""&f1SumMontoBaseOtros&","
	fila1=fila1 & ""&f1SumMontoTotalOtros&","
	fila1=fila1 & ""&f1TipMonOtros&","
	fila1=fila1 & ""&f1ImporteTotal&","
	fila1=fila1 & ""&f1O&","
	fila1=fila1 & ""&f1P&","
	fila1=fila1 & ""&f1TipOperacion&","
	fila1=fila1 & ""&f1R&","
	fila1=fila1 & ""&f1S&","
	fila1=fila1 & ""&f1T&","
	fila1=fila1 & ""&f1TotalOpeExpor&","
	fila1=fila1 & ""&f1TotalOpeGrav&","
	fila1=fila1 & ""&f1TotalOpeInaf&","
	fila1=fila1 & ""&f1TotalOpeExon&","
	fila1=fila1 & ""&f1TotalOpeGrat&","
	fila1=fila1 & ""&f1Z&","
	fila1=fila1 & ""&f1MontoBasPercep&","
	fila1=fila1 & ""&f1MontoTotPercep&","
	fila1=fila1 & ""&f1ImporteTotalPercep&","
	fila1=fila1 & ""&f1CodDetraccion&","
	fila1=fila1 & ""&f1MontoDetraDetraccion&","
	fila1=fila1 & ""&f1ProcenDetra&","
	fila1=fila1 & ""&f1NroBancoDetracc&","
	fila1=fila1 & ""&f1ImporteTotalDetracc&","
	fila1=fila1 & ""&f1AI&","
	fila1=fila1 & ""&f1AJ&","
	fila1=fila1 & ""&f1AK&","
	fila1=fila1 & ""&f1CantLineasDoc&","
	fila1=fila1 & ""&f1CodRegPerc&","
	fila1=fila1 & ""&f1CantGuias&","
	fila1=fila1 & ""&f1CantAnticipos&","
	fila1=fila1 & ""&f1TotalAnticipos&","
	fila1=fila1 & ""&f1CantPuntPartiLlegada&","
	fila1=fila1 & ""&f1MontoDescGlobalAB&","
	fila1=fila1 & ""&f1MontoDescGlobalNOAB&","
	fila1=fila1 & ""&f1MontoAnticipoGravadoIGV&","
	fila1=fila1 & ""&f1MontoAnticipoExon&","
	fila1=fila1 & ""&f1MontoAnticipoInaf&","
	fila1=fila1 & ""&f1MontoBaseFise&","
	fila1=fila1 & ""&f1MontoTotalFise&","
	fila1=fila1 & ""&f1RecargoPropina&","
	fila1=fila1 & ""&f1MontoCargoGlobalAB&","
	fila1=fila1 & ""&f1MontoCargoGlobalNOAB&","
	fila1=fila1 & ""&f1MontoTotalImpuestos&","
	fila1=fila1 & ""&f1TotalValorVenta&","
	fila1=fila1 & ""&f1TotalPercioVenta&","
	fila1=fila1 & ""&f1TotalDescuentos&","
	fila1=fila1 & ""&f1TotalCargos&","
	fila1=fila1 & ""&f1MontoParaRedondeoImporteTotal&","
	fila1=fila1 & ""&f1TotalDescuentosAB


	response.write(fila1)
	
	'FILA2
	'=============================================================================================================='
	'=============================================================================================================='
	'=============================================================================================================='

	'FILA 2 A:Código ubigeo punto llegada 
	f2CodUbigeo 							= ""
	'FILA 2 B:Dirección punto llegada 
	f2DireccionPuntoLLegada 				= ""
	'FILA 2 C:Urbanizacion punto llegada 
	f2UrbanizacionPuntoLLegada 				= ""
	'FILA 2 D:Provincia punto llegada 
	f2ProvinciaPuntoLLegada 				= ""
	'FILA 2 E:Departamento punto llegada 
	f2DepartamentoPuntoLLegada 				= ""
	'FILA 2 F:Distrito punto llegada 
	f2DistritoPuntoLLegada 					= ""
	'FILA 2 G:Código país punto llegada
	f2CodPaisPuntoLLegada 					= ""


	'Fila 2 H :Código ubigeo punto llegada 
	Fila2H									=""
	'Fila 2 I :Dirección punto llegada 
	Fila2I									=""
	'Fila 2 J :Urbanizacion punto llegada 
	Fila2J									=""
	'Fila 2 K :Provincia punto llegada 
	Fila2K									=""
	'Fila 2 L :Departamento punto llegada 
	Fila2L									=""
	'Fila 2 M :Distrito punto llegada 
	Fila2M									=""
	'Fila 2 N :Código país punto llegada 
	Fila2N									=""
	'Fila 2 O :Número de placa 
	Fila2O									=""
	'Fila 2 P :Autorización del vehículo
	Fila2P									=""
	'Fila 2 Q :Marca del vehículo 
	Fila2Q									=""
	'Fila 2 R :Número de licencia Conductor 
	Fila2R									=""
	'Fila 2 S :RUC del transportista 
	Fila2S									=""
	'Fila 2 T :Tipo documento del transportista 
	Fila2T									=""
	'Fila 2 U :Razón social del transportista 
	Fila2U									=""
	'Fila 2 V :Registro MTC del transportista 
	Fila2V									=""
	'Fila 2 W :Código de motivo de traslado 
	Fila2W									=""
	'Fila 2 X :Descripcion motivo de traslado 
	Fila2X									=""
	'Fila 2 Y :Peso bruto total 
	Fila2Y									=""
	'Fila 2 Z :Código modalidad de transporte 
	Fila2Z									=""
	'Fila 2 AA :Descripción modalidad detransporte
	Fila2AA									=""
	'Fila 2 AB :Fecha de inicio del traslado ofecha de entrega de bienes altransportista
	Fila2AB									=""
	'Fila 2 AC :Número de documento delconductor
	Fila2AC									=""
	'Fila 2 AD :Tipo de documento del conductor 
	Fila2AD									=""
	'Fila 2 AE :Nombres y apellidos delconductor
	Fila2AE									=""
	'Fila 2 AF :Indicador de subcontratación
	Fila2AF									=""




	fila2=f2CodUbigeo
	fila2 = fila2 &","&f2DireccionPuntoLLegada
	fila2 = fila2 &","&f2UrbanizacionPuntoLLegada
	fila2 = fila2 &","&f2ProvinciaPuntoLLegada
	fila2 = fila2 &","&f2DepartamentoPuntoLLegada
	fila2 = fila2 &","&f2DistritoPuntoLLegada
	fila2 = fila2 &","&f2CodPaisPuntoLLegada
	fila2 = fila2 &","&Fila2H	
	fila2 = fila2 &","&Fila2I	
	fila2 = fila2 &","&Fila2J	
	fila2 = fila2 &","&Fila2K	
	fila2 = fila2 &","&Fila2L	
	fila2 = fila2 &","&Fila2M	
	fila2 = fila2 &","&Fila2N	
	fila2 = fila2 &","&Fila2O	
	fila2 = fila2 &","&Fila2P	
	fila2 = fila2 &","&Fila2Q	
	fila2 = fila2 &","&Fila2R	
	fila2 = fila2 &","&Fila2S	
	fila2 = fila2 &","&Fila2T	
	fila2 = fila2 &","&Fila2U	
	fila2 = fila2 &","&Fila2V	
	fila2 = fila2 &","&Fila2W	
	fila2 = fila2 &","&Fila2X	
	fila2 = fila2 &","&Fila2Y	
	fila2 = fila2 &","&Fila2Z	
	fila2 = fila2 &","&Fila2AA
	fila2 = fila2 &","&Fila2AB
	fila2 = fila2 &","&Fila2AC
	fila2 = fila2 &","&Fila2AD
	fila2 = fila2 &","&Fila2AE
	fila2 = fila2 &","&Fila2AF
	response.write("<br/>")
	response.write(fila2)


	'FILA3
	'=============================================================================================================='
	'=============================================================================================================='
	'=============================================================================================================='

	
	'FILA 3 A: Tipo moneda anticipo 
	f3TipMonAnti							=""
	'FILA 3 B: Monto anticipado 
	f3MontoAnti								=""
	'FILA 3 C: Tipo de anticipo 
	f3TipAnti 								=""
	'FILA 3 D: Serie y correlativoanticipo
	f3SerieCorre 							=""
	'FILA 3 E: Fecha PAGO anticipo
	f3PagoEmi 								=""
	'FILA 3 F: PREPAID_DOC
	f3Prepaid_doc 							=""

	fila3 = f3TipMonAnti
	fila3 = fila3 & ","&f3MontoAnti
	fila3 = fila3 & ","&f3TipAnti 
	fila3 = fila3 & ","&f3SerieCorre
	fila3 = fila3 & ","&f3PagoEmi 
	fila3 = fila3 & ","&f3Prepaid_doc

	response.write("<br/>")
	response.write(fila3)





	'FILA4
	'=============================================================================================================='
	'=============================================================================================================='
	'=============================================================================================================='

	'FILA 4 A:Numero de guia 
	f4NroGuia 								=""
	'FILA 4 B:Código de la guia 
	f4CodGuia 								=""
	'FILA 4 C:Número otrodocumento
	f4NroOtroDoc							=""
	'FILA 4 D:Código del tipo otrodocumento
	f4CodOtroDoc							=""
	'FILA 4 E:ATTACHDOC
	f4AttachDoc								=""

	fila4 = f4NroGuia
	fila4 = fila4 &","&f4CodGuia
	fila4 = fila4 &","&f4NroOtroDoc
	fila4 = fila4 &","&f4CodOtroDoc
	fila4 = fila4 &","&f4AttachDoc
	response.write("<br/>")
	response.write(fila4)




	'FILA5
	'=============================================================================================================='
	'=============================================================================================================='
	'=============================================================================================================='

	'FILA 5 A:Apellidos y nombres,denominación o razónsocial
	f5RazonSocial 							=miRS
	'FILA 5 B:Nombre comercial 
	f5NomComercial 							=miNombreComercial
	'FILA 5 C:Número de RUC 
	f5NumRuc	 							=miRuc
	'FILA 5 D:Código Ubigeo 
	f5CodUbigeo 							=codUBIGUEO
	'FILA 5 E:Dirección 
	f5Direccion 							=miDireccion
	'FILA 5 F:Urbanización 
	f5Urbanizacion 							=""
	'FILA 5 G:Departamento 
	f5Departamento 							="LIMA"
	'FILA 5 H:Provincia 
	f5Provincia 							="LIMA"
	'FILA 5 I:Distrito 
	f5Distrito 								="ATE"
	'FILA 5 J:Codigo del pais 
	f5CodPais	 							="PE"
	'FILA 5 K:Código delestablecimiento
	f5CodEst	 							="0000"




	fila5 = f5RazonSocial
	fila5 = fila5 &","&f5NomComercial
	fila5 = fila5 &","&f5NumRuc	
	fila5 = fila5 &","&f5CodUbigeo
	fila5 = fila5 &","&f5Direccion
	fila5 = fila5 &","&f5Urbanizacion
	fila5 = fila5 &","&f5Departamento
	fila5 = fila5 &","&f5Provincia
	fila5 = fila5 &","&f5Distrito
	fila5 = fila5 &","&f5CodPais	
	fila5 = fila5 &","&f5CodEst
	response.write("<br/>")
	response.write(fila5)

	
	'FILA6
	'=============================================================================================================='
	'=============================================================================================================='
	'=============================================================================================================='


	'FILA 6  A:Número de documento
	f6NroDoc 								=clienteDOC
	'FILA 6  B:Tipo de documento 
	f6TipDoc 								=clienteTipDoc
	'FILA 6  C:Razón social 
	f6RazonSocial 							=clienteRazon
	'FILA 6  D:Nombre comercial 
	f6NomComercial 							=""
	'FILA 6  E:Código ubigeo 
	f6CodUbigeo 							=""
	'FILA 6  F:Dirección 
	f6Direccion 							=clienteDireccion
	'FILA 6  G:Urbanización 
	f6Urbanizacion 							=""
	'FILA 6  H:Departamento 
	f6Departamento 							=""&""
	'FILA 6  I:Provincia 
	f6Provincia 							=""&""
	'FILA 6  J:Distrito 
	f6Distrito 								=""&""
	'FILA 6  K:Código de país 
	f6CodPais 								="PE"
	'FILA 6  L:Correo
	f6Correo 								=clienteMail

	fila6 = f6NroDoc
	fila6 = fila6 &","&f6TipDoc 					
	fila6 = fila6 &","&f6RazonSocial 
	fila6 = fila6 &","&f6NomComercial
	fila6 = fila6 &","&f6CodUbigeo 		
	fila6 = fila6 &","&f6Direccion 		
	fila6 = fila6 &","&f6Urbanizacion
	fila6 = fila6 &","&f6Departamento
	fila6 = fila6 &","&f6Provincia 		
	fila6 = fila6 &","&f6Distrito 			
	fila6 = fila6 &","&f6CodPais 				
	fila6 = fila6 &","&f6Correo 				

	response.write("<br/>")
	response.write(fila6)

	
	'FILA7
	'=============================================================================================================='
	'=============================================================================================================='
	'=============================================================================================================='

	f71000 = Numlet(cdbl(total))& " soles."
	f71002 = ""
	f72000 = ""
	f72001 = ""
	f72002 = ""
	f72003 = ""
	f72004 = ""
	f72005 = ""
	f72006 = ""
	f72007 = ""
	f72008 = ""
	f72009 = ""
	f72010 = ""

	fila7 = f71000
	fila7 = fila7&","&f71002
	fila7 = fila7&","&f72000
	fila7 = fila7&","&f72001
	fila7 = fila7&","&f72002
	fila7 = fila7&","&f72003
	fila7 = fila7&","&f72004
	fila7 = fila7&","&f72005
	fila7 = fila7&","&f72006
	fila7 = fila7&","&f72007
	fila7 = fila7&","&f72008
	fila7 = fila7&","&f72009
	fila7 = fila7&","&f72010
	response.write("<br/>")
	response.write(fila7)



	'FILA8
	'=============================================================================================================='
	'=============================================================================================================='
	'=============================================================================================================='


	'FILA 8  A: Observaciones 
	F8A								= "No se acepta cambios ni devoluciones en ropa interior ni prendas con descuento pijamas solo cambios con documento de vta. max 5 dias."
	'FILA 8  B: Orden de compra 
	F8B								= ""
	'FILA 8  C: Fecha de vencimiento opago
	F8C								= ""
	'FILA 8  D: Codigodecliente
	F8D								= ""
	'FILA 8  E: Códigodevendedor
	F8E								= ""
	'FILA 8  F: Motivodeventa
	F8F								= ""
	'FILA 8  G: Ordendeventa
	F8G								= ""
	'FILA 8  H: Condicióndeventa
	F8H								= ""
	'FILA 8  I: Condición general 
	F8I								= ""
	'FILA 8  J: Número interno 
	F8J								= ""
	'FILA 8  K: Número pedido 
	F8K								= ""
	'FILA 8  L: Condición de pago 
	F8L								= ""
	'FILA 8  M: Fecha de pago 
	F8M								= ""
	'FILA 8  N: Tipo de cambio 
	F8N								= ""
	'FILA 8  O: Usuario 
	F8O								= ""
	'FILA 8  P: Emitido por 
	F8P								= ""
	'FILA 8  Q: Código SAP 
	F8Q								= ""
	'FILA 8  R: Entrega factura 
	F8R								= ""
	'FILA 8  S: Sede 
	F8S								= ""
	'FILA 8  T: Ruta 
	F8T								= ""
	'FILA 8  U: Fax 
	F8U								= ""
	'FILA 8  V: Número Teléfono 
	F8V								= ""
	'FILA 8  W: Código alumno 
	F8W								= ""
	'FILA 8  X: Nombre alumno 
	F8X								= ""
	'FILA 8  Y: Sección 
	F8Y								= ""
	'FILA 8  Z: Efectivo 
	F8Z								= ""
	'FILA 8  AA: Vuelto 
	F8AA							= ""
	'FILA 8  AB: Contrato 
	F8AB							= ""
	'FILA 8  AC: Proyecto 
	F8AC							= ""
	'FILA 8  AD: Número registro 
	F8AD							= ""
	'FILA 8  AE: Certificación 
	F8AE							= ""
	'FILA 8  AF: Producto 
	F8AF							= ""
	'FILA 8  AG: Nave 
	F8AG							= ""
	'FILA 8  AH: Puerto embarque 
	F8AH							= ""
	'FILA 8  AI: Puerto destino 
	F8AI							= ""
	'FILA 8  AJ: Puerto Entrega / Delivery 
	F8AJ							= ""
	'FILA 8  AK: Numero contenedor 
	F8AK							= ""
	'FILA 8  AL: Consignatario 
	F8AL							= ""
	'FILA 8  AM: Notificante 
	F8AM							= ""
	'FILA 8  AN: Flete
	F8AN							= ""
	'FILA 8  AO: Seguro 
	F8AO							= ""
	'FILA 8  AP: Total CFR/CPT 
	F8AP							= ""
	'FILA 8  AQ: Total FOB/FCA 
	F8AQ							= ""
	'FILA 8  AR: Partida Arancelaria 
	F8AR							= ""
	'FILA 8  AS: Temperatura 
	F8AS							= ""
	'FILA 8  AT: Nro BL/AWB 
	F8AT							= ""
	'FILA 8  AU: Incoterms 
	F8AU							= ""
	'FILA 8  AV: Nro DUA 
	F8AV							= ""
	'FILA 8  AW: Peso bruto 
	F8AW							= ""
	'FILA 8  AX: Total bultos 
	F8AX							= ""
	'FILA 8  AY: Total artículos 
	F8AY							= ""
	'FILA 8  AZ: Fecha ingreso 
	F8AZ							= ""
	'FILA 8  BA: Fecha salida 
	F8BA							= ""
	'FILA 8  BB: Intereses 
	F8BB							= ""
	'FILA 8  BC: Comisiones 
	F8BC							= ""
	'FILA 8  BD: Almacén 
	F8BD							= ""
	'FILA 8  BE: Lote 
	F8BE							= ""
	'FILA 8  BF: O 
	F8BF							= ""
	'FILA 8  BG: C 
	F8BG							= ""
	'FILA 8  BH: Z - OF 
	F8BH							= ""
	'FILA 8  BI: G 
	F8BI							= ""
	'FILA 8  BJ: T/ENT 
	F8BJ							= ""
	'FILA 8  BK: Punto partida 2 
	F8BK							= ""
	'FILA 8  BL: Punto llegada 2 
	F8BL							= ""
	'FILA 8  BM: Razón social transportistFinal
	F8BM							= ""
	'FILA 8  BN: Ruc transportista Final 
	F8BN							= ""
	'FILA 8  BO: Chofer/Licenciatransportista Final
	F8BO							= ""
	'FILA 8  BP: Marca/Placa Transportistfinal
	F8BP							= ""
	'FILA 8  BQ: Número de documentootro participante
	F8BQ							= ""
	'FILA 8  BR: Tipo de documento otrosparticipante
	F8BR							= ""
	'FILA 8  BS: Apellidos y nombres deotro participante
	F8BS							= ""
	'FILA 8  BT: ID de orden de compra 
	F8BT							= ""
	'FILA 8  BU: Referencia de cliente 
	F8BU							= ""
	'FILA 8  BV: Comprador númerodocumento identidad
	F8BV							= ""
	'FILA 8  BW: Comprador tipo dedocumento identidad
	F8BW							= ""
	'FILA 8  BX: Número de RUC del Agentede Ventas
	F8BX							= ""
	'FILA 8  BY: Tipo de documento delAgente de Ventas
	F8BY							= ""
	'FILA 8  BZ: Código ubigeo entrega delbien
	F8BZ							= ""
	'FILA 8  CA: Dirección entrega del bien 
	F8CA							= ""
	'FILA 8  CB: Urbanización entrega delbien
	F8CB							= ""
	'FILA 8  CC: Provincia entrega del bien 
	F8CC							= ""
	'FILA 8  CD: Departamento entrega delbien
	F8CD							= ""
	'FILA 8  CE: Distrito entrega del bien 
	F8CE							= ""
	'FILA 8  CF: Código país entrega delbien
	F8CF							= ""
	'FILA 8  CG: Medio de pago 
	F8CG							= ""
	'FILA 8  CH: Número de autorización dela transacción
	F8CH							= ""
	'FILA 8  CI: Código del País del uso,explotación oaprovechamiento delservicio.
	F8CI							= ""








	fila8 = F8A	
	fila8 = fila8 & "," &F8B	
	fila8 = fila8 & "," &F8C	
	fila8 = fila8 & "," &F8D	
	fila8 = fila8 & "," &F8E	
	fila8 = fila8 & "," &F8F	
	fila8 = fila8 & "," &F8G	
	fila8 = fila8 & "," &F8H	
	fila8 = fila8 & "," &F8I	
	fila8 = fila8 & "," &F8J	
	fila8 = fila8 & "," &F8K	
	fila8 = fila8 & "," &F8L	
	fila8 = fila8 & "," &F8M	
	fila8 = fila8 & "," &F8N	
	fila8 = fila8 & "," &F8O	
	fila8 = fila8 & "," &F8P	
	fila8 = fila8 & "," &F8Q	
	fila8 = fila8 & "," &F8R	
	fila8 = fila8 & "," &F8S	
	fila8 = fila8 & "," &F8T	
	fila8 = fila8 & "," &F8U	
	fila8 = fila8 & "," &F8V	
	fila8 = fila8 & "," &F8W	
	fila8 = fila8 & "," &F8X	
	fila8 = fila8 & "," &F8Y	
	fila8 = fila8 & "," &F8Z	
	fila8 = fila8 & "," &F8AA
	fila8 = fila8 & "," &F8AB
	fila8 = fila8 & "," &F8AC
	fila8 = fila8 & "," &F8AD
	fila8 = fila8 & "," &F8AE
	fila8 = fila8 & "," &F8AF
	fila8 = fila8 & "," &F8AG
	fila8 = fila8 & "," &F8AH
	fila8 = fila8 & "," &F8AI
	fila8 = fila8 & "," &F8AJ
	fila8 = fila8 & "," &F8AK
	fila8 = fila8 & "," &F8AL
	fila8 = fila8 & "," &F8AM
	fila8 = fila8 & "," &F8AN
	fila8 = fila8 & "," &F8AO
	fila8 = fila8 & "," &F8AP
	fila8 = fila8 & "," &F8AQ
	fila8 = fila8 & "," &F8AR
	fila8 = fila8 & "," &F8AS
	fila8 = fila8 & "," &F8AT
	fila8 = fila8 & "," &F8AU
	fila8 = fila8 & "," &F8AV
	fila8 = fila8 & "," &F8AW
	fila8 = fila8 & "," &F8AX
	fila8 = fila8 & "," &F8AY
	fila8 = fila8 & "," &F8AZ
	fila8 = fila8 & "," &F8BA
	fila8 = fila8 & "," &F8BB
	fila8 = fila8 & "," &F8BC
	fila8 = fila8 & "," &F8BD
	fila8 = fila8 & "," &F8BE
	fila8 = fila8 & "," &F8BF
	fila8 = fila8 & "," &F8BG
	fila8 = fila8 & "," &F8BH
	fila8 = fila8 & "," &F8BI
	fila8 = fila8 & "," &F8BJ
	fila8 = fila8 & "," &F8BK
	fila8 = fila8 & "," &F8BL
	fila8 = fila8 & "," &F8BM
	fila8 = fila8 & "," &F8BN
	fila8 = fila8 & "," &F8BO
	fila8 = fila8 & "," &F8BP
	fila8 = fila8 & "," &F8BQ
	fila8 = fila8 & "," &F8BR
	fila8 = fila8 & "," &F8BS
	fila8 = fila8 & "," &F8BT
	fila8 = fila8 & "," &F8BU
	fila8 = fila8 & "," &F8BV
	fila8 = fila8 & "," &F8BW
	fila8 = fila8 & "," &F8BX
	fila8 = fila8 & "," &F8BY
	fila8 = fila8 & "," &F8BZ
	fila8 = fila8 & "," &F8CA
	fila8 = fila8 & "," &F8CB
	fila8 = fila8 & "," &F8CC
	fila8 = fila8 & "," &F8CD
	fila8 = fila8 & "," &F8CE
	fila8 = fila8 & "," &F8CF
	fila8 = fila8 & "," &F8CG
	fila8 = fila8 & "," &F8CH
	fila8 = fila8 & "," &F8CI


	response.write("<br/>")
	response.write(fila8)






	'FILA9 (Datos de la lina)
	'=============================================================================================================='
	'=============================================================================================================='
	'=============================================================================================================='



	'FILA 9  A:Número de orden
	fDetalleNroOrden 										=""

	'FILA 9  B:Unidad de medida
	fDetalleUnidMed											=""																										

	'FILA 9  C:Cantidad
	fDetalleCantid											=""										

	'FILA 9  D:Descripción detallada
	fDetalleDescrip											=""									

	'FILA 9  E:Precio venta unitario
	fDetallePreVentUnit										=""									

	'FILA 9  F:Código de precio de ventaunitario
	fDetalleCodPreVentUnit 									="01"																	

	'FILA 9  G:Valor referencial unitario
	fDetalleValorRefUnit									=""								

	'FILA 9  H:Código del valor referencial unitario
	fDetalleCodValRefUnit									=""							

	'FILA 9  I:Monto base IGV o IVAP
	fDetalleMontoBaseIgv									=""						

	'FILA 9  J:Monto total IGV o IVAP
	fDetalleMontoTotalIgv									=""					

	'FILA 9  K:Afectacion IGV
	fDetalleTipoAfectIgv									="10"				

	'FILA 9  L:Código de tributos
	fDetalleCodigoTibutos									="1000"			

	'FILA 9  M:Porcentaje IGV o IVAP
	fDetalleProcentIgv										=""	

	'FILA 9  N:Monto base ISC
	fDetalleMontBaseISC										=""

	'FILA 9  O:Monto total ISC
	fDetalleMontoTotalIsc									=""

	'FILA 9  P:Código de tipos de sistemade cálculo ISC
	fDetalleCodigoTipoSistCalcIsc							=""	

	'FILA 9  Q:Código tributo ISC
	fDetalleCodTributoISC									=""

	'FILA 9  R:Código de Producto SUNAT
	fDetalleCodProdSunat									=""									

	'FILA 9  S:Código de producto 
	fDetalleCodProd											=""										

	'FILA 9  T:Valor unitario 
	fDetalleValUnit											=""										

	'FILA 9  U:Valor de venta 
	fDetalleValVent											=""										

	'FILA 9  V:Monto base Otros tributo 
	fDetalleMontoBaseOtros									=""												

	'FILA 9  W:Porcentaje Otros tributo 
	fDetalleProcentOtros									=""												

	'FILA 9  X:Monto total Otros tributo 
	fDetalleMontoTotalOtros									=""												

	'FILA 9  Y:MONTO BASE DESCUENTO AB
	fDetalleMontoBaseDescuentoAB							=""												

	'FILA 9  Z:factor descuento ab
	fDetalleFactorDescuentoAB								=""													

	'FILA 9  AA:Monto total descuento ab 
	fDetalleMontoTotalDescuentoAB							=""														

	'FILA 9  AB:Monto base descuento NO AB 
	fDetalleMontoBaseDescuentoNOAB							=""													

	'FILA 9  AC:Factor descuento no AB 
	fDetalleFactorDescuentoNOAB								=""												

	'FILA 9  AD:Monto total descuento no AB 
	fDetalleMontoTotalDescuentoNOAB							=""												

	'FILA 9  AE:Monto base del cargo AB
	fDetalleMontBaseCargoAB									=""												

	'FILA 9  AF:Factor del cargo AB
	fDetalleFactorCargoAB									=""												

	'FILA 9  AG:Monto total del cargo AB
	fDetalleMontTotalCargoAB								=""		

	'FILA 9  AH:Monto base del cargo no AB
	fDetalleMontBaseCargoNOAB								=""											

	'FILA 9  AI:Factor del cargo no AB
	fDetalleFactorCargoNoAB 								=""

	'FILA 9  AJ:Monto total del cargo no AB
	fDetalleMontTotalCargoNOAB 								=""

	'FILA 9  AK:Monto total impuesto
	fDetalleMontTotalImp 									=""

	'FILA 9  AL:Total de la Línea
	fDetalleTotaldelaLinea 									=""


	'FILA 9  AM:Número de placa del vehículo
	FILA9AM			= ""
	'FILA 9  AN:Cantidad Und Emp 
	FILA9AN			= ""
	'FILA 9  AO:Cantidad total und 
	FILA9AO			= ""
	'FILA 9  AP:Descuento % 
	FILA9AP			= ""
	'FILA 9  AQ:Descuento importe 
	FILA9AQ			= ""
	'FILA 9  AR:Descuento 1 
	FILA9AR			= ""
	'FILA 9  AS:Descuento 2 
	FILA9AS			= ""
	'FILA 9  AT:Descuento 3 
	FILA9AT			= ""
	'FILA 9  AU:Código cliente 
	FILA9AU			= ""
	'FILA 9  AV:Lote 
	FILA9AV			= ""
	'FILA 9  AW:Peso total 
	FILA9AW			= ""
	'FILA 9  AX:Numero guia 
	FILA9AX			= ""
	'FILA 9  AY:Campo adicional 
	FILA9AY			= ""
	'FILA 9  AZ:Código país de residencia delsujeto no domiciliado
	FILA9AZ			= ""
	'FILA 9  BA:Fecha ingreso al país 
	FILA9BA			= ""
	'FILA 9  BB:Fecha ingreso alestablecimiento
	FILA9BB			= ""
	'FILA 9  BC:Fecha salida delestablecimiento
	FILA9BC			= ""
	'FILA 9  BD:Número de días depermanencia
	FILA9BD			= ""
	'FILA 9  BE:Fecha de consumo 
	FILA9BE			= ""
	'FILA 9  BF:Código país de emisión delpasaporte
	FILA9BF			= ""
	'FILA 9  BG:Apellidos y nombres o razónsocial del huésped
	FILA9BG			= ""
	'FILA 9  BH:Tipo de documento delhuésped
	FILA9BH			= ""
	'FILA 9  BI:Número documento delhuésped
	FILA9BI			= ""
	'FILA 9  BJ:N° de expediente: Ventas sectorpúblico
	FILA9BJ			= ""
	'FILA 9  BK:Código unidad ejecutora:Ventas sector público
	FILA9BK			= ""
	'FILA 9  BL:N° de contrato: Ventas sectorpúblico
	FILA9BL			= ""
	'FILA 9  BM:N° de proceso de selección:Ventas sector público
	FILA9BM			= ""
	'FILA 9  BN:N° de contrato: Ventas sectorpúblico
	FILA9BN			= ""
	'FILA 9  BO:Fecha de otorgamiento delcrédito
	FILA9BO			= ""
	'FILA 9  BP:Código del tipo de préstamo 
	FILA9BP			= ""
	'FILA 9  BQ:Número de la partida registral 
	FILA9BQ			= ""
	'FILA 9  BR:Código de indicador de primeravivienda
	FILA9BR			= ""
	'FILA 9  BS:Predio: Código de ubigeo 
	FILA9BS			= ""
	'FILA 9  BT:Predio: Dirección completa ydetallada
	FILA9BT			= ""
	'FILA 9  BU:Predio: Urbanización 
	FILA9BU			= ""
	'FILA 9  BV:Predio: Provincia 
	FILA9BV			= ""
	'FILA 9  BW:Predio: Distrito 
	FILA9BW			= ""
	'FILA 9  BX:Predio: Departamento 
	FILA9BX			= ""
	'FILA 9  BY:Origen código ubigeo 
	FILA9BY			= ""
	'FILA 9  BZ:Origen dirección 
	FILA9BZ			= ""
	'FILA 9  CA:Destino código ubigeo 
	FILA9CA			= ""
	'FILA 9  CB:Destino dirección 
	FILA9CB			= ""
	'FILA 9  CC:Pasajero: Apellidos y nombres 
	FILA9CC			= ""
	'FILA 9  CD:Pasajero: Numero documentosidentidad
	FILA9CD			= ""
	'FILA 9  CE:Pasajero: Tipo documentoidentidad
	FILA9CE			= ""
	'FILA 9  CF:Origen código ubigeo 
	FILA9CF			= ""
	'FILA 9  CG:Origen dirección 
	FILA9CG			= ""
	'FILA 9  CH:Destino código ubigeo 
	FILA9CH			= ""
	'FILA 9  CI:Destino dirección 
	FILA9CI			= ""
	'FILA 9  CJ:Servicio de transporte: Númerde asiento
	FILA9CJ			= ""
	'FILA 9  CK:Servicio de transporte: Fechaprogramada de inicio de viaje
	FILA9CK			= ""
	'FILA 9  CL:Servicio de transporte: Horaprogramada de inicio de viaje
	FILA9CL			= ""
	'FILA 9  CM:Decreto supremo deaprobación del contrato
	FILA9CM			= ""
	'FILA 9  CN:Área de contrato - Lote 
	FILA9CN			= ""
	'FILA 9  CO:Periodo de pago Fecha de inicio 
	FILA9CO			= ""
	'FILA 9  CP:Periodo de pago Fecha de fin 
	FILA9CP			= ""
	'FILA 9  CQ:Partida arancelaria 
	FILA9CQ			= ""
	'FILA 9  CR:Númerodeplaca
	FILA9CR			= ""
	'FILA 9  CS:Categoría 
	FILA9CS			= ""
	'FILA 9  CT:Marca 
	FILA9CT			= ""
	'FILA 9  CU:Modelo 
	FILA9CU			= ""
	'FILA 9  CV:Color 
	FILA9CV			= ""
	'FILA 9  CW:Motor 
	FILA9CW			= ""
	'FILA 9  CX:Combustible 
	FILA9CX			= ""
	'FILA 9  CY:Form. Rodante
	FILA9CY			= ""
	'FILA 9  CZ:VIN  
	FILA9CZ			= ""
	'FILA 9  DA:Serie / Chasis 
	FILA9DA			= ""
	'FILA 9  DB:Año de fabricación 
	FILA9DB			= ""
	'FILA 9  DC:Año modelo 
	FILA9DC			= ""
	'FILA 9  DD:Versión 
	FILA9DD			= ""
	'FILA 9  DE:Ejes 
	FILA9DE			= ""
	'FILA 9  DF:Asientos 
	FILA9DF			= ""
	'FILA 9  DG:Pasajeros 
	FILA9DG			= ""
	'FILA 9  DH:Ruedas 
	FILA9DH			= ""
	'FILA 9  DI:Carrocería 
	FILA9DI			= ""
	'FILA 9  DJ:Potencia 
	FILA9DJ			= ""
	'FILA 9  DK:Cilindros 
	FILA9DK			= ""
	'FILA 9  DL:Cilindrada 
	FILA9DL			= ""
	'FILA 9  DM:Peso bruto 
	FILA9DM			= ""
	'FILA 9  DN:Peso Neto 
	FILA9DN			= ""
	'FILA 9  DO:Carga útil 
	FILA9DO			= ""
	'FILA 9  DP:Longitud 
	FILA9DP			= ""
	'FILA 9  DQ:Altura 
	FILA9DQ			= ""
	'FILA 9  DR:Ancho 
	FILA9DR			= ""
	'FILA 9  DS:Número de asiento 
	FILA9DS			= ""
	'FILA 9  DT:Información del manifiesto depasajeros
	FILA9DT			= ""
	'FILA 9  DU:Número de documento delpasajero
	FILA9DU			= ""
	'FILA 9  DV:Tipo de documento delpasajero
	FILA9DV			= ""
	'FILA 9  DW:Nombres y apellidos delpasajero
	FILA9DW			= ""
	'FILA 9  DX:Código de ubigeo origen delpasajero
	FILA9DX			= ""
	'FILA 9  DY:Dirección de origen delpasajero
	FILA9DY			= ""
	'FILA 9  DZ:Código de ubigeo destino delpasajero
	FILA9DZ			= ""
	'FILA 9  EA:Dirección del destino delpasajero
	FILA9EA			= ""
	'FILA 9  EB:Fecha inicio programado 
	FILA9EB			= ""
	'FILA 9  EC:Hora de inicio programado 
	FILA9EC			= ""
	'FILA 9  ED:Detracción: Matrícula de laembarcación pesquera
	FILA9ED			= ""
	'FILA 9  EE:Detracción: Nombre de laembarcación pesquera
	FILA9EE			= ""
	'FILA 9  EF:Detracción: Descripción del tipode la especie vendida
	FILA9EF			= ""
	'FILA 9  EG:Detracción: Lugar de descarga 
	FILA9EG			= ""
	'FILA 9  EH:Detracción: Cantidad de laespecie vendida
	FILA9EH			= ""
	'FILA 9  EI:Detracción: Fecha de descarga 
	FILA9EI			= ""
	'FILA 9  EJ:Detracción: Código ubigeoorigen
	FILA9EJ			= ""
	'FILA 9  EK:Detracción: Dirección de origen 
	FILA9EK			= ""
	'FILA 9  EL:Detracción: Código ubigeodestino
	FILA9EL			= ""
	'FILA 9  EM:Detracción: Dirección dedestino
	FILA9EM			= ""
	'FILA 9  EN:Detracción: Detalle del viaje 
	FILA9EN			= ""
	'FILA 9  EO:Detracción: Valor referencia delservicio de transporte
	FILA9EO			= ""
	'FILA 9  EP:Detracción: Valor referencialsobre la carga afectiva
	FILA9EP			= ""
	'FILA 9  EQ:Detracción: Valor referencialsobre la carga útil nominal
	FILA9EQ			= ""
	'FILA 9  ER:Código de producto GS1 
	FILA9ER			= ""
	'FILA 9  ES:Tipo de estructura GTIN delcódigo de producto GS1
	FILA9ES			= ""
	'FILA 9  ET:Porcentaje del ISC
	FILA9ET			= ""
	'FILA 9  Eu:CANTIDAD BOLSAS'
	FILA9EU			= ""
	'FILA 9  Ev:MONTO UNIT BOLSA'
	FILA9EV			= ""
	'FILA 9  Ew:MONTO TOTAL ICBPER'
	FILA9EW			= ""


	


	






	'FILA ULTIMA (Datos de la lina)
	fultima = "FF00FF"

	dim fs,f
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	'set f=fs.CreateTextFile("d:\VENTAS_NEW\efact\"&nombre&".csv",true)
	'set f=fs.CreateTextFile("d:\efact\daemon\documents\"&nombre&".csv",true)
	set f=fs.CreateTextFile("d:\EFACT_MODULO\daemon\documents\in\invoice\"&nombre&".csv",true)
	
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
	f.WriteLine(utf8_simbom(fila7))
	'LINEA 8'
	f.WriteLine(utf8_simbom(fila8))
	'LINEA DETALLE'


	rsDet.movefirst
	FOR I = 0 TO rsDet.recordcount-1 

		'RESETEO DE DATOS PARA LA NUEVA LINEA'
		fDetalleNroOrden 							= ""
		fDetalleUnidMed								= ""
		fDetalleCantid								= ""
		fDetalleDescrip								= ""
		fDetallePreVentUnit							= ""
		fDetalleMontoBaseIgv						= ""
		fDetalleMontoTotalIgv						= ""
		fDetalleTipoAfectIgv						= ""			
		fDetalleCodigoTibutos						= ""
		fDetalleProcentIgv							= ""
		fDetalleCodProd 							= ""
		fDetalleValUnit								= ""
		fDetalleValVent								= ""
		fDetalleMontoBaseDescuentoAB 				= ""
		fDetalleFactorDescuentoAB					= ""
		fDetalleMontoTotalDescuentoAB				= ""
		fDetalleMontTotalImp						= ""
		fDetalleTotaldelaLinea						= ""
		FILA9EU			= ""
		FILA9EV			= ""
		FILA9EW			= ""
		FILA9AP			= ""
		'FIN DE RESETEO DE DATOS PARA NUEVA LINEA'







		fDetalleNroOrden 							= cint(rsDet("item"))
		fDetalleUnidMed								= "C62"		'APLICANDO CODIGO INTERNACIONAL'	
		fDetalleCantid								= cdbl(rsDet("SALE"))
		fDetalleDescrip								= ucase(replace(replace(trim(rsDet("descri")),",",""),"/"," "))


		fDetallePreVentUnit							= formatnumber( ( CDBL(rsDet("PRECIO"))   ) / CDBL(rsDet("sale")),2,,,false )
		'fDetallePreVentUnit							= formatnumber( ( CDBL(rsDet("PRECIO"))   ) / CDBL(rsDet("sale")),2,,,false )

		'XXX'
		'fDetalleMontoBaseIgv						= formatnumber(CDBL(rsDet("PRECIO")) - CDBL(rsDet("igv")) + CDBL(rsDet("descuento")),2,,,false)
		fDetalleMontoBaseIgv						= formatnumber(CDBL(rsDet("PRECIO")) - CDBL(rsDet("igv")),2,,,false)
		
		fDetalleMontoTotalIgv						= formatnumber(CDBL(rsDet("igv")),2,,,false) 
		fDetalleTipoAfectIgv						= "10"			
		fDetalleCodigoTibutos						= "1000"
		fDetalleProcentIgv							= formatnumber(cdbl(rs("porigv")),0,,,false)
		fDetalleCodProd 							= ucase(trim(rsDet("codart")))
		fDetalleValUnit								= formatnumber( ( CDBL(rsDet("PRECIO"))  + CDBL(rsDet("descuento")) - CDBL(rsDet("igv"))  ) / CDBL(rsDet("sale")),2,,,false )
		fDetalleValVent								= formatnumber(CDBL(fDetalleValUnit)*cdbl(fDetalleCantid) - cdbl(rsDet("descuento")),2,,,false)
		

		'SOLO APARECERA SI HAY DESCUENTO
		if cdbl(rsDet("descuento")) > 0 then
			'fDetallePreVentUnit							= formatnumber( ( CDBL(rsDet("PRECIO"))   ) / CDBL(rsDet("sale")),2,,,false )
			'fDetalleValVent								= formatnumber((CDBL(fDetalleValUnit)*cdbl(fDetalleCantid) ) - cdbl(rsDet("descuento")) ,2,,,false)
			fDetalleMontoBaseDescuentoAB 				= formatnumber(cdbl(rsDet("descuento"))  - CDBL(rsDet("igv"))  + cdbl(rsDet("PRECIO")  ),2,,,false)
			fDetalleFactorDescuentoAB					= formatnumber(cdbl(rsDet("pordes"))/100,2,,,false)
			fDetalleMontoTotalDescuentoAB				= formatnumber(cdbl(rsDet("descuento")),2,,,false)
			FILA9AP										= formatnumber(cdbl(rsDet("pordes")),2,,,false)
		end if
		
		



		fDetalleMontTotalImp						= formatnumber(CDBL(rsDet("igv")),2,,,false)
		fDetalleTotaldelaLinea						= fDetalleMontoBaseIgv

		if ucase(left(trim(rsDet("codart")),3)) = "BOL" then
			fDetalleMontTotalImp = formatnumber(CDBL(rsDet("igv")),2,,,false) + cdbl(fDetalleMontoTotalIgv)
			fDetalleTotaldelaLinea = cdbl(fDetalleMontTotalImp) + formatnumber(CDBL(rsDet("precio")),2,,,false)
			FILA9EU			= fDetalleCantid

            'POR MOTIVOS DE 99% DE DESCUENTO AL PRECIO DE LA BOLSA COLOCARE 0.01 YA QUE NO SE PUEDE PONER PRECIO 0 EN LA COLUMNA DE PRECI0 E IGV
            '                                                                           ======================================================== 
            if cdbl(fDetalleMontoBaseIgv) = 0 then
                fDetalleMontoBaseIgv = "0.01"
            end if
            if cdbl(fDetallePreVentUnit) = 0 then
                fDetallePreVentUnit = "0.01"
            end if
            
            if cdbl(fDetalleMontoTotalIgv) = 0 then
                fDetalleMontoTotalIgv = "0.01"
            end if

            if cdbl(fDetalleValVent) = 0 then
                fDetalleValVent = "0.01"
            end if

            if cdbl(fDetalleMontTotalImp) = 0 then
                fDetalleMontTotalImp = "0.01"
            end if
            if cdbl(fDetalleTotaldelaLinea) = 0 then
                fDetalleTotaldelaLinea = "0.01"
            end if

			VALORICBPER = "0.10"

			set rsICB = RsNuevo
			if rsICB.state > 0 then 
				rsICB.close
			end if

			rsICB.open "select isc from PARAMETROS",cnn
			if rsICB.recordcount > 0 then
				rsICB.movefirst
				VALORICBPER = rsICB("isc")
			end if

			FILA9EV			= VALORICBPER
			FILA9EW			= CDBL(VALORICBPER) * CDBL(fDetalleCantid)


			'FILA9EV			= "0.1"
			'FILA9EW			= CDBL("0.1") * CDBL(fDetalleCantid)
		end if



		filaDD = fDetalleNroOrden										
		filaDD = filaDD &","&fDetalleUnidMed											
		filaDD = filaDD &","&fDetalleCantid											
		filaDD = filaDD &","&fDetalleDescrip											
		filaDD = filaDD &","&fDetallePreVentUnit										
		filaDD = filaDD &","&fDetalleCodPreVentUnit 									
		filaDD = filaDD &","&fDetalleValorRefUnit									
		filaDD = filaDD &","&fDetalleCodValRefUnit									
		filaDD = filaDD &","&fDetalleMontoBaseIgv									
		filaDD = filaDD &","&fDetalleMontoTotalIgv									
		filaDD = filaDD &","&fDetalleTipoAfectIgv									
		filaDD = filaDD &","&fDetalleCodigoTibutos									
		filaDD = filaDD &","&fDetalleProcentIgv										
		filaDD = filaDD &","&fDetalleMontBaseISC										
		filaDD = filaDD &","&fDetalleMontoTotalIsc									
		filaDD = filaDD &","&fDetalleCodigoTipoSistCalcIsc							
		filaDD = filaDD &","&fDetalleCodTributoISC									
		filaDD = filaDD &","&fDetalleCodProdSunat									
		filaDD = filaDD &","&fDetalleCodProd 											
		filaDD = filaDD &","&fDetalleValUnit											
		filaDD = filaDD &","&fDetalleValVent											
		filaDD = filaDD &","&fDetalleMontoBaseOtros									
		filaDD = filaDD &","&fDetalleProcentOtros									
		filaDD = filaDD &","&fDetalleMontoTotalOtros		
		'y'							
		filaDD = filaDD &","&fDetalleMontoBaseDescuentoAB							
		filaDD = filaDD &","&fDetalleFactorDescuentoAB								
		filaDD = filaDD &","&fDetalleMontoTotalDescuentoAB							
		filaDD = filaDD &","&fDetalleMontoBaseDescuentoNOAB							
		filaDD = filaDD &","&fDetalleFactorDescuentoNOAB								
		filaDD = filaDD &","&fDetalleMontoTotalDescuentoNOAB						
		filaDD = filaDD &","&fDetalleMontBaseCargoAB									
		filaDD = filaDD &","&fDetalleFactorCargoAB									
		filaDD = filaDD &","&fDetalleMontTotalCargoAB								
		filaDD = filaDD &","&fDetalleMontBaseCargoNOAB								
		filaDD = filaDD &","&fDetalleFactorCargoNoAB 								
		filaDD = filaDD &","&fDetalleMontTotalCargoNOAB 								
		filaDD = filaDD &","&fDetalleMontTotalImp 									
		filaDD = filaDD &","&fDetalleTotaldelaLinea 									
		filaDD = filaDD &","&FILA9AM
		filaDD = filaDD &","&FILA9AN
		filaDD = filaDD &","&FILA9AO
		filaDD = filaDD &","&FILA9AP
		filaDD = filaDD &","&FILA9AQ
		filaDD = filaDD &","&FILA9AR
		filaDD = filaDD &","&FILA9AS
		filaDD = filaDD &","&FILA9AT
		filaDD = filaDD &","&FILA9AU
		filaDD = filaDD &","&FILA9AV
		filaDD = filaDD &","&FILA9AW
		filaDD = filaDD &","&FILA9AX
		filaDD = filaDD &","&FILA9AY
		filaDD = filaDD &","&FILA9AZ
		filaDD = filaDD &","&FILA9BA
		filaDD = filaDD &","&FILA9BB
		filaDD = filaDD &","&FILA9BC
		filaDD = filaDD &","&FILA9BD
		filaDD = filaDD &","&FILA9BE
		filaDD = filaDD &","&FILA9BF
		filaDD = filaDD &","&FILA9BG
		filaDD = filaDD &","&FILA9BH
		filaDD = filaDD &","&FILA9BI
		filaDD = filaDD &","&FILA9BJ
		filaDD = filaDD &","&FILA9BK
		filaDD = filaDD &","&FILA9BL
		filaDD = filaDD &","&FILA9BM
		filaDD = filaDD &","&FILA9BN
		filaDD = filaDD &","&FILA9BO
		filaDD = filaDD &","&FILA9BP
		filaDD = filaDD &","&FILA9BQ
		filaDD = filaDD &","&FILA9BR
		filaDD = filaDD &","&FILA9BS
		filaDD = filaDD &","&FILA9BT
		filaDD = filaDD &","&FILA9BU
		filaDD = filaDD &","&FILA9BV
		filaDD = filaDD &","&FILA9BW
		filaDD = filaDD &","&FILA9BX
		filaDD = filaDD &","&FILA9BY
		filaDD = filaDD &","&FILA9BZ
		filaDD = filaDD &","&FILA9CA
		filaDD = filaDD &","&FILA9CB
		filaDD = filaDD &","&FILA9CC
		filaDD = filaDD &","&FILA9CD
		filaDD = filaDD &","&FILA9CE
		filaDD = filaDD &","&FILA9CF
		filaDD = filaDD &","&FILA9CG
		filaDD = filaDD &","&FILA9CH
		filaDD = filaDD &","&FILA9CI
		filaDD = filaDD &","&FILA9CJ
		filaDD = filaDD &","&FILA9CK
		filaDD = filaDD &","&FILA9CL
		filaDD = filaDD &","&FILA9CM
		filaDD = filaDD &","&FILA9CN
		filaDD = filaDD &","&FILA9CO
		filaDD = filaDD &","&FILA9CP
		filaDD = filaDD &","&FILA9CQ
		filaDD = filaDD &","&FILA9CR
		filaDD = filaDD &","&FILA9CS
		filaDD = filaDD &","&FILA9CT
		filaDD = filaDD &","&FILA9CU
		filaDD = filaDD &","&FILA9CV
		filaDD = filaDD &","&FILA9CW
		filaDD = filaDD &","&FILA9CX
		filaDD = filaDD &","&FILA9CY
		filaDD = filaDD &","&FILA9CZ
		filaDD = filaDD &","&FILA9DA
		filaDD = filaDD &","&FILA9DB
		filaDD = filaDD &","&FILA9DC
		filaDD = filaDD &","&FILA9DD
		filaDD = filaDD &","&FILA9DE
		filaDD = filaDD &","&FILA9DF
		filaDD = filaDD &","&FILA9DG
		filaDD = filaDD &","&FILA9DH
		filaDD = filaDD &","&FILA9DI
		filaDD = filaDD &","&FILA9DJ
		filaDD = filaDD &","&FILA9DK
		filaDD = filaDD &","&FILA9DL
		filaDD = filaDD &","&FILA9DM
		filaDD = filaDD &","&FILA9DN
		filaDD = filaDD &","&FILA9DO
		filaDD = filaDD &","&FILA9DP
		filaDD = filaDD &","&FILA9DQ
		filaDD = filaDD &","&FILA9DR
		filaDD = filaDD &","&FILA9DS
		filaDD = filaDD &","&FILA9DT
		filaDD = filaDD &","&FILA9DU
		filaDD = filaDD &","&FILA9DV
		filaDD = filaDD &","&FILA9DW
		filaDD = filaDD &","&FILA9DX
		filaDD = filaDD &","&FILA9DY
		filaDD = filaDD &","&FILA9DZ
		filaDD = filaDD &","&FILA9EA
		filaDD = filaDD &","&FILA9EB
		filaDD = filaDD &","&FILA9EC
		filaDD = filaDD &","&FILA9ED
		filaDD = filaDD &","&FILA9EE
		filaDD = filaDD &","&FILA9EF
		filaDD = filaDD &","&FILA9EG
		filaDD = filaDD &","&FILA9EH
		filaDD = filaDD &","&FILA9EI
		filaDD = filaDD &","&FILA9EJ
		filaDD = filaDD &","&FILA9EK
		filaDD = filaDD &","&FILA9EL
		filaDD = filaDD &","&FILA9EM
		filaDD = filaDD &","&FILA9EN
		filaDD = filaDD &","&FILA9EO
		filaDD = filaDD &","&FILA9EP
		filaDD = filaDD &","&FILA9EQ
		filaDD = filaDD &","&FILA9ER
		filaDD = filaDD &","&FILA9ES
		filaDD = filaDD &","&FILA9ET
		filaDD = filaDD &","&FILA9Eu
		filaDD = filaDD &","&FILA9Ev
		filaDD = filaDD &","&FILA9Ew

		f.WriteLine(utf8_simbom(filaDD))
		response.write("<br>"+filaDD+"<br>")
		rsDet.movenext
	next






	'LINEA ULTIMA'
	f.WriteLine(fultima)

	f.close
	set f=nothing
	set fs=nothing

	Response.Redirect "/apijf/public/index.php/auth?tipdoc=invoice&ope="&ope&"&file="&nombre&".xml"


else 
	response.write("No existe tal documento")
end if


%>  