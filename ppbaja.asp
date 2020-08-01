<%@ Language=VBScript %>
<!--#include file="includes/Cnn.inc"-->
<!--#include file="includes/funcionesVBscript.asp"-->
<!--#include file="funcionesefact.asp"-->
<%  

Response.CharSet = "UTF-8"
ope = request.querystring("ope")


cadValidaBaja = "select * from JACINTA..efact_bajas where convert(int,estado) between 200 and 300 and operacion = '"&ope&"'"
if rs.state > 0 then 
	rs.close
end if
rs.open cadValidaBaja,cnn
if rs.recordcount>0 then
	response.write("Este documento ya ha sido dajo de baja, valida en el Portal de Efact por favor.")
	response.end
end if



cad = "select *,getdate() as fecact from JACINTA..movimcab where OPERACION='"&ope&"' and GETDATE() < dateadd(day,30,FECDOC)"
'response.write(cad)
if rs.state > 0 then 
	rs.close
end if

rs.open cad,cnn

if rs.recordcount > 0 then 
	rs.movefirst

	'VALIDANDO SI ES UNA BOLETA'
	

	cadDet = "select d.*,l.AR_CDESCRI as descri from JACINTA..movimdet d full outer join rsfaccar..AL0012ARTI l on d.CODART = l.AR_CCODIGO where OPERACION='"&ope&"'"

	set rsDet = RsNuevo
	rsDet.open cadDet,cnn

	cadCli = "select * from JACINTA..CLIENTES where CLIENTE='"&RS("CLIENTE")&"' and estado ='A'"
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
	coddoc = "03"

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



	cadIdentificadorDoc = "select * from mostrar_cod_nuevo_bajas"
	set rsCadIdeDoc = RsNuevo
	rsCadIdeDoc.open cadIdentificadorDoc,cnn

	identificacion = rsCadIdeDoc("codigo")


	nombre = miRuc&"-"&identificacion



	pvp 	= FORMATNUMBER(cdbl(RS("pvp")),2,,,false)
	SUBTOTAL 	= FORMATNUMBER(cdbl(RS("SUBTOT")),2,,,false)
	IGV 		= FORMATNUMBER(cdbl(RS("IGV")),2,,,false)

    ' no estaba considerando si no tenia ISC o estaba ebn null --> mabel 23-12-2019
    if isnull(rs("isc"))  then 
        icbper = 0
    else
	    ICBPER 		= FORMATNUMBER((RS("isc")),2,,,false)
    end if
	Total 		= FORMATNUMBER(cdbl(RS("TOTAL")),2,,,false)
	MON 		= "PEN"
	TIPOPERACION = "0101"
	cntLinDet = rsDet.recordcount
	cntLinDet = 1
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
	elseif len(clienteDOC) = 11 then
		clienteTipDoc		= "6"
		
	else
		clienteTipDoc		= "0"
	end if



	'***********CLIENTE VARIOS*******'

	if CDBL(clienteDOC) = 0 then
		clienteTipDoc		= "0"
	end if

	if clienteTipDoc = "0" and cdbl(total)>699 then
		response.write("No se puede emitir un documento de mas de 700 soles a un cliente 'VARIOS'")
		response.end
	end if

	clienteDireccion	= AlphaNumericOnly(trim(rsCLI("direccion")))
	clienteRazon		= AlphaNumericOnly(trim(rsCLI("nombre")))
	clienteMail			= AlphaNumericOnly(trim(rsCLI("mail")))


	'FILA1
	'=============================================================================================================='
	'=============================================================================================================='
	'=============================================================================================================='

	


	'FILA 1  A :FECHA DE EMISION'
	F1A  	= CStr(year(RS("fecact"))&"-"&right("00"&month(RS("fecact")),2)&"-"&right("00"&day(RS("fecact")),2))
	'FILA 1  B :FECHA DE EMISION COMPROBANTE RELACIONADO'
	F1B		= CStr(year(RS("FECDOC"))&"-"&right("00"&month(RS("FECDOC")),2)&"-"&right("00"&day(RS("FECDOC")),2))
	'FILA 1  C :IDENTIFICACION'
	F1C		= identificacion
	'FILA 1  D :CANT ITEMS'
	F1D	 	= cntLinDet

	fila1 = F1A
	fila1 = fila1&","&F1B
	fila1 = fila1&","&F1C
	fila1 = fila1&","&F1D

	response.write(FILA1)
	
	'FILA2
	'=============================================================================================================='
	'=============================================================================================================='
	'=============================================================================================================='


	'FILA 5 A:Apellidos y nombres,denominación o razónsocial
	F2RazonSocial 							=miRS
	'FILA 5 B:Nombre comercial 
	F2NomComercial 							=miNombreComercial
	'FILA 5 C:Número de RUC 
	F2NumRuc	 							=miRuc
	'FILA 5 D:Código Ubigeo 
	F2CodUbigeo 							=codUBIGUEO
	'FILA 5 E:Dirección 
	F2Direccion 							=miDireccion
	'FILA 5 F:Urbanización 
	F2Urbanizacion 							=""
	'FILA 5 G:Departamento 
	F2Departamento 							="LIMA"
	'FILA 5 H:Provincia 
	F2Provincia 							="LIMA"
	'FILA 5 I:Distrito 
	F2Distrito 								="ATE"
	'FILA 5 J:Codigo del pais 
	F2CodPais	 							="PE"
	
	FILA2 = F2RazonSocial
	FILA2 = FILA2 &","&F2NomComercial
	FILA2 = FILA2 &","&F2NumRuc	 				
	FILA2 = FILA2 &","&F2CodUbigeo 		
	FILA2 = FILA2 &","&F2Direccion 		
	FILA2 = FILA2 &","&F2Urbanizacion
	FILA2 = FILA2 &","&F2Departamento
	FILA2 = FILA2 &","&F2Provincia 		
	FILA2 = FILA2 &","&F2Distrito 			
	FILA2 = FILA2 &","&F2CodPais	 			

	response.write("<br/>")
	response.write(FILA2)


	'FILA3
	'=============================================================================================================='
	'=============================================================================================================='
	'=============================================================================================================='

	
	'FILA 3 A: NUMRO DE ORDEN 
	f3A							="1"
	'FILA 3 B: CODIGO DEL TIPO DE COMPROBANTE 

	CODTIPCOM = ""
	serie = ""
	IF UCASE(TRIM(RS("CODDOC"))) = "FC" THEN
		CODTIPCOM= "01"
		serie = ""&trim(RS("SERIE"))
	ELSEIF UCASE(TRIM(RS("CODDOC"))) = "BL" THEN
		CODTIPCOM= "03"
		serie = ""&trim(RS("SERIE"))
	ELSEIF UCASE(TRIM(RS("CODDOC"))) = "NC" OR UCASE(TRIM(RS("CODDOC"))) = "ND"  THEN
		CODTIPCOM= "07"
		serie = trim(RS("SERIE"))
	END IF

	f3B							=CODTIPCOM
	'FILA 3 C: SERIE
	f3C							=SERIE
	'FILA 3 D: CORRELATIVO
	f3D							=trim(RS("NUMDOC"))
	'FILA 3 E: MOTIVO DE BAJA
	f3E							="Por cambio de artículo"

	fila3 = f3A
	fila3 = fila3 & ","&f3B
	fila3 = fila3 & ","&f3C
	fila3 = fila3 & ","&f3D 
	fila3 = fila3 & ","&f3E

	response.write("<br/>")
	response.write(fila3)
	fultima = "FF00FF"
	response.write("<br/>")
	response.write(fultima)


	dim fs,f
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	'set f=fs.CreateTextFile("d:\VENTAS_NEW\efact\"&nombre&".csv",true)
	'set f=fs.CreateTextFile("d:\efact\daemon\documents\"&nombre&".csv",true)
	set f=fs.CreateTextFile("d:\EFACT_MODULO\daemon\documents\in\"&nombre&".csv",true)
	
	'LINEA 1'
	f.WriteLine(utf8_simbom(fila1))
	'LINEA 2'
	f.WriteLine(utf8_simbom(fila2))
	'LINEA 3'
	f.WriteLine(utf8_simbom(fila3))
	'LINEA 4'
	f.WriteLine(utf8_simbom(fultima))

	f.close
	set f=nothing
	set fs=nothing

	Response.Redirect "/apijf/public/index.php/baja?ope="&ope&"&file="&nombre&".csv"


else 
	response.write("No existe tal documento")
end if


%>  