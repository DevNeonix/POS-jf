<%


Dim Unidades(20), Decenas(20), Oncenas(20)
Dim Veintes(20), Centenas(20)
'CantidadenTexto = Numlet(CCur(133)) ' ***** con esto se convierte
'Response.write CantidadenTexto
Function Numlet(NUM)
	Dim DEC, MILM, MILL, MILE, UNID
ReDim SALI(11)
Dim var, I, AUX
'NUM# = Round(NUM#, 2)
var = Trim(NUM)
If InStr(var, ".") = 0 Then
	var = var + ".00"
End If
If InStr(var, ".") = Len(var) - 1 Then
	var = var + "0"
End If
var = String(15 - Len(LTrim(var)), "0") + LTrim(var)
DEC = Mid(var, 14, 2)
MILM = Mid(var, 1, 3)
MILL = Mid(var, 4, 3)
MILE = Mid(var, 7, 3)
UNID = Mid(var, 10, 3)
For I = 1 To 11: SALI(I) = " ": Next 
	I = 0
	Unidades(1) = "UNO "
	Unidades(2) = "DOS "
	Unidades(3) = "TRES "
	Unidades(4) = "CUATRO "
	Unidades(5) = "CINCO "
	Unidades(6) = "SEIS "
	Unidades(7) = "SIETE "
	Unidades(8) = "OCHO "
	Unidades(9) = "NUEVE "
	
	Decenas(1) = "DIEZ "
	Decenas(2) = "VEINTE "
	Decenas(3) = "TREINTA "
	Decenas(4) = "CUARENTA "
	Decenas(5) = "CINCUENTA "
	Decenas(6) = "SESENTA "
	Decenas(7) = "SETENTA "
	Decenas(8) = "OCHENTA "
	Decenas(9) = "NOVENTA "
	
	Oncenas(1) = "ONCE "
	Oncenas(2) = "DOCE "
	Oncenas(3) = "TRECE "
	Oncenas(4) = "CATORCE "
	Oncenas(5) = "QUINCE "
	Oncenas(6) = "DIECISEIS "
	Oncenas(7) = "DIECISIETE "
	Oncenas(8) = "DIECIOCHO "
	Oncenas(9) = "DIECINUEVE "
	
	Veintes(1) = "VEINTIUNO "
	Veintes(2) = "VEINTIDOS "
	Veintes(3) = "VEINTITRES "
	Veintes(4) = "VEINTICUATRO "
	Veintes(5) = "VEINTICINCO "
	Veintes(6) = "VEINTISEIS "
	Veintes(7) = "VEINTISIETE "
	Veintes(8) = "VEINTIOCHO "
	Veintes(9) = "VEINTINUEVE "
	
	Centenas(1) = " CIENTO "
	Centenas(2) = " DOSCIENTOS "
	Centenas(3) = " TRESCIENTOS "
	Centenas(4) = "CUATROCIENTOS "
	Centenas(5) = " QUINIENTOS "
	Centenas(6) = " SEISCIENTOS "
	Centenas(7) = " SETECIENTOS "
	Centenas(8) = " OCHOCIENTOS "
	Centenas(9) = " NOVECIENTOS "
	
	If NUM > 999999999999.99 Then Numlet = " ": Exit Function
	If MILM >= 1 Then
		SALI(2) = " MIL ": '** MILES DE MILLONES
		SALI(4) = " MILLONES "
		If MILM <> 1 Then
			Unidades(1) = "UN "
			Veintes(1) = "VEINTIUN "
			SALI(1) = Descifrar(Val(MILM))
		End If
	End If
	If MILL >= 1 Then
		If MILL < 2 Then
			SALI(3) = "UN ": '*** UN MILLON
			If Trim(SALI(4)) <> "MILLONES" Then
				SALI(4) = " MILLON "
			End If
		Else
			SALI(4) = " MILLONES ": '*** VARIOS MILLONES
			Unidades(1) = "UN "
			Veintes(1) = "VEINTIUN "
			SALI(3) = Descifrar(MILL)
		End If
	End If
	For I = 2 To 9
		
		
		Centenas(I) = replace(Mid(Centenas(I), 1, 11) + "OS","OO","O")
		Centenas(I) = replace(Centenas(I),"OSOS","OS")
	Next 
	If MILE > 0 Then
		SALI(6) = " MIL ": '*** MILES
		If MILE <> 1 Then
			SALI(5) = Descifrar(MILE)
		End If
	End If
	Unidades(1) = "UN "
	Veintes(1) = "VEINTIUNO"
	'If UNID >= 1 Then
		SALI(7) = Descifrar(UNID): '*** CIENTOS
	''	If DEC >= 10 Then
			SALI(8) = " CON ": '*** DECIMALES
			SALI(10) = (DEC)
	''	End If
	'End If
	If MILM = 0 And MILL = 0 And MILE = 0 And UNID = 0 Then SALI(7) = " CERO "
	AUX = ""
	For I = 1 To 11
		AUX = AUX + SALI(I)
	Next 
	Numlet = Trim(AUX)&"/100 "

End Function

Function Descifrar(numero)
	Dim SAL(4)
	Dim I, CT , DC , DU , UD 
	Dim VARIABLE

	For I = 1 To 4: SAL(I) = " ": Next 
	VARIABLE = String(3 - Len(Trim(numero)), "0") + Trim(numero)
	CT = Mid(VARIABLE, 1, 1): '*** CENTENA
	DC = Mid(VARIABLE, 2, 1): '*** DECENA
	DU = Mid(VARIABLE, 2, 2): '*** DECENA + UNIDAD
	UD = Mid(VARIABLE, 3, 1): '*** UNIDAD
	If numero = 100 Then
	SAL(1) = "CIEN "
	Else
	If CT <> 0 Then SAL(1) = Centenas(CT)
	If DC <> 0 Then
	If DU <> 10 And DU <> 20 Then
	If DC = 1 Then
	SAL(2) = Oncenas(UD)
	Descifrar = Trim(SAL(1) + " " + SAL(2))
	Exit Function
	End If
	If DC = 2 Then
	SAL(2) = Veintes(UD)
	Descifrar = Trim(SAL(1) + " " + SAL(2))
	Exit Function
	End If
	End If
	SAL(2) = " " + Decenas(DC)
	If UD <> 0 Then SAL(3) = "Y "
	End If
	If UD <> 0 Then SAL(4) = Unidades(UD)
	End If
	Descifrar = Trim(SAL(1) + SAL(2) + SAL(3) + SAL(4))
End Function



function sleep(scs)
    Dim lo_wsh, ls_cmd
    Set lo_wsh = CreateObject( "WScript.Shell" )
    ls_cmd = "%COMSPEC% /c ping -n " & 1 + scs & " 127.0.0.1>nul"
    lo_wsh.Run ls_cmd, 0, True 
End Function

%>