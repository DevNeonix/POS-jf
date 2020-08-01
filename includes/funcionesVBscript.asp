<%
Function RsNuevo
        Dim Connection, Recordset
        Set Recordset = Server.CreateObject("ADODB.Recordset")
        recordset.ActiveConnection = cnn
        Recordset.CursorType       = 3 'CONST adOpenStatic = 3
	    Recordset.LockType         = 1 'CONST adReadOnly = 1
	    Recordset.CursorLocation   = 3 'CONST adUseClient = 3
        Set RsNuevo = Recordset
End Function
Function fnFileSize(cArchivo)
	On Error Resume Next
	Dim fso, f, cAppPhyPath
	cAppPhyPath=Request.ServerVariables("APPL_PHYSICAL_PATH")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFile(cAppPhyPath&cArchivo)
	If Err.number = 0 Then
		fnFileSize = f.Size
	Else
		fnFileSize = 0
	End If
	Response.Write(err.number&" - "&f.Size)
End Function


Function decimaldocena(num)
	num =CDBL(num)
	nume = Int(num)
	VALOR = NUM - NUME
	VALOR= VALOR *12
	VALOR = CINT(abs(VALOR))
	CAD = TRIM(CSTR(NUME))
	decimaldocena= (CAD+ " " + LEFT(CSTR(VALOR),2))
end Function

function prepara_str_sql(str)
	IF NOT ISNULL(str) THEN
	str = trim(str)
	str = replace(str,"'","''")
	str=replace(str,"--","")
	END IF
	prepara_str_sql = str
end function

function muestra_cadena(str)
	IF  NOT ISNULL(str) THEN
	str = trim(str)
	str = replace(str,"""","&quot;")
	str = replace(str,"<","&lt;")
	str = replace(str,">","&gt;")
    str = replace(str,"á","a")
    str = replace(str,"é","e")
    str = replace(str,"í","i")
    str = replace(str,"ó","o")
    str = replace(str,"ú","u")
    str = replace(str,"ñ","n")
	muestra_cadena = str
	END IF
end function

function utf8_simbom(cadena)
	if not isNull(cadena) then
	    ' Eliminamos los espacios a ambos lados de la cadena
	    strCadena = Trim(lCase(cadena))
	    ' Reemplazamos carácteres especiales
	    'strCadena = replace(replace(strCadena,"'",""),"""","")
	    'strCadena = replace(replace(strCadena,"&quot;",""),vbcrlf,"")
	    'strCadena = replace(replace(strCadena,"<br>","")," ","-")
	    set expReg = New RegExp
	    ' Todas las ocurrencias
	    expReg.Global = True
	    expReg.Pattern = "[àáâãäå]"
	    strCadena = expReg.Replace(strCadena, "a")
	    expReg.Pattern = "[èéêë]"
	    strCadena = expReg.Replace(strCadena, "e")
	    expReg.Pattern = "[ìíîï]"
	    strCadena = expReg.Replace(strCadena, "i")
	    expReg.Pattern = "[òóôõö]"
	    strCadena = expReg.Replace(strCadena, "o")
	    expReg.Pattern = "[ùúûü]"
	    strCadena = expReg.Replace(strCadena, "u")
	    expReg.Pattern = "[ñ]"
	    strCadena = expReg.Replace(strCadena, "n")
	    expReg.Pattern = "[ç]"
	    strCadena = expReg.Replace(strCadena, "c")
	    ' Todo lo que no cumpla este patron
	    expReg.Pattern = "[^a-zA-Z0-9-,./% ]"
	    strCadena = expReg.Replace(strCadena, "")
	    set expReg = nothing
	    utf8_simbom = ucase(strCadena)
	  else
	    utf8_simbom = ""
	  end if
end function

function reemplazaNull(str,nval)
	IF  ISNULL(str) THEN
		str = nval
	END IF
	reemplazaNull = str
end function

Function AlphaNumericOnly(strSource) 
if isnull(strSource) then
	AlphaNumericOnly = "" 
else
	Dim i 
	Dim strResult 
	For i = 1 To Len(strSource) 
	    Select Case Asc(Mid(strSource, i, 1)) 
	        Case 32,37,46,47, 48,49,50,51,52,53,54,55,56,57,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89, 90, 97,98,99,100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,121, 122
	            strResult = strResult & Mid(strSource, i, 1) 
	        End Select 
	Next 
	    AlphaNumericOnly = strResult 
end if
End Function


%>