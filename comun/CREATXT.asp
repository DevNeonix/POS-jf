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
<%
cpos = TRIM(request.form("POS"))	' numero de documento origen
strFileName = "../Guias/"+CPOS+".txt"
%>
<body>
<center>
<input type="button" onclick="guardar('<%=strFileName %>');" style="cursor:pointer;" value="guardar este documento">
<input type="button" onclick="window.location.replace('../exportaguia.asp');" style="cursor:pointer;" value="regresar">
</center>


<%
cpos = TRIM(request.form("POS"))	' numero de documento origen
Const ForReading = 1, ForWriting = 2, ForAppending = 8
set FSO = server.createObject ( "Scripting.FileSystemObject" ) 
ARCHIVO = "D:\VENTAS\Guias\"+CPOS+".txt"
'response.write(ARCHIVO)

Set FD = FSO.OpenTextFile ( ARCHIVO , ForWriting, True)
'Escribimos su contenido 

lin = Request.QueryString("lin")
' HAY QUE LLENAR LAS LINEAS DE DETALLE DE MOVIMIENTO.....
' Y ACTUALIZAR STOCKS
for i = 1 to lin
	CODI= TRIM(Request.Form("COD"&I))
    DESI= TRIM(Request.Form("DES"&I))
	canI= (Request.Form("CAN"&I))
	PreI= (Request.Form("PRE"&I))
    if len(trim(prei)) = 0 then prei = 0
    IT  = RIGHT("00"+ CSTR(I),2) 

    DATO = CODI+","+DESI+","+CANI+","+PREI
    FD.WriteLine DATO
    response.write(dato)
    RESPONSE.Write("<BR>")	
NEXT
Set FSO = Nothing
Cnn.Close	
set Cnn = Nothing
SET RS = Nothing	
'**************************************************************
'   hasta arriba, genera el archivo TXT en el servidor WEB
'**************************************************************
strFileName = "../Guias/"+CPOS+".txt"
'RESPONSE.Write(strFilePath)
cad= "saveasTXT.asp?file=" + strFileName
%>
<script type="text/javascript">
    function guardar(doc) {
    alert(doc)
    document.execCommand('SaveAs', false, doc); // ó ejemplo.txt, pero no ejemplo.png
    alert("Documento Almacenado Correctamente")
    }
</script>





</BODY>
</HTML>
