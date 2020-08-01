<%@ Language=VBScript %>
<%Usuario = Request.Cookies("tienda")("usr")%>
<%tda     = Request.Cookies("tienda")("pos")%>
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
CHK = Request.QueryString("chk")
COD = ucase(TRIM(Request.QueryString("cod")))
DES = ucase(TRIM(Request.QueryString("DES")))
DIR = ucase(TRIM(Request.QueryString("DIR")))
COR = ucase(TRIM(Request.QueryString("COR")))
TEL = ucase(TRIM(Request.QueryString("TEL")))
TIP = ucase(TRIM(Request.QueryString("TIP")))
DT1 = ucase(TRIM(Request.QueryString("DT1")))
DT2 = ucase(TRIM(Request.QueryString("DT2")))
URL = TRIM(Request.QueryString("URL"))
largo = TRIM(Request.QueryString("largo"))
FECHA ="{ fn NOW() }"


cad = "select CLIENTE, nombre from clientes where nombre = '"&des&"' and estado= 'a'  "
'response.write (cad)
rs.open cad,cnn

if rs.recordcount > 0 and chk = "0" then
    rs.movefirst
    cod= rs.fields.item(0)
%>
<script type="text/jscript" language="jscript">
parent.window.alert("Descriptor duplicado")
cad = trim('<%=url%>')+ 'perfil=1&pos='+ '<%=trim(cod)%>'
window.parent.frames[1].location= cad
</script>

<%
response.End
end if
rs.close

CAD = " SELECT CLIENTE FROM CLIENTES WHERE CLIENTE = '"&cod&"' "
response.Write(des)

rs.open cad,cnn
'response.Write(rs.recordcount)
'response.end
TOTALREG = rs.recordcount
RS.CLOSE

des = replace(des,"*", "'+'&'+'")

IF TOTALREG>0 THEN	
	if chk = "" then
		CAD =	" UPDATE  CLIENTES set          " & _
				" NOMBRE = '"&DES&"',           " & _
                " DIRECCION = '"&DIR&"',        " & _ 
                " mail = '"&COR&"',             " & _  
                " TIPOCLI = '"&TIP&"',          " & _ 
                " FONO = '"&TEL&"',             " & _ 
                " DCTO1 = "&dt1&",              " & _
                " DCTO2 = "&DT2&",              " & _
				" ESTADO = 'A',                 " & _
				" USUARIO = '"&USUARIO&"',      " & _
				" FECHA = "&fecha&"             " & _
				" WHERE   CLIENTE = '"&COD&"' ; "
	else
		CAD =	" UPDATE CLIENTES               " & _
				" SET  USUARIO = '"&USUARIO&"', " & _
				" FECHA = "&fecha&" ,           " & _
				" ESTADO = 'E'	                " & _ 
				" WHERE   cliente = '"&COD&"';   "
				cod=""
	end if	
else
   'COD = RIGHT("00000000000000"+ TRIM(COD),11)
	CAD = 	" insert into CLIENTES  (CLIENTE,   " & _
			" NOMBRE, DIRECCION, MAIL, TIPOCLI, " & _
            " FONO, ESTADO, usuario,  fecha,    " & _
			" DCTO1, DCTO2 )    values(         " & _
			" '"&cod&"','"&DES&"', '"&DIR&"',   " & _
            " '"&COR&"', '"&TIP&"', '"&TEL&"',  " & _
			"  'A', '"&USUARIO&"', "&fecha&" ,  " & _
            " "&dt1&", "&dt2&");                " 
End if
' Inicia transacción , para que los datos no queden a medias
RESPONSE.WRITE("<br>")
RESPONSE.WRITE(CAD)
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

cad = trim('<%=url%>')+ '?perfil=1&pos='+ '<%=trim(cod)%>'
parent.window.location.replace = cad

</script>
</BODY>
</HTML>
