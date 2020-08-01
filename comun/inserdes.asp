<%@LANGUAGE="VBSCRIPT"  CODEPAGE="65001"%>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<%	tienda = Request.Cookies("tienda")("pos") 	%>
<%	USUARIO = Request.Cookies("tienda")("USR") 	%>
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
Dct = ucase(TRIM(Request.QueryString("Dct")))
TBL = TRIM(Request.QueryString("TBL"))
PK  = TRIM(Request.QueryString("PK"))
DS  = TRIM(Request.QueryString("DS"))
URL = TRIM(Request.QueryString("URL"))
largo = TRIM(Request.QueryString("largo"))
FECHA ="{ fn NOW() }"


cad = "select codigo, descripcion from descuentos where DESCRIPCION = '"&des&"' and estado= 'a'"
rs.open cad,cnn

if rs.recordcount > 0 and chk = "0" then
    rs.movefirst
    cod= rs.fields.item(0)
    response.Write(cad)
%>
<script type="text/jscript" language="jscript">
parent.window.alert("Descriptor duplicado")
cad = trim('<%=url%>')+ 'perfil=1&pos='+ '<%=trim(cod)%>'
window.parent.frames[0].location = cad
</script>

<%
response.End
end if
rs.close

CAD = " SELECT CODIGO FROM DESCUENTOS WHERE CODIGO='"&COD&"' "
'response.Write(cad)

rs.open cad,cnn
'response.Write(rs.recordcount)
'response.end
TOTALREG = rs.recordcount
RS.CLOSE
IF TOTALREG>0 THEN	
	if chk = "" then
		CAD =	" UPDATE  DESCUENTOS set        " & _
				" DESCRIPCION = '"&DES&"',      " & _ 
                " VALOR = "&dct&",              " & _
				" ESTADO = 'A',                 " & _
				" USUARIO = '"&USUARIO&"',      " & _
				" FECHA = "&fecha&"             " & _
				" WHERE   CODIGO = '"&COD&"';   "
	else
		CAD =	" UPDATE DESCUENTOS                " & _
				" SET  USUARIO = '"&USUARIO&"', " & _
				" FECHA = "&fecha&" ,           " & _
				" ESTADO = 'E'	                " & _ 
				" WHERE CODIGO = '"&COD&"';     "
				cod=""
	end if	
else
    if 	trim(cod) = "AUTO" then
        rs.open "select count(*) from DESCUENTOS" , cnn
        xx = ""
        for i=0 to largo
            xx = xx + "0"
        next
        if rs.recordcount > 0  then
            cod = xx + trim(cstr(rs.fields.item(0)+1))    
            cod = right(cod,largo)
        ELSE
            cod = xx
        end if    
    end if        

	CAD = 	" insert into DESCUENTOS            " & _
			" (codigo, descripcion, ESTADO,          " & _
			" usuario, fecha, VALOR )           " & _
			" values('"&COD&"', '" & DES & "',  " & _
			"  'A', '"&USUARIO&"', "&fecha&",   " & _
            " '"&dct&"' )  " 
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

cad = trim('<%=url%>')+ 'perfil=1&pos='+ '<%=trim(cod)%>'

parent.window.location.replace = cad

</script>
</BODY>
</HTML>
