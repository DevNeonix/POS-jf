<%@LANGUAGE="VBSCRIPT"  CODEPAGE="65001"%>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<%Usuario = Request.Cookies("tienda")("usr")%>
<%tda     = Request.Cookies("tienda")("pos")%>
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
' usuario
cod = Request.QueryString("cod")
' pwd
pwd = Request.QueryString("pwd")
aCan = split(can,",")
' perfil
per = Request.QueryString("per")

' nombre del usuario
des = Request.QueryString("des")
url = Request.QueryString("url")
chk = Request.QueryString("chk")


cad = "select USUARIO from usuarios where usuario= '"&cod&"' "
rs.open cad,cnn

if rs.recordcount > 0 and chk = "0" then
    rs.movefirst
    cod= rs.fields.item(0)
    response.Write(cad)
%>
<script type="text/jscript" language="jscript">
    parent.window.alert("Descriptor duplicado")
    cad = trim('<%=url%>') + 'perfil=1&pos=' + '<%=trim(cod)%>'
    window.parent.frames[0].location = cad
</script>

<%
response.End
end if
rs.close

CAD = " SELECT usuario FROM usuarios WHERE usuario='"&COD&"' "
'response.Write(cad)

rs.open cad,cnn
'response.Write(rs.recordcount)
'response.end
TOTALREG = rs.recordcount
RS.CLOSE
IF TOTALREG>0 THEN	
	if chk = "" then
		CAD =	" UPDATE  usuarios set          " & _
				" nombres = '"&DES&"',          " & _ 
                " clave = '"&pwd&"',            " & _
                " perfil = '"&per&"'            " & _
				" WHERE   USUARIO= '"&COD&"';   "
	else
		CAD =	" DELETE usuarios               " & _
				" WHERE usuario = '"&COD&"';    "
				cod=""
	end if	
else
   

	CAD = 	" insert into usuarios               " & _
			" (USUARIO, CLAVE, NOMBRES,          " & _
			" PERFIL)                            " & _
			" values('"&COD&"', '" & PWD & "',   " & _
			"   '"&DES&"', '"&PER&"')             " 
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

    cad = trim('<%=url%>') + 'perfil=1&pos=' + '<%=trim(cod)%>'

   parent.window.location.replace = cad

</script>
</BODY>
</HTML>

