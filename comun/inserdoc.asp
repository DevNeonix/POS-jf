<%@ Language=VBScript %>
<%Session.LCID=2058%>
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
SER = ucase(TRIM(Request.QueryString("SER")))
COR = ucase(TRIM(Request.QueryString("COR")))
TIP = ucase(TRIM(Request.QueryString("TIP")))

URL = TRIM(Request.QueryString("URL"))
largo = TRIM(Request.QueryString("largo"))
FECHA ="{ fn NOW() }"


cad = "select CODIGO, DESCRIPCION from DOCUMENTO where DESCRIPCION = '"&des&"' " & _
        " AND CIA  = '"&TDA&"'   "
response.write (cad)
rs.open cad,cnn

if rs.recordcount > 0 and chk = "0" then
    rs.movefirst
    cod= rs.fields.item(0)
    response.Write(cad)
%>
<script type="text/jscript" language="jscript">
parent.window.alert("Descriptor duplicado")
cad = trim('<%=url%>')+ 'perfil=1&pos='+ '<%=trim(cod)%>'
parent.window.location.replace = cad
</script>

<%
response.End
end if
rs.close

CAD = " SELECT CODIGO FROM DOCUMENTO WHERE CODIGO = '"&cod&"' AND CIA  = '"&TDA&"'  "
'response.Write(cad)
 COR = RIGHT("00000000000000"+ TRIM(CSTR(COR)),7)
rs.open cad,cnn
'response.Write(COR)
'response.end
TOTALREG = rs.recordcount
RS.CLOSE
IF TOTALREG>0 THEN	
	if chk = "" then   
		CAD =	" UPDATE  DOCUMENTO set         " & _
				" DESCRIPCION = '"&DES&"',      " & _
                " TIPMOV = '"&TIP&"',           " & _ 
                " SERIE= '"&SER&"',             " & _ 
                " CORREL = '"&COR&"',           " & _
				" ESTADO = 'A',                 " & _
				" USUARIO = '"&USUARIO&"',      " & _
				" FECHA = "&fecha&"             " & _
				" WHERE CODIGO = '"&cod&"' AND CIA  = '"&TDA&"'  ; "
	else
		CAD =	" UPDATE DOCUMENTO               " & _
				" SET  USUARIO = '"&USUARIO&"', " & _
				" FECHA = "&fecha&" ,           " & _
				" ESTADO = 'E'	                " & _ 
				" WHERE CODIGO = '"&cod&"' AND CIA  = '"&TDA&"'  ; "
				cod=""
	end if	
  
else
    
	CAD = 	" insert into DOCUMENTO  (CIA,      " & _
            " CODIGO, DESCRIPCION, SERIE,       " & _
            " CORREL, TIPMOV, USUARIO, FECHA,   " & _
            " ESTADO)    values(  '"&TDA&"' , '"&cod&"',    " & _
			" '"&DES&"', '"&SER&"', '"&COR&"',  " & _
            " '"&TIP&"','"&USUARIO&"',"&fecha&"," & _
            " 'A');                             " 
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
