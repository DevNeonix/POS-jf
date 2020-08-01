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
CHK = Request.QueryString("chk")
COD = ucase(TRIM(Request.QueryString("cod")))
MNN = ucase(TRIM(Request.QueryString("MNN")))
L1 = ucase(TRIM(Request.QueryString("L1")))
L2 = ucase(TRIM(Request.QueryString("L2")))
L3 = ucase(TRIM(Request.QueryString("L3")))
pla = ucase(TRIM(Request.QueryString("pla")))
STK = ucase(TRIM(Request.QueryString("STK")))
FECHA ="{ fn NOW() }"


 CAD =   " SELECT * FROM  ARTICULOS   WHERE   " & _
         " tienda = '"&TDA&"' AND codigo = '"&cod&"'   "
'response.write (cad)
rs.open cad,cnn

TOTALREG = rs.recordcount
RS.CLOSE
IF TOTALREG>0 THEN	
	if chk = "" then
		CAD =	" UPDATE articulos set          " & _
				" lista1 = '"&l1&"',            " & _
                " lista2 = '"&l2&"',            " & _ 
                " lista3 = '"&l3&"',            " & _  
                " minimo = '"&mnn&"',           " & _ 
                " planilla = "&pla&",           " & _
                " STOCK = "&STK&",           " & _
                " ESTADO = 'A',                 " & _
				" USUARIO = '"&USUARIO&"',      " & _
				" FECHA = "&fecha&"             " & _
				" WHERE  tienda = '"&TDA&"' AND " & _
                " codigo = '"&cod&"'   ;        "
	else
		CAD =	" UPDATE articulos              " & _
				" SET  USUARIO = '"&USUARIO&"', " & _
				" FECHA = "&fecha&" ,           " & _
				" ESTADO = 'E'	                " & _ 
				" WHERE  tienda = '"&TDA&"' AND " & _
                " codigo = '"&cod&"'   ;        "
				cod=""
	end if	
else   
	CAD = 	" insert into articulos  (TIENDA,   " & _
			" CODIGO, MINIMO, STOCK, LISTA1,    " & _
            " LISTA2, LISTA3, ESTADO, planilla, " & _
			" usuario, fecha )    values(       " & _
			" '"&TDA&"','"&COD&"', '"&MNN&"',   " & _
            " "&STK&",   '"&l1&"', '"&l2&"',    " & _
            "  '"&l3&"', 'A', "&pla&",          " & _
            " '"&USUARIO&"',    "&fecha&" );    " 
End if
' Inicia transacci�n , para que los datos no queden a medias
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
'regresa a la p�gina de donde fu� llamado, para que vea que agreg� el registro
cad = "..\articulos.asp?pos="+ trim(cod)
'window.response.Redirect(cad)

%>

<script language="javascript" type="text/jscript">

cad = '../ARTICULOS.ASP?POS='+ '<%=trim(cod)%>'
//top.window.parent.frames[0].location = cad
parent.window.location.replace = cad

</script>
</BODY>
</HTML>
