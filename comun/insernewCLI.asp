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
COD = ucase(TRIM(Request.QueryString("cod")))
DES = ucase(TRIM(Request.QueryString("DES")))
DIR = ucase(TRIM(Request.QueryString("DIR")))
COR = ucase(TRIM(Request.QueryString("COR")))
TEL = ucase(TRIM(Request.QueryString("TEL")))
TIP = ucase(TRIM(Request.QueryString("TIP")))
DT1 = ucase(TRIM(Request.QueryString("DT1")))
DT2 = ucase(TRIM(Request.QueryString("DT2")))
vta = ucase(TRIM(Request.QueryString("opcion")))
FECHA ="{ fn NOW() }"


cad =   " select * from clientes where nombre = '"&des&"' or " & _
        " cliente  = '"&cod&"'  "
response.write (des)
response.write ("<Br>")
response.write ("<br>")
des = replace(des,"*", "'+'&'+'")
response.write (des)
response.write ("<Br>")
response.write ("<br>")


rs.open cad,cnn

if rs.recordcount > 0  then
    rs.movefirst
    %>
    <script type="text/jscript" language="jscript">
        window.opener.window.document.all.CLI.value = '<%=trim(rs("cliente")) %>'
        window.opener.window.document.all.DES.value = '<%=trim(rs("nombre")) %>'
        window.opener.window.document.all.DIR.value = '<%=trim(rs("direccion"))%>'
      //  window.opener.window.document.all.DS1.value = '<%=trim(rs("dcto1"))%>'
      //  window.opener.window.document.all.DS2.value = '<%=trim(rs("dcto2"))%>'
        window.close()
    </script>
    <%
end if
rs.close

	CAD = 	" insert into CLIENTES  (CLIENTE,   " & _
			" NOMBRE, DIRECCION, MAIL, TIPOCLI, " & _
            " FONO, ESTADO, usuario,  fecha,    " & _
			" DCTO1, DCTO2 )    values(         " & _
			" '"&cod&"','"&DES&"', '"&DIR&"',   " & _
            " '"&COR&"', '"&tip&"', '"&TEL&"',        " & _
			"  'A', '"&USUARIO&"', "&fecha&" ,  " & _
            " "&dt1&", "&dt2&"); exec sp_envia_mail 'se ha creado el cliente "&cod&" - "&DES&" ','rmalaga@elmodelador.com.pe;jquevedo@elmodelador.com.pe'"                

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
    window.opener.window.document.all.CLI.value = '<%=trim(cod) %>'
    window.opener.window.document.all.DES.value = '<%=trim(des) %>'
    window.opener.window.document.all.DIR.value = '<%=trim(DIR)%>'
 
  
  
 window.close()

</script>
</BODY>
</HTML>
