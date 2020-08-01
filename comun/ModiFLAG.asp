<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = true%>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos")%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<link REL="stylesheet" TYPE="text/css" HREF="../ventas.CSS" >
<!--#include file="funcionescomunes.asp"-->
<!--#include file="../includes/funcionesVBscript.asp"-->
<!--#include file="../includes/cnn.inc"-->

<%
ope = request.QueryString("ope")
lin = request.QueryString("lin")

response.write(lin)

cad =   " SET DATEFORMAT DMY;           " & _
        " UPDATE movimdet SET           " & _
        " flag = ''                     " & _
        " where operacion = '"&ope&"'   " & _
        " and convert(int,item) in ("&lin&") "
        response.write(cad)
CNN.EXECUTE CAD    
    
%>

<body></body>
<script language="jscript" type="text/jscript">

alert("datos actualizados")
parent.window.location.replace('../blanco.htm')
</script>
</html>
