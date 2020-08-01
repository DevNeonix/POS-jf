<%@LANGUAGE="VBSCRIPT"  CODEPAGE="65001"%>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<!--#include file="includes/Cnn.inc"-->
<!--#include file="comun/funcionescomunes.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<body>
<% 
tbl = request.QueryString("pos")
CAD = CAD + " select * from sunat.."&TBL&" order by it; "
rs.open cad,cnn
rs.movefirst
LU= 0

do while not rs.eof
    cod = rs("codigo")
    des = rs("descrip")  
    response.write(ISNULL(COD))
    ' RESPONSE.END
    RS.MOVENEXT
    do while not rs.eof AND ISNULL(RS("CODIGO"))
	    id = rs("it")
        response.write(COD)
        cad =  "update sunat.."&tbl&" set codigo = '"&cod&"', descrip = '"&des&"' where it = "&id&""
        
        response.write("<br />")
         response.write(cad) 
      '  cnn.execute (cad)
		RS.MOVENEXT
        'RESPONSE.END
		if rs.eof then exit do
	LOOP

    if rs.eof then exit do
LOOP












Cnn.Close	
set Cnn = Nothing
SET RS = Nothing	

%>


</BODY>
</HTML>
