<%@ Language=VBScript %>
<!--#include file="../includes/Cnn.inc"-->
<!--#include file="../COMUN/FUNCIONESCOMUNES.ASP"-->
<html>
<head>
</head>
<html xmlns="http://www.w3.org/1999/xhtml">
<body>

<% 	DATO	= Request.QueryString("pos")
	pk		= Request.QueryString("pk")
	ds		= Request.QueryString("ds")	
	tbL		= trim(Request.QueryString("ALIAS"))
	url		= Request.QueryString("url")
	pag     = Request.QueryString("PAG")
	tt      = len(trim(application("owner")))
    tt      = tt+1
    tabla = right(tbl,(len(tbl)- tt))
	errr = ""	
	UBI = inStr( PK, " ")
	IF ubi > 0 THEN
		PK = LEFT(PK,UBI)+ "+" + RIGHT(PK, LEN(PK)-UBI)
	END IF	
CAD =	" SELECT top "&pag&" "&pk&"		" & _
		" from "&tbl&"			        " & _
		" where "&pk&" <= '"&dato&"'    " & _
		" order by "&pk&" 	DESC        "
		Response.Write(cad)
		RS.Open  CAD
	IF rs.recordcount <= 0 THEN 
	    ss=""
	else 
	    RS.movelast
	    ss= rs.fields.item(0)	
	END IF    
	
%>
	<script>
	cad =  '<%=URL %>'+'pos=' + '<%=ss%>'
	cad += '&tbl='+ '<%=tabla%>'
	cad += '&pk='+ '<%=pk%>'
	cad += '&ds='+ '<%=ds%>'
	cad += '&col=COD'
	//alert(cad)
	window.location.replace(cad)
	</script>
</BODY>
</HTML>
