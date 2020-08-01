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
	tbL		= trim(Request.QueryString("tbl"))
	url		= Request.QueryString("url")
	ppp     = Request.QueryString("size")
	tt = len(trim(application("owner")))
    tt = tt+1
    tabla = right(tbl,(len(tbl)- tt))
	errr = ""	
CAD =	" SELECT top "&ppp&" "&pk&"		" & _
		" from "&tbl&"			        " & _
		" where "&pk&" <= '"&dato&"'    " & _
		" order by "&pk&" 	DESC        "
		Response.Write(cad)
		RS.Open  CAD
	IF rs.recordcount <= 0 THEN 
	    errr = "XX" 
	else 
	    RS.MOVELAST
	    errr= rs.fields.item(0)	
	END IF    
	
%>
	<script>
		ss =trim('<%=errr%>') 
		if (ss!='XX')
		{	cad =  '../tablas.asp?pos=' + ss
			cad += '&tbl='+ '<%=tabla%>'
			cad += '&pk='+ '<%=pk%>'
			cad += '&ds='+ '<%=ds%>'
			cad += '&col=COD'
			top.window.location.replace(cad)
		}		
	</script>
</BODY>
</HTML>
