<%@ Language=VBScript %>
<!--#include file="../includes/Cnn.inc"-->
<!--#include file="../COMUN/FUNCIONESCOMUNES.ASP"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
</head>
<body>
<html>
<%	DATO	= Request.QueryString("pos")
	pk		= Request.QueryString("pk")
	pag		= Request.QueryString("pag")
	ds		= Request.QueryString("ds")	
	tbL		= trim(Request.QueryString("alias"))
	url		= Request.QueryString("url")
	errr = ""	
	tt = len(trim(application("owner")))
    tt = tt+1
	'response.Write(tbl)
	UBI = inStr( PK, " ")
	IF ubi > 0 THEN
		PK = LEFT(PK,UBI)+ "+" + RIGHT(PK, LEN(PK)-UBI)
	END IF	
    tabla = right(tbl,(len(tbl)- tt))
CAD =	" SELECT top "&pag&" "&pk&"		" & _
		" from "&tbl&"			    	" & _
		" order by "&pk&"   desc	    "
		Response.Write(cad)
		RS.Open  CAD
	IF rs.recordcount <= 0 THEN 
	    errr = "XX" 
	else 
	    rs.movelast
	    errr= rs.fields.item(0)	
	end if
	
	'response.end
%>
	<script language="javascript" type="text/jscript">
		ss =trim('<%=errr%>') 
		if (ss!='XX')
		{	//alert("CODIGO YA EXISTE")
			cad =  '<%=URL %>'+'?pos=' + ss
			window.location.replace(cad)
		}		
	</script>
</body>
</html>
