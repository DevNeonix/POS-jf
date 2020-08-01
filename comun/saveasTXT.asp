<%@ Language=VBScript%>
<%Response.Buffer = true %>
<%Session.LCID=2058%>
<%Usuario = Request.Cookies("tienda")("usr")%>
<%tda     = Request.Cookies("tienda")("pos")%>
<%destda  = Request.Cookies("tienda")("tda") %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<%
cfile = TRIM(request.form("file"))	
Response.Charset = "UTF-8"
Response.ContentType = "text/html"
Response.AddHeader "Content-Disposition", "attachment; filename=" & cfile 
%>
<body>
<script language="javascript" type="text/jscript">
</script>
</BODY>
</HTML>
