<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<% tienda = Request.Cookies("tienda")("pos") %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<frameset rows="150,*" frameborder="no" border="0" framespacing="0">
  <frameset rows="*,40" frameborder="no" border="0" framespacing="0">
    <frame src="notaHEADconti.asp" name="topFrame" scrolling="No" noresize="noresize" id="topFrame" title="topFrame" />
    <frame src="mensajeconti.asp" name="bottomFrame" scrolling="No" noresize="noresize" id="bottomFrame" title="bottomFrame" />
  </frameset>
  <frame src="notaDETAconti.asp" name="mainFrame" id="mainFrame" title="mainFrame" />
</frameset>

<noframes>
    <body>
    </body>
</noframes>
</html>
