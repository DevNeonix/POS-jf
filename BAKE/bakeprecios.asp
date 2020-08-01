<%@ Language=VBScript %>
<%TDA = Request.Cookies("TIENDA")("POS")%>
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
pos = ucase(TRIM(Request.QueryString("cod")))
OP = ucase(TRIM(Request.QueryString("OP")))

if len(trim(pos)) = 5 then
 ' es por grupo
    CAD =   " SELECT top 1 * FROM  VIEW_GRUPOS   WHERE   " & _
           " GRUPO = '"&POS&"'   "
            response.write(cad)
            response.write("<br>")
        RS.OPEN CAD,CNN
        IF rs.recordcount > 0 THEN
            des = rs("descri")
            %>
            <script language="jscript" type="text/jscript">
                op = parseInt(trim('<%=cint(op)%>'), 10)
                cad = '<%=DES %>'
                eval("window.parent.document.all.DS" + op + ".value =cad")
                eval("window.parent.document.all.LS" + op + ".focus()")
            </script>	
        <%ELSE%>
            <script language="jscript" type="text/jscript">
                op = parseInt(trim('<%=cint(op)%>'), 10)        
                alert("Articulo no ingresado a la tienda")
                eval("window.parent.document.all.CD" + op + ".focus()")
            </script>
        <%END IF

else
   CAD =   " SELECT top 1 * FROM  VIEW_ARTICULOS_TIENDA   WHERE   " & _
           " codigo = '"&POS&"'   "
            response.write(cad)
            response.write("<br>")
        RS.OPEN CAD,CNN
        IF rs.recordcount > 0 THEN
            des = rs("descri")
            pre  = rs("lista1")
            %>
            <script language="jscript" type="text/jscript">
                op = parseInt(trim('<%=cint(op)%>'), 10)
                cad = '<%=DES %>'
                pre = '<%=formatnumber(pre,2,,true) %>'
                eval("window.parent.document.all.DS" + op + ".value =cad")
                eval("window.parent.document.all.LS" + op + ".value=pre")
                eval("window.parent.document.all.LS" + op + ".focus()")
            </script>	
        <%ELSE%>
            <script language="jscript" type="text/jscript">
                op = parseInt(trim('<%=cint(op)%>'), 10)
                alert("Articulo no ingresado a la tienda")
                eval("window.parent.document.all.CD" + op + ".focus()")
            </script>
        <%END IF%>

<%END IF %>

</BODY>
</HTML>
