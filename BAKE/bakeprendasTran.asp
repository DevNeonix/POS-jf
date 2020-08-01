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
pos = ucase(TRIM(Request.QueryString("pos")))
OP = ucase(TRIM(Request.QueryString("OP")))
marka = 0
    CAD =   " SELECT * FROM  VIEW_ARTICULOS_TIENDA   WHERE   " & _
            " tienda = '"&TDA&"' AND codigo = '"&POS&"' and estado = 'A';  "
            response.write(cad)
            response.write("<br>")
    RS.OPEN CAD,CNN
    IF rs.recordcount > 0 THEN
        des = rs("descri")
        stk = rs("stock") 
        if cdbl(stk) <=0 then marka = 1
         response.write(marka)%>
        <script language="jscript" type="text/jscript">
        
            op = trim('<%=cint(op)%>')
            dd = parseInt('<%=trim(marka)%>', 10)
            if (dd == 1) 
            {   cad ='<%=pos %>'+" sin stock en el sistema"
                alert(cad)
                eval("window.parent.document.all.DES" + op + ".value =cad")
                eval("window.parent.document.all.COD" + op + ".value =''")
                eval("window.parent.document.all.COD" + op + ".focus()")
            }
            else 
            {   eval("window.parent.document.all.DES" + op + ".value = '<%=trim(des)%>'")
                eval("window.parent.document.all.STK" + op + ".value = '<%=trim(STK)%>' ")
            }

        </script>	
    <%
    ELSE%>
      <script language="jscript" type="text/jscript">
            alert("Articulo no ingresado a la tienda")
    </script>
<%END IF%>

</BODY>
</HTML>
