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
MARKA=0

if ltrim(rtrim(ucase(TRIM(Request.QueryString("tda"))))) <> "" then 
    Tda = ltrim(rtrim(ucase(TRIM(Request.QueryString("tda")))))

END IF

    CAD =   " SELECT * FROM  VIEW_ARTICULOS_TIENDA   WHERE   " & _
            " tienda = '"&TDA&"' AND codigo = '"&POS&"' and estado = 'A';  "
            response.write(cad)
            response.write("<br>")
    RS.OPEN CAD,CNN
    IF rs.recordcount > 0 THEN
        des = rs("descri")
        l1  = rs("lista1")
        stk = rs("stock") 
        IsC = rs("ISC") 
        if cdbl(stk) <=0 then marka = 1
         response.write(marka)%>
        <script language="jscript" type="text/jscript">
        
            op = trim('<%=cint(op)%>')
            dd = parseInt('<%=trim(marka)%>', 10)
            l1 = parseFloat('<%=trim(l1)%>', 10)
            cd = trim('<%=trim(pos)%>')


            if (dd == 1) {
            cad ='<%=pos %>'+" sin stock en el sistema"
                alert(cad)
                eval("window.parent.document.all.DES" + op + ".value =cad")
                eval("window.parent.document.all.COD" + op + ".value =''")
                eval("window.parent.document.all.COD" + op + ".focus()")
            }
            else {
                if (l1 == 0  && cd != 'SALDO00000' ) {
                    cad = '<%=pos %>' + " sin PRECIO en el sistema"
                    alert(cad)
                    eval("window.parent.document.all.DES" + op + ".value =cad")
                    eval("window.parent.document.all.COD" + op + ".value =''")
                    eval("window.parent.document.all.COD" + op + ".focus()")
                }
                else {
                    l1 = Math.round(parseFloat('<%=trim(l1)%>') * 100) / 100
                    eval("window.parent.document.all.DES" + op + ".value = '<%=trim(des)%>'")
                    eval("window.parent.document.all.PVP" + op + ".value = l1")

                    eval("window.parent.document.all.STK" + op + ".value = '<%=trim(STK)%>' ")
                    eval("window.parent.document.all.ISC" + op + ".value = '<%=trim(ISC)%>' ")
                    eval("window.parent.document.all.III" + op + ".value = '<%=trim(ISC)%>' ")
                    eval("window.parent.document.all.CAN" + op + ".focus()")

                    //alert('<%=trim(STK)%>')
                }
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
