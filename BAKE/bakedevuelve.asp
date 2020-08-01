<%@ Language=VBScript %>
<%TDA = Request.Cookies("TIENDA")("POS")%>
<%Response.Buffer = TRUE %>
<%Session.LCID=2058%>
<!--#include file="../includes/Cnn.inc"-->
<!--#include file="../comun/funcionescomunes.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<body>
<script language="javascript" type="text/jscript">
    var OPE
    var COD
    var PDV
    var CASO
    CASO = 0
</script>
<%
'*********************** OJO **********************
'REVISAR EL LARGO DE LAS COLUMNAS de la tabla
'*********************** OJO **********************
ser = ucase(TRIM(Request.QueryString("ser")))
nro = ucase(TRIM(Request.QueryString("nro")))
cod = ucase(TRIM(Request.QueryString("cod")))
tip = ucase(TRIM(Request.QueryString("tip")))
caso = 0

'RESPONSE.WRITE(POS)
 
' vemos si el producto EXISTE
 cad = "select top 1 descri from view_articulos_tienda where codigo = '"&cod&"'"
 rs.open cad, cnn
       
IF rs.recordcount <= 0 THEN %>
    <script language="javascript" type="text/jscript">
        parent.window.alert("Código de Artículo no encontrado en ninguna tienda")
    </script>
<%  response.end
else
' existe en el maestro de ARTICULOS....
%>
      <script language="javascript" type="text/jscript">
          top.parent.window.frames[0].window.frames[0].window.document.all.DES.value = '<%=trim(rs("descri"))%>'
    </script>
<%end if

' VEMOS SI EL DOCUMENTO EXISTE
 CAD = " select * from movimcab where serie = '"&ser&"' and numdoc= '"&nro&"' and coddoc = '"&tip&"'   "
 rs.close
 rs.open cad, cnn
      response.write(cad)
      response.write("<br>")
IF rs.recordcount <= 0 THEN 
' NO EXISTE PARA NINGUNA TIENDA
%>
    <script language="javascript" type="text/jscript">
        window.parent.alert("Documento no existe")
    </script>
<%  response.end
else
    rs.movefirst
    ope = rs("operacion")
    tienda = rs("tienda")

end if
' VERIFICA QUE EL PRODUCTO SE HAYA VENDIDO CON EL DOCUMENTO DE LA REFERENCIA
cad = "select * from movimdet where operacion ='"&ope&"' and codart = '"&cod&"' "
    rs.close
    rs.open cad, cnn 
    response.write(cad)
IF rs.recordcount <= 0 THEN %>
    <script language="javascript" type="text/jscript">
      
        //top.parent.window.alert("Producto NO pertenece al Documento informado")
        top.parent.window.frames[0].window.frames[1].window.document.all.MSG.value = 'ATIPICO .- ARTICULO no vendido con el Documento informado '
        top.parent.window.frames[0].window.frames[2].window.location.replace('../blanco.htm')
    </script>
<%RESPONSE.End
END IF
    if tienda <> tdA THEN%>
     <script language="javascript" type="text/jscript">
      //   top.parent.window.alert("Producto NO VENDIDO en esta tienda")
    </script>
    <%end if' el documento Y EL ARTÍCULO existeN
    ' CASO 1 .- SI pertenece a la misma tienda Y EL FLAG DE DEVOLUCION ESTA EN BLANCO
    cad = "select * from movimdet where operacion ='"&ope&"' and codart = '"&cod&"' and flag = '' and tienda = '"&tda&"'"
    rs.close
    rs.open cad, cnn 
    response.write(cad)
    response.write("<br>")

IF rs.recordcount <= 0 THEN 
  cad = "select top 1 * from movimdet BB INNER JOIN movimcab as dd ON BB.OPERACION = DD.OPERACION where BB.codart = '"&cod&"' and BB.flag = '' and BB.tienda = '"&tda&"' and dd.coddoc <> 'VL' AND DD.TIPMOV = 'S' "
    rs.close
    rs.open cad, cnn    
    IF rs.recordcount <= 0 THEN %>
        <script language="javascript" type="text/jscript">
          //  window.parent.alert("No hay Artículo liberado vendidos en esta tienda\nBuscando otras tiendas")
        </script>
        <%  cad = "select top 1 * from movimdet BB  INNER JOIN movimcab as dd ON BB.OPERACION = DD.OPERACION where BB.codart = '"&cod&"' and BB.flag = ''  AND DD.TIPMOV='S'"
        rs.close
        rs.open cad, cnn    
        IF rs.recordcount <= 0 THEN %>
            <script language="javascript" type="text/jscript">
                window.parent.alert("No hay Artículo liberado/vendido en ninguna tienda")
            </script>
        <%ELSE%>
        <script language="javascript" type="text/jscript">
            cad = 'CASO 3 .- Producto Vendido en otra tienda Y EL FLAG DE DEVOLUCION ESTA EN BLANCO '
            top.parent.window.frames[0].window.frames[1].window.document.all.MSG.value = cad
            OPE = '<%=RS("OPERACION")%>'
            COD = '<%=RS("CODART")%>'
            PDV = '<%=rs("tienda")%>'
            CASO = '3'
        </script><%
        end if
    ELSE%>
        <script language="javascript" type="text/jscript">               
            cad ='CASO 2 .- Producto Vendido en la misma tienda Y EL FLAG DE DEVOLUCION ESTA EN BLANCO se aplica a documento nuevo'
            top.parent.window.frames[0].window.frames[1].window.document.all.MSG.value = cad
            OPE = '<%=RS("OPERACION")%>'
            COD = '<%=RS("CODART")%>'
            PDV = '<%=rs("tienda")%>'
            CASO = '2'
            alert(ope)
        </script><%
    end if

ELSE %>
    <script language="javascript" type="text/jscript">
        top.parent.window.frames[0].window.frames[1].window.document.all.MSG.value = 'CASO 1 .- DOCUMENTO pertenece a la misma tienda Y EL FLAG DE DEVOLUCION ESTA EN BLANCO'
        OPE = '<%=RS("OPERACION")%>'
        COD = '<%=RS("CODART")%>'
        PDV = '<%=rs("tienda")%>'
        CASO = '1'
    </script><%
end if
%>
<script language="javascript" type="text/jscript">
    cad = '../notadeta.asp?pos=' + trim(COD) + '&caso=' + trim(CASO)
    cad += '&ope=' + trim(OPE)
    cad += '&PDV=' + trim(PDV)
   // alert(cad)
    top.parent.window.frames[0].window.frames[2].window.location.replace(cad)




</script><%
RS.CLOSE	
Cnn.Close	
set Cnn = Nothing
SET RS = Nothing	
'regresa a la página de donde fué llamado, para que vea que agregó el registro
%>

</BODY>
</HTML>
