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
<iframe frameborder="1" style="display:block" height="100" width="100%" id="aec" name="aec"></iframe>
<%
'*********************** OJO **********************
'REVISAR EL LARGO DE LAS COLUMNAS de la tabla
'*********************** OJO **********************
pos = ucase(TRIM(Request.QueryString("cod")))
'RESPONSE.WRITE(POS)
      ops = right("00000000000"+pos,11)
CAD = " SELECT dni, bcp.fecha as fec, *                  " & _
      " FROM   dbo.BCP INNER JOIN " & _
      " dbo.TIENDAS ON dbo.BCP.TIENDA = dbo.TIENDAS.CODIGO" & _
      " WHERE dni ='"&POS&"'      "
 response.write(cad)
 rs.open cad, cnn
    
IF rs.recordcount > 0 THEN
    fec= rs("fecha")
    tda= rs("descripcion")
    compro = 1

%>
<script language="jscript" type="text/jscript">
    Pcad = 'FECHA DE COMPRA = ' + '<%=formatdatetime(rs("fec"),2)%>'
    window.parent.document.all.m1.value = ''
    window.parent.document.all.o1.value =''
    window.parent.document.all.COD.value =''
    window.parent.document.all.m1.value = Pcad
    window.parent.document.all.o1.value = ' Tienda --> ' + '<%=ucase(tda)%>'
    window.parent.document.all.COD.value = ''
</script>
<%ELSE  %>
   <script language="javascript" type="text/jscript">
       window.parent.document.all.o1.value = 'TARJETA VALIDA'
       window.parent.document.all.m1.value = ''
       cad = '../comun/inserBCP.asp?dni=' + '<%=trim(pos)%>'
       document.all.aec.src = cad
   </script>
<%END IF

response.Write(compro)	

%>

<script language="javascript" type="text/jscript">
    //alert(parseInt('<%=cint(compro)%>', 10))
   compro = (parseInt('<%=cint(compro)%>', 10))
   //alert(compro)
    window.parent.document.all.mal.style.display = 'block'
    window.parent.document.all.ok.style.display = 'block'
    window.parent.document.all.dn1.style.display = 'block'
    window.parent.document.all.dni.value = 'DNI : ' + '<%=trim(pos)%>'
  
 </script>
 <%Cnn.Close	
set Cnn = Nothing
SET RS = Nothing	 %>
</BODY>
</HTML>
