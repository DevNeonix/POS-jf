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
'RESPONSE.WRITE(POS)
      response.write("<br>")
      ops = right("00000000000"+pos,8)
      opp = right("00000000000"+pos,11)
CAD = " SELECT *                  " & _
      " FROM  CLIENTES            " & _
      " WHERE (CLIENTE ='"&POS&"' or cliente = '"&ops&"' or cliente = '"&opp&"') and estado = 'a'"
 rs.open cad, cnn

      response.write(RS.RECORDCOUNT)
      response.write("<br>")

IF rs.recordcount > 0 THEN
    rs.movefirst
    cli = rs("cliente")
    des = rs("NOMBRE")
    DIR = rs("DIRECCION")
    DC1 = rs("DCTO1")
    DC2 = rs("DCTO2")
   RS.CLOSE
ELSE%>
    <script language="javascript" type="text/jscript">
   
       var opc = "directories=no,height=400,width=600, ";
        opc = opc + "hotkeys=no,location=no,";
        opc = opc + "menubar=no,resizable=no,";
        opc = opc + "left=0,top=0,scrollbars=no,";
        opc = opc + "status=no,titlebar=no,toolbar=no,";
        cad = "../help/newcliente.asp?COD=" + '<%=trim(pos)%>'
        window.parent.document.all.CLI.value = ''
      //  alert(cad)
        parent.open(cad,"",'opc,_new')     
        

      //  window.parent.NEWCLI()
    </script>
<%  END IF	
Cnn.Close	
set Cnn = Nothing
SET RS = Nothing	
'regresa a la página de donde fué llamado, para que vea que agregó el registro
%>
<script language="javascript" type="text/jscript">
//alert()
   window.parent.document.all.DES.value = '<%=trim(des)%>'
   window.parent.document.all.DIR.value = '<%=trim(DIR)%>'
  // window.parent.document.all.CLI.value = '<%=trim(cli)%>'
 //  window.parent.document.all.DS2.value = '<%=trim(DC2)%>'
 </script>
</BODY>
</HTML>
