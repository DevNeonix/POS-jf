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

CAD = " SELECT ARTI.AR_CDESCRI AS DESCRI,       " & _
      " ARTI.AR_CUNIDAD, ARTI.AR_CMONVTA,       " & _
      " ARTI.AR_NPRECI1 AS L1,                  " & _
      " ARTI.AR_NPRECI2 AS L2,                  " & _
      " ARTI.AR_NPRECI3 AS L3                   " & _
      " FROM  RSFACCAR.dbo.AL0012ARTI as ARTI   " & _
      " WHERE AR_CCODIGO ='"&POS&"'             "
      response.write(cad)
            response.write("<br>")
rs.open cad, cnn
IF rs.recordcount > 0 THEN

    des = rs("descri")
    uni = rs("ar_cunidad")
    mon = rs("ar_cmonvta")
    l1  = 0
    l2  = rs("l2")
    l3  = rs("l3")
    stk = 0
    mnn = 1
    RS.CLOSE
    CAD =   " SELECT * FROM  VIEW_ARTICULOS_TIENDA   WHERE   " & _
            " tienda = '"&TDA&"' AND codigo = '"&POS&"'   "
            response.write(cad)
            response.write("<br>")
    RS.OPEN CAD,CNN
    IF rs.recordcount > 0 THEN
        des = rs("descri")
        uni = rs("ar_cunidad")
        mon = rs("ar_cmonvta")
        l1  = rs("lista1")
        l2  = rs("lista2")
        l3  = rs("lista3")
        stk = rs("stock")
        mnn = rs("minimo")
        CAD =   " UPDATE ARTICULOS SET ESTADO ='A' WHERE TIENDA = '"&TDA&"' " & _
                " AND CODIGO = '"&POS&"'  "
                response.write(cad)
            response.write("<br>")
        CNN.EXECUTE CAD
        %><script language="jscript" type="text/jscript">
              if (trim(top.window.parent.frames[0].document.all.thisForm.DES.value)=='')
              { alert("Ya existe un registro con este dato") }
              top.window.parent.frames[0].document.all.thisForm.COD.value = ''
              pos = '<%=trim(POS)%>'
        </script>	
        <% response.end
	END IF
ELSE
 %><script language="jscript" type="text/jscript">
       alert("Hay que crear el articulo en el Realsoft para que \nlo pueda jalar al almacen de la Tienda")
       
        </script>	
        <%
END IF	
Cnn.Close	
set Cnn = Nothing
SET RS = Nothing	
'regresa a la p�gina de donde fu� llamado, para que vea que agreg� el registro
%>
<script language="javascript" type="text/jscript">
    top.window.parent.frames[0].document.all.thisForm.DES.value = '<%=trim(des) %>'

    var subcadena = '<%=trim(uni) %>'
    var elemento =  top.window.parent.frames[0].document.all.thisForm.UNI;
    top.window.parent.frames[0].document.all.thisForm.UNI.selectedIndex = seleindice(subcadena, elemento);
    var subcadena = '<%=trim(mon) %>'
    var elemento = top.window.parent.frames[0].document.all.thisForm.MON;
    top.window.parent.frames[0].document.all.thisForm.MON.selectedIndex = seleindice(subcadena, elemento); 
    top.window.parent.frames[0].document.all.thisForm.L1.value  = '<%=trim(l1)%>'
    top.window.parent.frames[0].document.all.thisForm.L2.value  = '<%=trim(l2)%>'
    top.window.parent.frames[0].document.all.thisForm.L3.value  = '<%=trim(l3)%>'
    top.window.parent.frames[0].document.all.thisForm.MNN.value = '<%=trim(mnn)%>'
    top.window.parent.frames[0].document.all.thisForm.STK.value = '<%=trim(stk)%>'
    top.window.parent.frames[0].document.all.thisForm.PLA.selectedIndex = 2
</script>
</BODY>
</HTML>
