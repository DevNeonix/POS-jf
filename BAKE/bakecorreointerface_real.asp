<%@ Language=VBScript %>
<%TDA = Request.Cookies("TIENDA")("POS")%>
<%Response.Buffer = TRUE %>
<!--#include file="../includes/Cnn.inc"-->
<!--#include file="../comun/funcionescomunes.asp"-->
<link REL="stylesheet" TYPE="text/css" HREF="../ventas.CSS" >
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>

<body>
<script type="text/jscript" language="jscript">
//alert("entro")
</script>
<%
'*********************** OJO **********************
'REVISAR EL LARGO DE LAS COLUMNAS de la tabla
'*********************** OJO **********************
TDA = ucase(TRIM(Request.QueryString("pos")))
INI = ucase(TRIM(Request.QueryString("INI")))
FIN = ucase(TRIM(Request.QueryString("FIN")))

cuerpo = "Se ha corrido el proceso para la tienda " + tda + " entre las fechas :" + ini + " y " + fin
tit = "Interface de Facturación"
RECIBEN = "sistemas@elmodelador.com.pe; jchuquillanqui@elmodelador.com.pe; csaba@elmodelador.com.pe; rmalaga@elmodelador.com.pe;"
CAD =   " set dateformat dmy;                                           " & _
        " SELECT DISTINCT  TOP 100 PERCENT  M1.coddoc, M1.SERIE,        " & _
        " M1.NUMDOC, M1.FECDOC, m1.Pvp                                  " & _
        " FROM MOVIMCAB  AS M1                                          " & _
        " INNER JOIN TIENDAS AS T1 ON T1.CODIGO = M1.TIENDA             " & _
        " INNER JOIN CLIENTES AS C1 ON C1.CLIENTE = M1.CLIENTE          " & _
        " WHERE FECdoc between '"&ini&"' AND DateAdd(day,1,'"&fin&"')   " & _
        " AND CONTA = '*' and M1.TIENDA =ltrim(rtrim('"&tda&"'))        " & _
        " and tipmov = 'S' AND  PVP >0                                  " & _
        " ORDER BY 	 1,2,3                                              "
rs.open cad, cnn
IF RS.RECORDCOUNT > 0 THEN
     cuerpo = "<p> Documentos transferidos a Contabilidad Tienda : " +  tda + " entre las fechas :" + ini + " y " + fin + "</p>"
     cuerpo = cuerpo +  " <table border=1 cellpadding=5 cellspacing=0> " & _
                        " <tr style=''background: gainsboro; color: black;font-weight: bold;''> " & _
                        " <td>FECHA</td> <td>DOC</td> <td>SERIE</td> <td>NUMERO</td> </tr>     " 
    rs.movefirst
    do while not rs.eof
        cuerpo = cuerpo + "<tr>"
        cuerpo = cuerpo + "<td>"+formatdatetime(rs("fecdoc"),2)+"</td>"
        cuerpo = cuerpo + "<td>"+rs("coddoc")+"</td>"
        cuerpo = cuerpo + "<td>"+rs("serie")+"</td>"
        cuerpo = cuerpo + "<td>"+rs("numdoc")+"</td>"
        cuerpo = cuerpo + "</tr>"
        rs.movenext
    loop
    cuerpo = cuerpo + "</table>"       
    
    cad =  "  EXEC msdb.dbo.sp_send_dbmail                  " & _
           " @profile_name	= 'DBMailProfile',              " & _
	       " @recipients	= '"&RECIBEN&"',                " &  _
           " @body_format	= 'HTML',                       " & _
           " @importance	= 'High',                       " & _
	       " @body			= '"&cuerpo&"',	                " & _
	       " @subject		= '"&tit&"'                     "
    'response.write(cad)
    cnn.EXECUTE (CAD)
    rs.close
    '*********************************************************************************
    ' FALTA BUSCAR DOCUMENTOS ANULADOS ENTRE LAS FECHAS DADAS
    ' y MANDAR CORREO CON LOS DOCUMENTOS ANULADOS PARA QUE LOS HAGAN MANUALMENTE
    '*********************************************************************************
    CAD =   " set dateformat dmy;                                           " & _
            " SELECT DISTINCT  TOP 100 PERCENT  M1.coddoc, M1.SERIE,        " & _
            " M1.NUMDOC, M1.FECDOC, m1.Pvp                                  " & _
            " FROM MOVIMCAB  AS M1                                          " & _
            " INNER JOIN TIENDAS AS T1 ON T1.CODIGO = M1.TIENDA             " & _
            " INNER JOIN CLIENTES AS C1 ON C1.CLIENTE = M1.CLIENTE          " & _
            " WHERE FECdoc between '"&ini&"' AND DateAdd(day,1,'"&fin&"')   " & _
            " AND CONTA = '*' and M1.TIENDA =ltrim(rtrim('"&tda&"'))        " & _
            " and tipmov = 'S' AND  PVP =0                                  " & _
            " ORDER BY 	 1,2,3                                              "
    rs.open cad, cnn
    IF RS.RECORDCOUNT > 0 THEN
        tit = "Documentos Anulados de Inteface de Facturación para la tienda " + tda + " entre las fechas :" + ini + " y " + fin
        RECI = "sistemas@elmodelador.com.pe;csaba@elmodelador.com.pe; rmalaga@elmodelador.com.pe;"
        cuerpo =    " <table border=1 cellpadding=5 cellspacing=0> " & _
                    " <tr style=''background: gainsboro; color: black;font-weight: bold;''> " & _
                    " <td>FECHA</td> <td>DOC</td> <td>SERIE</td> <td>NUMERO</td> </tr>     " 
                   rs.movefirst
                   do while not rs.eof
                       cuerpo = cuerpo + "<tr>"
                       cuerpo = cuerpo + "<td>"+formatdatetime(rs("fecdoc"),2)+"</td>"
                       cuerpo = cuerpo + "<td>"+rs("coddoc")+"</td>"
                       cuerpo = cuerpo + "<td>"+rs("serie")+"</td>"
                       cuerpo = cuerpo + "<td>"+rs("numdoc")+"</td>"
                       cuerpo = cuerpo + "</tr>"
                       rs.movenext
                   loop
                  cuerpo = cuerpo + "</table>"       
         cad = "  EXEC msdb.dbo.sp_send_dbmail                  " & _
               " @body_format	= 'HTML',                       " & _
               " @importance	= 'High',                       " & _
	           " @body			= '"&CUERPO&"',	                " & _
               " @profile_name	= 'DBMailProfile',              " & _
	           " @recipients	= '"&RECI&"',                   " & _
	           " @subject		= '"&tit&"'                     "
        'response.write(cad)
        cnn.EXECUTE (CAD)
    END IF


end if

%>

<script type="text/jscript" language="jscript">
 //   alert("salio")
</script>

</BODY>
</HTML>
