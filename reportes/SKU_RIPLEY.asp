<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = true %>
<%Session.LCID=2058%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<LINK REL="stylesheet" TYPE="text/css" HREF="../VENTAS.CSS">
<!--#include file="../comun/funcionescomunes.asp"-->
<!--#include file="../includes/funcionesVBscript.asp"-->
<!--#include file="../includes/cnn.inc"-->

<%
ini = request.querystring("ini")
fin = request.querystring("fin")
if request.querystring("xl") = "1" then 

    archivo = "c:\temp\ripley.xls"
    Response.Charset = "UTF-8"
    Response.ContentType = "application/vnd.ms-excel" 
    Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo 
end if

Set RS2 = Server.CreateObject("ADODB.Recordset")
      RS2.ActiveConnection = Cnn
      RS2.CursorType       = 3 'CONST adOpenStatic = 3
      RS2.LockType         = 1 'CONST adReadOnly = 1
      RS2.CursorLocation   = 3 'CONST adUseClient = 3
ini = request.querystring("ini")
fin = request.querystring("fin")




cad =   " SELECT LEFT(AT_CCODIGO,5) AS COD,                                                                                         " & _
        " (SELECT TG_CDESCRI FROM RSFACCAR..AL0012TABL WHERE TG_CCOD = '39' AND TG_CCLAVE = SUBSTRING(AT_CCODIGO,1,5)) AS DESCRI,   " & _
        " (SELECT TG_CDESCRI FROM RSFACCAR..AL0012TABL WHERE TG_CCOD = 'D2' AND TG_CCLAVE = SUBSTRING(AT_CCODIGO,9,2)) AS COLOR,    " & _
        " (SELECT TG_CDESCRI FROM RSFACCAR..AL0012TABL WHERE TG_CCOD = 'D1' AND TG_CCLAVE = SUBSTRING(AT_CCODIGO,6,3 )) AS TALLA,   " & _
        " SUM(CANT) AS OCT_FEB                                                                                                      " & _
        " FROM SKU_RIPLEY                                                                                                           " & _
        " where anio+mes >= '"&ini&"' and anio+mes <= '"&fin&"'                                                                     " & _
        " GROUP BY  AR_CDESCRI, SUBSTRING(AT_CCODIGO,6,3 ), SUBSTRING(AT_CCODIGO,9,2), SUBSTRING(AT_CCODIGO,1,5)                    " & _
        " ORDER BY SUBSTRING(AT_CCODIGO,1,5), COLOR, TALLA     " 
'RESPONSE.WRITE (CAD)
rs.open cad,cnn
if rs.recordcount <=0 then
    response.write("no hay nada que mostrar")
    RESPONSE.End
end if

%>
<script type="text/jscript" language="jscript">
    excel = 'sku_ripley.asp?xl=1'+ '&ini=' + '<%=trim(ini) %>' + '&fin=' + '<%=fin%>'
</script>

<body>

<center>

<table align="center" cellpadding="2" cellspacing="0" border="0"  >
<tr><td class ="Estilo17">

<%
aMes = array("" ,"ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SET","OCT","NOV","DIC")

inicio = aMes(cint(right(trim(ini),2))) + "-" + left(trim(ini),4)
final  = aMes(cint(right(trim(fin),2))) + "-" + left(trim(fin ),4)%>

Reporte de Ventas por Articulo Ripley <%=inicio %> - <%=final %></td>
<td><%if request.querystring("xl") = "" then %><img src="../images/xl1.png" style="cursor:pointer" onclick="window.location.replace(excel)" /><%end if%></td>
</tr>
</table>



<table align="center" cellpadding="0" cellspacing="2" border="0" >
	<tr class="Estilo3"> 
   		<td align="center" >Codigo</td>
        <td align="center" >Descripcion</td>
        <td align="center" >Colores</td>
        <td colspan="10" align="center">Tallas</td>
        <td align="center" >TOTAL</td>
	</tr>
	<%lin =0 %>
    <%do while not rs.eof%> 
        <%cod = rs("cod")%>
        <tr class="Estilo25" style="height:30px;"> 
            <td align="center"><%=rs("cod") %></td>
            <td align="left" ><%=rs("descri") %></td>
            <td></td>
            <%kad = "SELECT DISTINCT  (SELECT TG_CDESCRI FROM RSFACCAR..AL0012TABL               " & _
                    " WHERE TG_CCOD = 'D1' AND TG_CCLAVE = SUBSTRING(AT_CCODIGO,6,3 )) AS TALLA  " & _
                    " FROM SKU_RIPLEY WHERE LEFT(AT_CCODIGO,5) = '"&cod&"'                       " & _
                    " and anio+mes >= '"&ini&"' and anio+mes <= '"&fin&"'                        " 
                    rs2.open kad,cnn %>
            <%rs2.movefirst
            i = 0
            aTal = array("","","","","","","","","","","")
            do while not rs2.eof%>
                <td><%=rs2("talla")%></td>
                <%aTal(i)=rs2("talla") %>
            <%rs2.movenext%>
            <%i=i +1 %>
            <%loop %>
            <%rs2.close%>
            
           <td colspan="<%=ubound(atal)-i %>">&nbsp;</td>
           <%KAD =  " SELECT DISTINCT LEFT(AT_CCODIGO,5) AS COD,      " & _
                    " SUM(CANT) AS TOTA FROM SKU_RIPLEY               " & _
                    " where  LEFT(AT_CCODIGO,5) = '"&COD&"'  and      " & _
                    " anio+mes >= '"&ini&"' and anio+mes <= '"&fin&"' " & _ 
                    " GROUP BY LEFT(AT_CCODIGO,5)                     "
              RS2.OPEN KAD,CNN
              RS2.MOVEFIRST%> 
               <td align="RIGHT"><%=rs2("TOTA") %></td> 
               <%RS2.CLOSE %>
        </tr>
        <% do while not rs.eof and trim(cod) = trim(rs("cod")) %>
        <%j=0%>
        <tr class="Estilo0">
            <td colspan="2"></td>
            <td  align="left"><%=rs("color") %></td> 
                <%col = rs("color") %>
                <%do while not rs.eof and trim(cod) = trim(rs("cod")) and trim(col) = trim(rs("color"))  %>
                    <%for i=j to ubound(aTal) %>  
                        <%if rtrim(ltrim(rs("talla"))) <> rtrim(ltrim(aTal(i))) then %>
                            <td>&nbsp;</td>
                            <% j= i+1%>
                        <%else%>
                            <td><%=rs("oct_feb")%></td>
                            <% j= j+1%>
                            <%exit for%>
                        <%end if%>
                    <%next %>
                    <%rs.movenext %>
                    <%IF RS.EOF THEN EXIT DO %>
                <%loop%>
        
        <%IF RS.EOF THEN EXIT DO %>
        <%loop%>
    <%IF RS.EOF THEN EXIT DO %>
    </tr>
<%loop%>
</table>
</center>

<script type="text/jscript" language="jscript">
    top.window.parent.window.calcHeight()
</script>
</body>
</html>

