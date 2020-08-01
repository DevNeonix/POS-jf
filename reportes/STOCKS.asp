<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%Session.LCID=2057%>
<% tienda = Request.Cookies("tienda")("pos") %>
<% destda = Request.Cookies("tienda")("tda") %>
<!--#include file="../includes/Cnn.inc"-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=titulo%></title>
<link rel="stylesheet" type="text/css" href="../ventas.CSS" />
<%TDA= REQUEST.QueryString("TDA") %>
</head>

<body topmargin="0" leftmargin="0" rightmargin="0" border="0">
<%IF  request.QueryString("EXCEL") = "1" THEN
  archivo = "c:\temp\stkexcel.xls"
    Response.Charset = "UTF-8"
    Response.ContentType = "application/vnd.ms-excel" 
    Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo 
END IF
IF LEN(LTRIM(RTRIM(TDA)))=0 THEN %>

        <table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="2"  cellspacing="4"  border="0">
	        <% CAD =   " SELECT  'TT' AS CODIGO, ' TODAS' AS DESCRIPCION UNION SELECT CODIGO, DESCRIPCION FROM TIENDAS WHERE ESTADO ='A'  order by descripcion "
                  '  response.write(cad)
                  '  response.write("<br>")
            RS.OPEN CAD,CNN
            IF rs.recordcount > 0 THEN rs.movefirst%>
    
            <tr valign="middle" >
                 <td class="Estilo11" valign="middle" align="right" rowspan="2">
                    <label for="Radio">Excel:&nbsp;</label></td> 
                    <td><input id="excel" type="checkbox" checked name="excel" /></td>
                    <td class="Estilo11" valign="middle" align="right" rowspan="2">
                        <label >2019:&nbsp;</label></td> 
                    <td><input id="chk19" type="checkbox"  name="chk19" /></td>
    	        <td class="Estilo11" valign="middle" align="right" rowspan="2">
                    <label for="Radio">Tiendas:&nbsp;</label></td> 
                <td  class="Estilo12" align="left"  rowspan="2">
                    <select  name="TT" id="TT">
                       
                        <%do while not rs.eof %>
                            <option value="<%=TRIM(RS("CODIGO"))%>"><%=TRIM(RS("DESCRIPCION")) %></option>
                            <%rs.movenext %>
                        <%loop %>
                        <%rs.close %>
                    </select>
                </td>
                   <td class="Estilo11" valign="middle" align="right">Temporada:&nbsp;</td> 
        <td  class="Estilo12" align="left">
                <select  name="tempo" id="tempo">
                    <option value = "" selected>TODAS</option>
                    <%CAD =   " SELECT * FROM temporadas WHERE ESTADO ='A' order by descripcion "
              '  response.write(cad)
              '  response.write("<br>")
                RS.OPEN CAD,CNN
                    IF rs.recordcount > 0 THEN rs.movefirst
                    do while not rs.eof %>
                        <option value=" <%=TRIM(RS("CODIGO"))%>"><%=TRIM(RS("DESCRIPCION")) %></option>
                        <%rs.movenext %>
                    <%loop %>
                
                </select>
            </td>
            </td>
            <td><img src="../images/ok.gif" onClick="reemplaza()" style="cursor:pointer;"/></td>       
        </tr>
        </tr>
        </table>
        <%RS.CLOSE %>

<% END IF%>



<%IF LEN(LTRIM(RTRIM(TDA)))>0 THEN %>
<%      IF TRIM(TDA) = "TT" THEN 
            destda = "TODAS"
        ELSE
            rs.open "select CODIGO, descripcion from tiendas where codigo = '"&tda&"' AND ESTADO ='A'  " , cnn
            destda = rs("descripcion")
            rs.close
        END IF
tem = request.QueryString("tem")
%>

<table width="100%">
	<tr><td  align="center" class="estilo6">Listado de Stock Artículos --> tienda: <%=ucase(destda) %></td></tr>
	<tr><td><hr /></td></tr>
</table>
<%
chk19 = request.querystring("chk19")
tbl = "ARTICULOS"
if chk19 <> "" then
    tbl = "STOCK2019"
end if

CAD =	" SELECT a.CODIGO, RSFACCAR..AL0012ARTI.AR_CDESCRI AS DESCRI,               " & _
        " SUM(a.STOCK) AS STOCK, MIN(a.MINIMO) AS MINIMO                    " & _
        " FROM RSFACCAR..AL0012ARTI INNER JOIN "&tbl&" as a ON RSFACCAR..AL0012ARTI.AR_CCODIGO " & _
        "  COLLATE Modern_Spanish_CI_AI = a.CODIGO                              " & _
        " WHERE    (RSFACCAR..AL0012ARTI.AR_CDESCRI like '%"&tem&"%'  )                     "
         IF TRIM(TDA) <> "TT" THEN CAD =  CAD + " AND (LTRIM(RTRIM(a.TIENDA)) = '"&tda&"' )  " 

CAD = CAD + " GROUP BY a.CODIGO, RSFACCAR..AL0012ARTI.AR_CDESCRI                             " & _
            " order by codigo                                                                        "	
RESPONSE.WRITE(tbl)
'response.end
RS.OPEN CAD ,Cnn
		
If rs.eof or rs.bof then
	Response.Write("Stock OK, no se requiere Reposición")		
	Response.End
end if	
rs.movefirst
%>
<table id="TABLA" align="center"  cellpadding="2" cellspacing="2" bordercolor='<%=application("color2") %>' border="0" >
<%'**************************%>
<%'LINEA DE CABECERA STANDAR %>
<%'**************************%>
<%cont = 1%>
<tr>
    <td align="center" class="Estilo8">IT</td>
	<td align="center" class="Estilo8">Codigo</td>
    <td align="center" class="Estilo8">Descripcion</td>
    <td align="center" class="Estilo8">Cant.</td>
    <td align="center" class="Estilo8">Min.</td>
</tr>
<%IF NOT RS.EOF THEN%>
    <%DO WHILE NOT RS.EOF%>
        <tr  bgcolor="<% if CONT mod 2  = 0 THEN 
                response.write(Application("color2"))
                else
	            response.write(Application("color1"))
	            end IF%>" id="fila<%=Trim(Cstr(cont))%>" 
            class="<%if cdbl(rs("stock")) < cdbl(rs("minimo")) then 
                    response.write("Estilo23") 
                 else 
                    response.write("estilo5") 
                 end if %>">
            <td  align="center"><%=Cont%></td>
	        <td  align="center"><%=rs("codigo") %></td>
            <td align="left"><%=ucase(rs("descri")) %></td>
            <td align="center"><%=rs("stock") %></td>
            <td  align="center"><%=rs("minimo") %></td>
            <%cont =cont +1 %>
        </tr>
        <%rs.movenext %>
    <%loop%>
<%end if %>



</table>

<%end if %>
</body>
<script language="jscript" type="text/jscript">
    function reemplaza() {
        cad = 'stocks.asp?tda=' + document.all.TT.value + '&tem=' + (document.all.tempo.value)
        if (document.all.excel.checked == true)
            cad += '&excel=1'
        if (document.all.chk19.checked == true)
            cad += '&chk19=1'
       // alert(cad)
       //     window.open(cad)
        window.location.replace(cad)
    }
</script>
</html>
