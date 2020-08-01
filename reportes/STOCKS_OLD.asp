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
	        <% CAD =   " SELECT * FROM TIENDAS WHERE ESTADO ='A' order by descripcion "
                  '  response.write(cad)
                  '  response.write("<br>")
            RS.OPEN CAD,CNN
            IF rs.recordcount > 0 THEN rs.movefirst%>
    
            <tr valign="middle" >
                 <td class="Estilo11" valign="middle" align="right" rowspan="2">
                    <label for="Radio">Excel:&nbsp;</label></td> 
                    <td><input id="excel" type="checkbox" checked name="excel" /></td>
    	        <td class="Estilo11" valign="middle" align="right" rowspan="2">
                    <label for="Radio">Tiendas:&nbsp;</label></td> 
                <td  class="Estilo12" align="left"  rowspan="2">
                    <select  name="TT" id="TT">
                        <option value = "" selected></option>
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
<% rs.open "select *  from tiendas where codigo = '"&tda&"'  " , cnn
destda = rs("descripcion")
rs.close
tem = request.QueryString("tem")
%>

<table width="100%">
	<tr><td  align="center" class="estilo6">Listado de Stock Artículos --> tienda: <%=ucase(destda) %></td></tr>
	<tr><td><hr /></td></tr>
</table>
<%
CAD =	" SELECT dbo.ARTICULOS.TIENDA, dbo.ARTICULOS.CODIGO, RSFACCAR.dbo.AL0012ARTI.AR_CDESCRI AS DESCRI, RSFACCAR.dbo.AL0012ARTI.AR_CFAMILI AS CODM,  " & _
        " TB.TG_CDESCRI AS MARCA, RSFACCAR.dbo.AL0012ARTI.AR_CGRUPO AS CLIN, TR.TG_CDESCRI AS LINEA, dbo.ARTICULOS.STOCK, dbo.ARTICULOS.MINIMO,         " & _
        " RSFACCAR.dbo.AL0012ARTI.AR_CUNIDAD, RSFACCAR.dbo.AL0012ARTI.AR_CMONVTA, RSFACCAR.dbo.AL0012ARTI.AR_NIGVPOR, dbo.ARTICULOS.LISTA1,             " & _
        " dbo.ARTICULOS.LISTA2, dbo.ARTICULOS.LISTA3, dbo.ARTICULOS.USUARIO, dbo.ARTICULOS.FECHA, dbo.ARTICULOS.ESTADO, dbo.ARTICULOS.PLANILLA,         " & _
        " RSFACCAR.dbo.AL0012ARTI.AR_CFAMILI AS Expr1, RSFACCAR.dbo.AL0012ARTI.AR_CCUENTA, RSFACCAR.dbo.FT0012CTAE.TC_CVENTAS,                          " & _
        " RSFACCAR.dbo.FT0012CTAE.TC_CCOSVEN                                                                                                            " & _
        " FROM RSFACCAR.dbo.AL0012ARTI INNER JOIN                                                                                                       " & _
        " RSFACCAR.dbo.AL0012TABL AS TB ON RSFACCAR.dbo.AL0012ARTI.AR_CFAMILI = TB.TG_CCLAVE AND TB.TG_CCOD = '38' INNER JOIN                           " & _
        " RSFACCAR.dbo.AL0012TABL AS TR ON RSFACCAR.dbo.AL0012ARTI.AR_CGRUPO = TR.TG_CCLAVE AND TR.TG_CCOD = '06' INNER JOIN                            " & _
        " RSFACCAR.dbo.FT0012CTAE ON RSFACCAR.dbo.AL0012ARTI.AR_CCUENTA = RSFACCAR.dbo.FT0012CTAE.TC_CEXISTE RIGHT OUTER JOIN                           " & _
        " dbo.ARTICULOS ON RSFACCAR.dbo.AL0012ARTI.AR_CCODIGO COLLATE Modern_Spanish_CI_AI = dbo.ARTICULOS.CODIGO                                       " & _
        " WHERE        (LTRIM(RTRIM(dbo.ARTICULOS.TIENDA)) = '"&tda&"'   ) and (RSFACCAR.dbo.AL0012ARTI.AR_CDESCRI like '%"&tem&"%'  )                                    " & _
		" order by codigo                                                                                                                               "	
'RESPONSE.WRITE(CAD)
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
       // alert(cad)
       //     window.open(cad)
        window.location.replace(cad)
    }
</script>
</html>
