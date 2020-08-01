<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = true %>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos") %>

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
<!--#include file="../comun/comunQRY.asp"-->
<script language="jscript" type="text/jscript">
function AGRANDA() {
    top.parent.window.document.getElementById('body0').height = 480
}
var oldrow = 1

function dd2(ff) {	// LLENA TEXTBOX ADICIONALES AL COMUN
    // LOS DEL COMUN SON CODIGO Y DESCRIPCION
    var t = document.all.TABLA;
    var pos = parseInt(ff);
    dd(ff, 0);
}
//alert()
</script>
<%
TDA = request.QueryString("TDA")
ini = request.QueryString("ini")
fin = request.QueryString("fin")
IF  request.QueryString("EXCEL") = "1" THEN
  archivo = "cTOexcel.xls"
    Response.Charset = "UTF-8"
    Response.ContentType = "application/vnd.ms-excel" 
    Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo 
END IF
'RESPONSE.WRITE(ini)
'RESPONSE.WRITE("<br>")
'RESPONSE.WRITE(fin)
'RESPONSE.WRITE("<br>")
'
'response.End

'*********************************************************************************************
CAD =   " SET DATEFORMAT DMY;                                           " & _
        " SELECT CODGRU,DESCRI,  SUM (CANT) AS CANT, COSTO,             " & _
        " SUM(CTO_TOT) as cto_tot , sum(TOT_DOC) AS TOT_DOC, SUM(DIF)   " & _
        " AS DIF                                                        " & _
        " FROM TMPCTO                                                   " & _
        " WHERE FECHA between '"&INI&"' AND '"&FIN&"'  +' 23:59:59.999'   "

IF LTRIM(RTRIM(TDA)) <> "TT" THEN     CAD = CAD + " AND TIENDA = '"&TDA&"'"
CAD = CAD + " GROUP BY CODGRU, DESCRI, COSTO ORDER BY 1 "
'*********************************************************************************************
'RESPONSE.WRITE(cAD)
'response.end
rs.open cad,cnn
'RESPONSE.WRITE(cAD)

if rs.recordcount <=0 then RESPONSE.End

IF  request.QueryString("EXCEL") = "1" THEN
  archivo = "c:\temp\cajaexcel.xls"
    Response.Charset = "UTF-8"
    Response.ContentType = "application/vnd.ms-excel" 
    Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo 
END IF
%>

<body onload="AGRANDA()">
<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="Table2" name="listado"  >
<tr>
<td><input type="button" value="Excel " onclick="REPORTE(1)" /></td>
</tr>
</table>

<center>

<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="TABLA" name="TABLA"  >
	<tr> 
    <%FOR I=0 TO RS.FIELDS.COUNT-1 %>
        <td align="center" class="Estilo8"><%=RS.FIELDS(I).NAME %></td>
    <%NEXT %>
    <td align="center" class="Estilo8">&nbsp;%&nbsp;</td>
	</tr >
<%CONT = 1 
tota =0 %>
<%do while not rs.eof %>
	<tr bgcolor="<% if CONT mod 2  = 0 THEN 
                response.write(Application("color1"))
                else
	            response.write(Application("color2"))
	            end IF%>"
	            onclick="dd('<%=(cont)%>',0)" id="fila<%=Trim(Cstr(cont))%>" >
		<%for i =0 to 1 %>
            <td align="LEFT" CLASS="EstiloT">&nbsp;<%=trim(RS.FIELDS.ITEM(i))%>&nbsp;</td>
		<%next %>
         <%if isnull(RS.FIELDS.ITEM(2)) then nume = 0 else nume = cdbl(RS.FIELDS.ITEM(2))
             tota = tota + nume%>
        <%for i =2 to RS.FIELDS.COUNT -1 %>
            <%if isnull(RS.FIELDS.ITEM(i)) then nume = 0 else nume = cdbl(RS.FIELDS.ITEM(i))
%>
            <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(nume,2,,,TRUE)%></td>
		<%next %>
        <%if cdbl(RS.FIELDS.ITEM(5)) = 0 then up=1 else up = cdbl(RS.FIELDS.ITEM(6))
        if cdbl(RS.FIELDS.ITEM(4)) = 0 then bot=1 else bot = cdbl(RS.FIELDS.ITEM(5))
        %>




        <% 
        if bot = 0 then bot=-1
        %>





         <td align="right" CLASS="EstiloT">&nbsp;<%=formatnumber(((up/bot))*100,2,,,true)%>&nbsp;%</td>
	</tr> 
    <%CONT =CONT + 1 %>
    <%rs.movenext%>
<%loop %>
<tr> 

        <td align="center" class="Estilo8" colspan="2">Total Unidades</td>
            <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(tota,2,,,TRUE)%></td>
	</tr >

</table>

<iframe  width="100%" src="" id="body0" name="body0" scrolling="yes" frameborder="1" height="40" align="middle" style="display:none" ></iframe>


<script language="jscript" type="text/jscript">
rec = parseInt('<%=rs.recordcount%>',10)
if (rec > 0 )
dd2('1');


</script>
</center>

<script language="jscript" type="text/jscript">
    function REPORTE(op) {
        if (op == '1')
            window.open('CTODETA.asp?TDA=' + '<%=TDA%>' + '&INI=' + '<%=INI%>' + '&FIN=' + '<%=FIN%>'  + '&EXCEL=1')
    }
</script>
</body>
</html>
