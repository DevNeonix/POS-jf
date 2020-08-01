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
</script>
<%
TDA = request.QueryString("TDA")
ini = request.QueryString("ini")
fin = request.QueryString("fin")
tem = request.QueryString("tem")

'*********************************************************************************************
CAD =   " SET DATEFORMAT DMY                                                            " & _
        " select  CODART, DESCRI, SUM (SALE) AS CANT, SUM(PRECIO)/SUM(SALE) AS UNIT,    " & _
        " SUM(PVP-IGV) AS PVP, AVG(PORDES) AS PORDES, SUM(DCT) AS DCT, SUM(IGV) AS IGV, " & _
        " SUM(PVP) AS TOT                                                               " & _
        " FROM VIEW_VENTAS_ARTICULO WHERE SALE > 0  and isnull(pvp,0) > 0               " & _
        " AND FECHA between '"&INI&"' AND '"&FIN&"'+' 23:59:59.999'                     " & _
        " and descri like '%"&tem&"%'                                                   "
IF LTRIM(RTRIM(TDA)) <> "TT" THEN     CAD = CAD + " AND TIENDA = '"&TDA&"'"

CAD = CAD + " GROUP BY CODART, DESCRI ORDER BY CODART "
'*********************************************************************************************
'RESPONSE.WRITE(cAD)
'response.end
rs.open cad,cnn
if rs.recordcount <=0 then RESPONSE.End
 archivo = "c:\temp\VTS_DETA_excel.xls"
    Response.Charset = "UTF-8"
    Response.ContentType = "application/vnd.ms-excel" 
    Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo 
%>

<body onload="AGRANDA()"  text="black">
<center>

<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="TABLA" name="TABLA"  >
	<tr> 
    <%FOR I=0 TO RS.FIELDS.COUNT-1 %>
        <td align="center" class="Estilo8"><%=RS.FIELDS(I).NAME %></td>
    <%NEXT %>
	</tr >
<%CONT = 1 %>
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
        <%for i =2 to RS.FIELDS.COUNT -1 %>
            <%if isnull(RS.FIELDS.ITEM(i)) then nume = 0 else nume = cdbl(RS.FIELDS.ITEM(i)) %>
            <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(nume,2,,,TRUE)%></td>
		<%next %>
        
	</tr> 
    <%CONT =CONT + 1 %>
    <%rs.movenext%>
<%loop %>


</table>

<iframe  width="100%" src="" id="body0" name="body0" scrolling="yes" frameborder="1" height="40" align="middle" style="display:none" ></iframe>


<script language="jscript" type="text/jscript">
rec = parseInt('<%=rs.recordcount%>',10)
if (rec > 0 )
dd2('1');


</script>
</center>
</body>
</html>
