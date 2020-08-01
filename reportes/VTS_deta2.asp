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
art = request.QueryString("art")
cf  = request.QueryString("cf")





xtda =""
if len(tda) > 2 then
    atda = split(tda,",")
    jatda=""
    for i=0 to ubound(atda)
        jatda = jatda&"'"&atda(i)&"'"
    next
    xtda = replace(jatda,"''","','")
   '' response.write(xtda&"<br>")
else
    xtda = tda
end if


xtempo = ""
atem = split(tem,",")
jatem=""
for i=0 to ubound(atem)
jatem = jatem&"'%"&atem(i)&"%'"
next
xtempo = replace(jatem,"''","','")

kddtmpo = "("
xatem = split(xtempo,",")

for ll=0 to ubound(xatem)
kddtmpo = kddtmpo & "descri like "&xatem(ll)&" or "
next
if len(kddtmpo) > 3 then
kddtmpo = left( kddtmpo ,  len(kddtmpo) - 3  ) &")"
else
kddtmpo = tem
end if

'response.write(kddtmpo&"<br>")





IF  request.QueryString("EXCEL") = "1" THEN
  archivo = "c:\temp\vtart_excel.xls"
    Response.Charset = "UTF-8"
    Response.ContentType = "application/vnd.ms-excel" 
    Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo 
END IF







'*********************************************************************************************
CAD =   " SET DATEFORMAT DMY;                                                            " & _
        "SELECT TIENDA, UPPER(RTRIM(LTRIM(CODART))) AS SKU, LEFT(CODART, 5) AS CODIGO,RIGHT(LTRIM(RTRIM(CODART)), 2) AS COLOR, t2.tg_cdescri as COLOR_DESC,SUBSTRING(LEFT(CODART, 8), 6, 3) AS TALLA, t3.tg_cdescri AS TALLA_DESC,DESCRI, SALE, LISTA1 * 1.18 AS UNIT, ISNULL(pvp, 0) - ISNULL(IGV,0) AS PVP, ISNULL(PORDES,0) AS PORDES, ISNULL(DCT,0) AS DCT, ISNULL(IGV,0) AS IGV, (ISNULL(pvp, 0) - ISNULL(IGV,0)) - ISNULL(DCT,0) + ISNULL(IGV,0) AS TOT" & _
        " FROM VIEW_VENTAS_ARTICULO_2 as t1 left outer join rsfaccar.dbo.AL0001TABL as t2 on (t2.TG_CCLAVE = RIGHT(LTRIM(RTRIM(t1.CODART)), 2) ) and t2.TG_CCOD = 'D2'   left outer join rsfaccar.dbo.AL0001TABL as t3 on (t3.TG_CCLAVE = sUBSTRING(LEFT(CODART, 8), 6, 3)) and t3.TG_CCOD = 'D1'  WHERE SALE > 0 and pvp > 0                " & _
        " AND FECHA between '"&INI&"' AND '"&FIN&"'  +' 23:59:59.999'                   "
IF LTRIM(RTRIM(TDA)) <> "TT" THEN   CAD = CAD + " AND TIENDA in ("&xtda&")   "
IF LTRIM(RTRIM(art)) <>   ""   THEN   CAD = CAD + " AND codart = '"&art&"'  "
IF LTRIM(RTRIM(TEM)) <>   ""   THEN   CAD = CAD + " and "&kddtmpo&"   "
'CAD = CAD + " GROUP BY CODART, DESCRI ORDER BY CODART                                   "
'*********************************************************************************************
if cf = "true" then

    CAD =   " SET DATEFORMAT DMY;                                                            " & _
            " SELECT tienda,UPPER(RTRIM(LTRIM(CODART))) AS SKU, LEFT(CODART, 5) AS CODIGO,YEAR(FECHA) AS ANO,MONTH(FECHA) AS MES,LEFT(CODART, 5) AS CODIGO,RIGHT(LTRIM(RTRIM(CODART)), 2) AS COLOR, t2.tg_cdescri as COLOR_DESC,SUBSTRING(LEFT(CODART, 8), 6, 3) AS TALLA, t3.tg_cdescri AS TALLA_DESC,DESCRI, SALE, LISTA1 * 1.18 AS UNIT, ISNULL(pvp, 0) - ISNULL(IGV,0) AS PVP, ISNULL(PORDES,0) AS PORDES, ISNULL(DCT,0) AS DCT, ISNULL(IGV,0) AS IGV, (ISNULL(pvp, 0) - ISNULL(IGV,0)) - ISNULL(DCT,0) + ISNULL(IGV,0) AS TOT" & _
            " FROM VIEW_VENTAS_ARTICULO_2 as t1 left outer join rsfaccar.dbo.AL0001TABL as t2 on (t2.TG_CCLAVE = RIGHT(LTRIM(RTRIM(t1.CODART)), 2) ) and t2.TG_CCOD = 'D2'  left outer join rsfaccar.dbo.AL0001TABL as t3 on (t3.TG_CCLAVE = sUBSTRING(LEFT(CODART, 8), 6, 3)) and t3.TG_CCOD = 'D1' WHERE SALE > 0 and pvp > 0            " & _
            " AND FECHA between '"&INI&"' AND '"&FIN&"'  +' 23:59:59.999'                   "
    IF LTRIM(RTRIM(TDA)) <> "TT" THEN   CAD = CAD + " AND TIENDA in ("&xtda&")"
    IF LTRIM(RTRIM(art)) <>   ""   THEN   CAD = CAD + " AND codart = '"&art&"'  "
    IF LTRIM(RTRIM(TEM)) <>   ""   THEN   CAD = CAD + " and "&kddtmpo&"    "
    'CAD = CAD + " GROUP BY CODART, DESCRI ORDER BY CODART                                   "
    '*********************************************************************************************

end if 
cad = cad & " order by 2"


'RESPONSE.WRITE(CAD)
'response.end
rs.open cad,cnn
if rs.recordcount <=0 then 
    RESPONSE.WRITE("<center>")
    RESPONSE.WRITE("<font color='magenta'>")
    RESPONSE.WRITE("No hay registros que cumplan con su criterio")
    RESPONSE.End
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
	</tr >
<%CONT = 1
tota =0
totaL =0  %>
<%if cf = "false" then %>
    <%do while not rs.eof %>
        <tr bgcolor="<% if CONT mod 2  = 0 THEN 
                    response.write(Application("color1"))
                    else
                    response.write(Application("color2"))
                    end IF%>"
                    onclick="dd('<%=(cont)%>',0)" id="fila<%=Trim(Cstr(cont))%>" >
            <%for i =0 to 7 %>
                <td align="LEFT" CLASS="EstiloT">&nbsp;<%=trim(RS.FIELDS.ITEM(i))%></td>
            <%next %>
            <td align="LEFT" CLASS="EstiloT"><%=trim(RS.FIELDS.ITEM(8))%></td>
            <td align="LEFT" CLASS="EstiloT"><%=trim(RS.FIELDS.ITEM(9))%></td>
            <td align="LEFT" CLASS="EstiloT"><%=trim(RS.FIELDS.ITEM(10))%></td>
            <%if isnull(RS.FIELDS.ITEM(11)) then nume = 0 else nume = cdbl(RS.FIELDS.ITEM(8))
                tota = tota + nume%>
            <%for i =11 to RS.FIELDS.COUNT -1 %>
                <%if isnull(RS.FIELDS.ITEM(i)) then nume = 0 else nume = cdbl(RS.FIELDS.ITEM(i)) %>
                <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(nume,2,,,TRUE)%></td>           
            <%next %>
            <%TOTAL =  TOTAL + cdbl(RS("TOT")) %>
        </tr> 
        
        <%CONT =CONT + 1 %>
        <%rs.movenext%>
    <%loop %>
<%else%>
    <%do while not rs.eof %>
    <tr bgcolor="<% if CONT mod 2  = 0 THEN 
                response.write(Application("color1"))
                else
                response.write(Application("color2"))
                end IF%>"
                onclick="dd('<%=(cont)%>',0)" id="fila<%=Trim(Cstr(cont))%>" >
        <%for i =0 to 7 %>
                <td align="LEFT" CLASS="EstiloT">&nbsp;<%=trim(RS.FIELDS.ITEM(i))%></td>
            <%next %>
            <td align="LEFT" CLASS="EstiloT"><%=trim(RS.FIELDS.ITEM(8))%></td>
            <td align="LEFT" CLASS="EstiloT"><%=trim(RS.FIELDS.ITEM(9))%></td>
            <td align="LEFT" CLASS="EstiloT"><%=trim(RS.FIELDS.ITEM(10))%></td>
        <%if isnull(RS.FIELDS.ITEM(11)) then nume = 0 else nume = cdbl(RS.FIELDS.ITEM(11))
            tota = tota + nume%>
         <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(nume,0,,,TRUE)%></td>   

        <%for i =12 to RS.FIELDS.COUNT -1 %>
            <%if isnull(RS.FIELDS.ITEM(i)) then nume = 0 else nume = cdbl(RS.FIELDS.ITEM(i)) %>
            <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(nume,2,,,TRUE)%></td>           
        <%next %>
        <%TOTAL =  TOTAL + cdbl(RS("TOT")) %>
    </tr> 

    <%CONT =CONT + 1 %>
    <%rs.movenext%>
    <%loop %>

<%end if%>
<%if cf = "false" then%>
<tr>
    <td align="center" class="Estilo8" colspan="7">Total Unidades</td>
    <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(tota,2,,,TRUE)%></td>
    <td colspan="5"></td>
    <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(totaL,2,,,TRUE)%></td>
</tr >
<%else%>
<tr>
    <td align="center" class="Estilo8" colspan="11">Total Unidades</td>
    <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(tota,2,,,TRUE)%></td>
    <td colspan="5"></td>
    <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(totaL,2,,,TRUE)%></td>
</tr >
<%end if%>
</table>

<iframe  width="100%" src="" id="body0" name="body0" scrolling="yes" frameborder="1" height="40" align="middle" style="display:none" ></iframe>


<script language="jscript" type="text/jscript">
rec = parseInt('<%=rs.recordcount%>',10)
if (rec > 0 )
dd2('1');


</script>
</center>

<script >
    
    function REPORTE(op) {
        window.location.replace('VTS_deta2.asp?tda=' + '<%=tda %>' + '&ini=' + '<%=ini %>' + '&fin=' + '<%=fin%>' + '&excel=1' + '&art='+ '<%=art%>' + '&tem='+'<%=tem %>' + '&cf='+ '<%=cf%>')
    }
</script>
</body>
</html>
