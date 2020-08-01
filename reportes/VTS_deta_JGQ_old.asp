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
'response.write(kddtmpo&"<br>")
if len(kddtmpo) > 3 then
kddtmpo = left( kddtmpo ,  len(kddtmpo) - 3  ) &")"
else
kddtmpo = tem
end if


'response.write(kddtmpo&"<br>")

'*********************************************************************************************
CAD =   " SET DATEFORMAT DMY                                                            " & _
        " select  ltrim(rtrim(CODART)) CODART, DESCRI, SUM (SALE) AS CANT, MAX(LISTA1*1.18) AS UNIT,         " & _
        " SUM(PVP-IGV) AS PVP, AVG(PORDES) AS PORDES, SUM(DCT) AS DCT, SUM(IGV) AS IGV, " & _
        "  SUM(PVP-IGV) -  SUM(DCT) +  SUM(IGV) AS TOT                                  " & _
        " FROM VIEW_VENTAS_ARTICULO WHERE SALE > 0  and isnull(pvp,0) > 0               " & _
        " AND FECHA between '"&INI&"' AND '"&FIN&"'  +' 23:59:59.999'                   "
IF LTRIM(RTRIM(TDA)) <> "TT" THEN   CAD = CAD + " AND TIENDA in ("&xtda&")  "
IF LTRIM(RTRIM(art)) <>   ""   THEN   CAD = CAD + " AND ltrim(rtrim(codart)) = '"&art&"'  "
IF LTRIM(RTRIM(TEM)) <>   ""   THEN   CAD = CAD + "and  "&kddtmpo
CAD = CAD + " GROUP BY ltrim(rtrim(CODART)), DESCRI ORDER BY 1,2                                   "

if cf = "true" then

CAD =   " SET DATEFORMAT DMY                                                            " & _
        " select  TIENDA,YEAR(FECHA) AS ANO,MONTH(FECHA) AS MES,ltrim(rtrim(CODART)) CODART, DESCRI, SUM (SALE) AS CANT, MAX(LISTA1*1.18) AS UNIT,         " & _
        " SUM(PVP-IGV) AS PVP, AVG(PORDES) AS PORDES, SUM(DCT) AS DCT, SUM(IGV) AS IGV, " & _
        "  SUM(PVP-IGV) -  SUM(DCT) +  SUM(IGV) AS TOT                                  " & _
        " FROM VIEW_VENTAS_ARTICULO WHERE SALE > 0  and isnull(pvp,0) > 0               " & _
        " AND FECHA between '"&INI&"' AND '"&FIN&"'  +' 23:59:59.999'                   "
IF LTRIM(RTRIM(TDA)) <> "TT" THEN   CAD = CAD + " AND TIENDA in ("&xtda&")  "
IF LTRIM(RTRIM(art)) <>   ""   THEN   CAD = CAD + " AND codart = '"&art&"'  "
IF LTRIM(RTRIM(TEM)) <>   ""   THEN   CAD = CAD + "and "&kddtmpo
CAD = CAD + " GROUP BY TIENDA,FECHA,ltrim(rtrim(CODART)) , DESCRI ORDER BY 3,1,2  "


end if


'*********************************************************************************************
RESPONSE.WRITE(CAD)
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
<% if cf="false" then%>

    <%do while not rs.eof %>
        <tr bgcolor="<% if CONT mod 2  = 0 THEN 
                    response.write(Application("color1"))
                    else
                    response.write(Application("color2"))
                    end IF%>"
                    onclick="dd('<%=(cont)%>',0)" id="fila<%=Trim(Cstr(cont))%>" >
            <%for i =0 to 1 %>
                <td align="LEFT" CLASS="EstiloT"><%=trim(RS.FIELDS.ITEM(i))%></td>
            <%next %>
            <%if isnull(RS.FIELDS.ITEM(3)) then nume = 0 else nume = cdbl(RS.FIELDS.ITEM(3))
                tota = tota + nume%>
            <%for i =2 to RS.FIELDS.COUNT -1 %>
                <%if isnull(RS.FIELDS.ITEM(i)) then nume = 0 else nume = cdbl(RS.FIELDS.ITEM(i)) %>
                <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(nume,2,,,TRUE)%></td>           
            <%next %>
            <%TOTAL =  TOTAL + cdbl(RS("TOT")) %>
        </tr> 
        
        <%CONT =CONT + 1 %>
        <%rs.movenext%>
    <%loop %>
    <%
    else
    %>
    <%do while not rs.eof %>
        <tr bgcolor="<% if CONT mod 2  = 0 THEN 
                    response.write(Application("color1"))
                    else
                    response.write(Application("color2"))
                    end IF%>"
                    onclick="dd('<%=(cont)%>',0)" id="fila<%=Trim(Cstr(cont))%>" >
            <%for i =0 to 4 %>
                <td align="LEFT" CLASS="EstiloT"><%=trim(RS.FIELDS.ITEM(i))%></td>
            <%next %>
            <%if isnull(RS.FIELDS.ITEM(5)) then nume = 0 else nume = cdbl(RS.FIELDS.ITEM(5))
                tota = tota + nume%>
            <%for i =5 to RS.FIELDS.COUNT -1 %>
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
    <td align="center" class="Estilo8" colspan="3">Total Unidades</td>
    <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(tota,2,,,TRUE)%></td>
    <td colspan="5"></td>
    <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(totaL,2,,,TRUE)%></td>
</tr >
<%else%>
<tr>
    <td align="center" class="Estilo8" colspan="5">Total Unidades</td>
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

<script language="jscript" type="text/jscript">
    function REPORTE(op) {
        if (op == '1')
            window.open('VTS_deta_excel.asp?tda=' + '<%=tda %>' + '&ini=' + '<%=ini %>' + '&fin=' + '<%=fin%>'+'&tem='+'<%=tem%>'+'&cf='+<%=cf%>)
    }
</script>
</body>
</html>
