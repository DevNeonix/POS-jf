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

IF  request.QueryString("EXCEL") = "1" THEN
  archivo = "c:\temp\stkexcel.xls"
    Response.Charset = "UTF-8"
    Response.ContentType = "application/vnd.ms-excel" 
    Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo 
END IF

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
kddtmpo = kddtmpo & "(select ar_cdescri from rsfaccar..al0012arti where AR_CCODIGO = dd.codart) like "&xatem(ll)&" or "
next
'response.write(kddtmpo&"<br>")
if len(kddtmpo) > 3 then
kddtmpo = left( kddtmpo ,  len(kddtmpo) - 3  ) &")"
else
kddtmpo = tem
end if


'response.write(kddtmpo&"<br>")
' " isnull(MAX(LISTA1*1.18),0) AS UNIT,                                              " & _
'*********************************************************************************************
    CAD =   " SET DATEFORMAT DMY                                                               " & _ 
            " select codart, (select ar_cdescri from rsfaccar..al0012arti                      " & _
            " where AR_CCODIGO = dd.codart) as descri,                                         " & _
            " isnull(sum(sale),0) as cant,                                                     " & _
            " isnull(MAX(precio/sale),0) AS UNIT,                                              " & _
            " isnull(SUM(LISTA1*SALE),0) AS pvp,                                               " & _
            " isnull(AVG(PORDES),0) AS PORDES,                                                 " & _
            " isnull(SUM(dd.descuento),0) AS DCT,                                              " & _
            " isnull(SUM(dd.IGV),0) AS IGV,                                                    " & _
            " isnull(SUM(PRECIO),0) AS TOT                                                     " & _
            " from view_movimcab_all as cc inner join view_movimdet_all as dd on cc.operacion = dd.operacion     " & _
            " full outer join ARTICULOS as aa on dd.CODART = aa.codigo and aa.TIENDA = cc.TIENDA    " & _
        " where sale > 0 and pvp>0 AND CONVERT(SMALLDATETIME,FECdoc,113) between '"&INI&"' AND " & _
        " '"&FIN&"'  +' 23:59:59.999'                                                          "

IF UCASE(TRIM(RTRIM(TDA))) <> "TT"    THEN   CAD = CAD + " AND cc.TIENDA in ("&xtda&")  "
IF UCASE(TRIM(RTRIM(art))) <>   ""    THEN   CAD = CAD + " AND ltrim(rtrim(codart)) = '"&art&"'  "
IF UCASE(TRIM(RTRIM(TEM))) <>   ""    THEN   CAD = CAD + " and  "&kddtmpo

' mm le quite la tienda porque cecilia lo pidio 09-10-2019
'CAD = CAD + " GROUP BY tienda,ltrim(rtrim(CODART)), DESCRI ORDER BY 1,2                           "
CAD = CAD + " GROUP BY CODART ORDER BY 1,2                                   "
if cf = "true" then

    CAD =   " SET DATEFORMAT DMY                                                               " & _ 
            " select codart, (select ar_cdescri from rsfaccar..al0012arti                      " & _
            " where AR_CCODIGO = dd.codart) as descri,                                         " & _
            " isnull(sum(sale),0) as cant,                                                     " & _
            " isnull(MAX(LISTA1*1.18),0) AS UNIT,                                              " & _
            " isnull(SUM(LISTA1*SALE),0) AS pvp,                                               " & _
            " isnull(AVG(PORDES),0) AS PORDES,                                                 " & _
            " isnull(SUM(dd.descuento),0) AS DCT,                                              " & _
            " isnull(SUM(dd.IGV),0) AS IGV,                                                    " & _
            " isnull(SUM(PRECIO),0) AS TOT                                                " & _
            " from  view_movimcab_all as cc                                                                  " & _
            " inner join view_movimdet_all as dd on cc.operacion =  dd.operacion                             " & _
            " full outer join ARTICULOS as aa on  dd.codart = aa.CODIGO  and aa.TIENDA = cc.TIENDA       " & _
            " where sale > 0 and pvp > 0 AND CONVERT(SMALLDATETIME,FECdoc,113) between '"&INI&"' AND" & _
            " '"&FIN&"'  +' 23:59:59.999'                                                           "
    IF LTRIM(RTRIM(TDA)) <>  "TT"  THEN   CAD = CAD + " AND cc.TIENDA in ("&xtda&")  "
    IF LTRIM(RTRIM(art)) <>   ""   THEN   CAD = CAD + " AND codart = '"&art&"'  "
    IF LTRIM(RTRIM(TEM)) <>   ""   THEN   CAD = CAD + " and "&kddtmpo
    ' mm le quite la tienda porque cecilia lo pidio 09-10-2019
    'CAD = CAD + " GROUP BY TIENDA,FECHA,ltrim(rtrim(CODART)) , DESCRI ORDER BY 3,1,2  "
    CAD = CAD + " GROUP BY cc.FECHA,CODART  ORDER BY 3,1,2  "

end if


'*********************************************************************************************
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
        <td align="center" class="Estilo8">sku</td>
    <%FOR I=0 TO RS.FIELDS.COUNT-1 %>
        <td align="center" class="Estilo8"><%=ucase(RS.FIELDS(I).NAME) %></td>
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
                <td align="LEFT" CLASS="EstiloT">&nbsp;<%=trim(left(RS("codart"),5))%></td>
                <td align="LEFT" CLASS="EstiloT"><%=trim(RS("codart"))%></td>
                <td align="LEFT" CLASS="EstiloT"><%=trim(RS("descri"))%> </td>
               
            
            <%if isnull(RS("cant")) then nume = 0 else nume = cdbl(RS("cant"))
              if cdbl(RS("pordes")) = 0 then porc = " " else porc = cstr(formatnumber(cdbl(RS("pordes")),0,,,true)) + " % "
                tota = tota + nume%>
                <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(cdbl(RS("cant")),0,,,TRUE)%></td>   
                <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(cdbl(rs("unit")),2,,,TRUE)%></td>
                <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(cdbl(rs("pvp")),2,,,TRUE)%></td>   
                <td align="RIGHT" CLASS="EstiloT"><%=porc%> &nbsp; </td>
                <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(cdbl(rs("dct")),2,,,TRUE)%></td>   
                <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(cdbl(rs("igv")),2,,,TRUE)%></td>            
                <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(cdbl(rs("tot")),2,,,TRUE)%></td>   
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
            <%if isnull(RS.FIELDS.ITEM(5)) then nume = 0 else nume = cdbl(RS.FIELDS.ITEM(4))
                tota = tota + nume%>
            <%for i =4 to RS.FIELDS.COUNT -1 %>
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
    <td align="center" class="Estilo8" colspan="4">Total Unidades</td>
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
            window.open('VTS_deta.asp?tda=' + '<%=tda %>' + '&ini=' + '<%=ini %>' + '&fin=' + '<%=fin%>'+'&tem='+'<%=tem%>'+'&cf='+<%=cf%> + '&excel=1')
    }
</script>
</body>
</html>
