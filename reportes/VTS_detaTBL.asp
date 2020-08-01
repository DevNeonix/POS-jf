<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = true %>
<%Session.LCID=2057%>
<%tienda = Request.Cookies("tienda")("pos") %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<LINK REL="stylesheet" TYPE="text/css" HREF="../VENTAS.CSS">
<!--#include file="../comun/funcionescomunes.asp"-->
<!--#include file="../includes/cnn.inc"-->
<!--#include file="../comun/comunQRY.asp"-->
<script language="jscript" type="text/jscript">
function AGRANDA() {
    top.parent.window.document.getElementById('body0').height = 680
    //alert()
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
tem = TRIM(request.QueryString("tem"))

IF  request.QueryString("EXCEL") = "1" THEN
  archivo = "c:\temp\cajaexcel.xls"
    Response.Charset = "UTF-8"
    Response.ContentType = "application/vnd.ms-excel" 
    Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo 
END IF

'*********************************************************************************************
kAD =   " exec SP_LENGUAJE ; SET DATEFORMAT DMY;    "  &_
	    " select DISTINCT YEAR(VV.FECHA) AS ANIO, MONTH(VV.FECHA) AS NMES, UPPER(DATENAME(mm, VV.FECHA)) AS MES,        " & _
	    " VV.TIENDA AS CTDA, TT.DESCRIPCION AS TIENDA, UPPER(VV.CODART) AS CODIGO,                                      " & _
	    " LEFT(VV.CODART,5) AS GRUPO, UPPER(SUBSTRING ( VV.CODART ,9, 2)) AS COLOR,                                     " & _
	    " UPPER(SUBSTRING ( VV.CODART ,6, 3)) AS TALLA, 	CODM,MARCA, aa.CLIN AS CTIP, aa.LINEA AS TIP_PDA,           " & _
	    " ISNULL(CGEN,'G') AS CGEN, ISNULL(DGEN,'GENERICO') AS GENERO,                                                  " & _
	    " VV.DESCRI, SUM (SALE) AS QTY_VTA, (AA.STOCK) AS STK ,                                                         " & _
	    " COSTO, MAX(AA.LISTA1) *(1+(MAX(AR_NIGVPOR)/100)) AS UNIT,                                                     " & _
        " SUM(PVP-IGV) AS PVP, AVG(PORDES) AS PORDES, SUM(DCT) AS DCT, SUM(IGV) AS IGV,                                 " & _
        " SUM((PVP-IGV) - DCT + IGV)  AS TOT, isnull(te.descripcion, '') as DTEM                                " & _
	    " FROM VIEW_VENTAS_ARTICULO AS VV                                                                               " & _
	    " INNER JOIN TIENDAS AS TT ON  TIENDA = CODIGO                                                                  " & _
	    " INNER JOIN VIEW_ARTICULOS_TIENDA AS AA ON AA.TIENDA = VV.TIENDA AND AA.CODIGO = VV.CODART                     " & _
	    " INNER JOIN COSTOS CC ON CC.CODIGO = LEFT(CODART,5)                                                            " & _
	    " FULL OUTER JOIN VIEW_GENERO_PRENDA GG ON LEFT(VV.CODART,5)  COLLATE Modern_Spanish_CI_Ai = LEFT(GG.ART,5)                                   " & _
        " FULL OUTER JOIN TEMPORADAS AS TE ON (ltrim(rtrim(te.codigo)) collate Modern_Spanish_CI_AI =  RIGHT(RTRIM(VV.DESCRI),4) " & _
		" OR  ltrim(rtrim(te.codigo)) collate Modern_Spanish_CI_AI =  RIGHT(RTRIM(VV.DESCRI),2))                        " & _
	    " WHERE SALE > 0 and isnull(pvp,0) > 0  AND VV.FECHA between '"&INI&"' AND '"&FIN&"'+' 23:59:59.999'            " 

IF LEN(TRIM(TEM)) > 0 THEN 
 kAD = kAD +  " AND RIGHT(ltrim(RTRIM(VV.DESCRI)),4) = '"&TEM&"'    " 
end if
'
IF TRIM(tda) <> "TT" THEN  
    kAD = kAD + "  AND VV.TIENDA = '"&tda&"' "
end if

kAD = kAD +	" AND VV.TIENDA < '20'                                                                                      " & _
	        " GROUP BY VV.CODART, VV.DESCRI,  YEAR(VV.FECHA), DATENAME(mm, VV.FECHA), TT.DESCRIPCION, MONTH(VV.FECHA),  " & _
            " VV.TIENDA, AA.STOCK, CODM, aa.CLIN, MARCA, aa.LINEA,COSTO, CGEN, DGEN,RIGHT(RTRIM(VV.DESCRI),4)           " & _
	        " , te.descripcion ORDER BY YEAR(VV.FECHA), MONTH(VV.FECHA), VV.TIENDA,  UPPER(VV.CODART)                   "
'*********************************************************************************************
'RESPONSE.WRITE(kad)
'response.end


rs.open kAD, cnn


if rs.recordcount <=0 then RESPONSE.End
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
tota =0 %>
<%do while not rs.eof %>
	<tr bgcolor="<% if CONT mod 2  = 0 THEN 
                response.write(Application("color1"))
                else
	            response.write(Application("color2"))
	            end IF%>"
	            onclick="dd('<%=(cont)%>',0)" id="fila<%=Trim(Cstr(cont))%>" >
		<%for i =0 to 15 %>
            <td align="LEFT" CLASS="EstiloT"><%=trim(RS.FIELDS.ITEM(i))%></td>
		<%next %>
        
        <%for i =16 to RS.FIELDS.COUNT -2 %>
            <%if isnull(RS.FIELDS.ITEM(i)) then nume = 0 else nume = cdbl(RS.FIELDS.ITEM(i)) %>
            <td align="RIGHT" CLASS="EstiloT"><%=FORMATNUMBER(nume,2,,,TRUE)%></td>
		<%next %>
         <td align="LEFT" CLASS="EstiloT"><%=trim(RS.FIELDS.ITEM(i))%></td>
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

<script language="jscript" type="text/jscript">
    function REPORTE(op) {
        window.location.replace('VTS_detaTBL.asp?tda=' + '<%=tda %>' + '&ini=' + '<%=ini %>' + '&fin=' + '<%=fin%>' + '&excel=1')

    }
</script>
</body>
</html>
