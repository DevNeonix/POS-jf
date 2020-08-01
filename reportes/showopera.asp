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
<!--#include file="../comun/comunqry.asp"-->
<!--#include file="../includes/funcionesVBscript.asp"-->
<!--#include file="../includes/cnn.inc"-->

<script language="jscript" type="text/jscript">
function AGRANDA() {
    top.parent.window.document.getElementById('body0').height = 480
}


function dd2(ff) {	// LLENA TEXTBOX ADICIONALES AL COMUN
    // LOS DEL COMUN SON CODIGO Y DESCRIPCION
    var t = document.all.TABLA;
    var pos = parseInt(ff,10);
    dd(ff);
}
oldrow = 1
</script>

<%op = request.QueryString("op")
TDA = request.QueryString("TDA")
tip = request.QueryString("tip")
ini = request.QueryString("ini")
fin = request.QueryString("fin")
cod = request.QueryString("cod")
'*************************************************************************************************************
cad =   " set dateformat dmy    ;                                                                        " & _
        " SELECT FECHA, ISNULL(TIPDOC,'') AS DOC, ISNULL(NUMDOC,'') AS NUMDOC, ISNULL(DOCORI,'') AS ORI, " & _
        " isnull(NUMORI,'') AS NUM_ORI,  CLIENTE, NOMBRE,                                                " & _
        " CODART, DESCRI, CANT = CASE WHEN SALE = 0 THEN ENTRA ELSE SALE END, LISTA1 AS PVP, DCT,        " & _
        " ISNULL(PORDES,0) AS PORDES, IGV, PRECIO,ITEM, GRUPO, TIENDA  as TDA, operacion as OPE          " & _
        " FROM VIEW_VENTAS_ARTICULO WHERE  FECHA between '"&INI&"' AND DateAdd(day,1,'"&fin&"')          " & _
        " AND TIPMOV = '"&OP&"'                                                                          " 

if tip = "PP" then
        cad = cad +"  and CODART =  '"&cod&"'"
else
        cad = cad +"  and grupo =  '"&cod&"'"
end if
if Tda <> "TT" then cad = cad +"  and tienda =  '"&tda&"'"
cad = cad + "ORDER BY TIENDA, FECHA, DOC, NUMDOC "        

'*************************************************************************************************************
'RESPONSE.WRITE(CAD)
'response.end
rs.open cad,cnn
if rs.recordcount <=0 then RESPONSE.End
%>

<body>
<center>

<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="TABLA" name="TABLA"  >
	<tr> 
    <%cont = 1 %>
    <%FOR I=0 TO RS.FIELDS.COUNT-1 %>
        <td align="center" class="Estilo8"><%=RS.FIELDS(I).NAME %></td>
    <%NEXT %>
	</tr >
    <%can= 0 %>
<%do while not rs.eof %>
	<tr  bgcolor="<% if CONT mod 2  = 0 THEN 
                response.write(Application("color1"))
                else
	            response.write(Application("color2"))
	            end IF%>"
	            onclick="dd('<%=(cont)%>')" id="fila<%=Trim(Cstr(cont))%>"   class="EstiloT" align="left" style="cursor:pointer" ondblclick="show('<%=cont %>')">
		<td align="center">&nbsp;<%=FORMATDATETIME(RS.FIELDS.ITEM(0),2)%>&nbsp;</td>
        <%FOR I=1 TO RS.FIELDS.COUNT-1 %>
    		<td>&nbsp;<%=ucase(RS.FIELDS.ITEM(I))%>&nbsp;</td>
        <%NEXT %>
	</tr> 
    <% can = can +  cdbl(RS.FIELDS.ITEM(9))%>
    <%cont = cont +1 %>
    <%rs.movenext%>
<%loop %>
<tr class="Estilo8">
<td colspan="9" style="text-align:right">Total&nbsp;</td>
    <td align="center" ><%=can%></td>
<td colspan="11"></td>
</tr>

</table>
<script language="jscript" type="text/jscript">
rec = '<%=rs.recordcount %>'    
    if (rec > 0)
        dd2('1');

function show(ff) {
    var t = document.all.TABLA;
    var op = parseInt(ff,10);
    var pos = t.rows(op).cells(18).innerText
    window.open('showmov.asp?ope='+trim(pos))
    }
</script>
</center>
</body>
<%RS.CLOSE %>

</html>
