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
'response.write(request.QueryString("docs"))
if ltrim(rtrim(request.QueryString("docs"))) <> "" then
doc = "''" + left(replace(request.QueryString("docs"),",","'',''"), len(replace(request.QueryString("docs"),",","'',''"))-3)
else
    doc = ""
end if
'*********************************************************************************************
 cad =   " EXEC SP_MOVIM_FECHA  '"&INI&"' , '"&FIN&"', '"&TDA&"', '"&DOC&"' "
'*********************************************************************************************
'RESPONSE.WRITE(cAD)
'response.end
rs.open cad,cnn
if rs.recordcount <=0 then RESPONSE.End
%>

<body onload="AGRANDA()">
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
	            onclick="dd('<%=(cont)%>',0)" id="fila<%=Trim(Cstr(cont))%>" ondblclick="SHOW('<%=CONT%>')">
		<%for i =0 to 6 %>
            <td align="center" CLASS="EstiloT">&nbsp;<%=trim(RS.FIELDS.ITEM(i))%>&nbsp;</td>
		<%next %>
        <td align="center" CLASS="EstiloT">&nbsp;<%=FORMATDATETIME(RS.FIELDS.ITEM(i),2)%>&nbsp;</td>
        <%for i =8 to RS.FIELDS.COUNT -3 %>
            <%if isnull(RS.FIELDS.ITEM(i)) then nume = 0 else nume = cdbl(RS.FIELDS.ITEM(i)) %>
            <td align="RIGHT" CLASS="EstiloT">&nbsp;<%=FORMATNUMBER(nume,2,,,TRUE)%>&nbsp;</td>
		<%next %>
        <%for i = RS.FIELDS.COUNT - 3 to RS.FIELDS.COUNT-1 %>
            <td align="LEFT" CLASS="EstiloT">&nbsp;<%=trim(RS.FIELDS.ITEM(i))%>&nbsp;</td>
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

function SHOW(op) {
    opc = "directories=no,status=no,titlebar=yes,toolbar=no,hotkeys=yes,location=no"; 
    var t = document.all.TABLA;
    var pos = parseInt(op);
    OPE = trim(eval("document.all.fila" + pos + ".cells(1).innerText"));
    DOC = trim(eval("document.all.fila" + pos + ".cells(2).innerText"));
    DOC += '-->' + trim(eval("document.all.fila" + pos + ".cells(3).innerText"));
    DOC += '-' + trim(eval("document.all.fila" + pos + ".cells(4).innerText"));
//    alert(OPE)
        cad = 'showmov.asp?ope=' + OPE + '&tit=' + DOC
//        alert(DOC)
        window.open(cad)
        return true
    }
</script>
</center>
</body>
</html>
