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

<LINK REL="stylesheet" TYPE="text/css" HREF="VENTAS.CSS">
<!--#include file="comun/funcionescomunes.asp"-->
<!--#include file="comun/comunqry.asp"-->
<!--#include file="includes/funcionesVBscript.asp"-->
<!--#include file="includes/cnn.inc"-->

<script language="jscript" type="text/jscript">
    var oldrow = 1
function AGRANDA() {
    top.parent.window.document.getElementById('body0').height = 480
}
function dd2(ff) {	// LLENA TEXTBOX ADICIONALES AL COMUN
    // LOS DEL COMUN SON CODIGO Y DESCRIPCION
    var t = document.all.TABLA;
    var pos = parseInt(ff);
    dd(ff);
}
</script>

<%
pos = request.QueryString("pos")  ' tienda
ini = request.QueryString("ini")  ' fecha de inicio
fin = request.QueryString("fin")  ' fecha de fin

'*************************************************************
' MM 13-ABR 2013 REIMPRESION DE DOCUMENTOS
cad = " set dateformat dmy; SELECT * FROM movimcab AS M1            " & _
      " INNER JOIN CLIENTES  AS C1 ON M1.CLIENTE = C1.CLIENTE       " & _
      " WHERE m1.FECdoc between '"&INI&"' AND DateAdd(day,1,'"&FIN&"')  " & _
      " and ltrim(rtrim(coddoc)) in ('BL','FC','NC','TR') "
if pos<>"TT" then cad = cad&" and TIENDA ='" & POS & "' "
cad = cad&"order by coddoc,serie,numdoc,operacion"
'*************************************************************
'response.write (cad)
'response.end
rs.open cad,cnn
if rs.recordcount <=0 then RESPONSE.End
%>

<body onload="AGRANDA()">
<center>

<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="TABLA" name="TABLA"  >
	<tr> 
        <td align="center" class="Estilo8">TIP</td>
        <td align="center" class="Estilo8">SERIE</td>
        <td align="center" class="Estilo8">NUM DOC</td>
        <td align="center" class="Estilo8">COD CLI</td>
	    <td align="center" class="Estilo8">NOMBRE</td>
        <td align="center" class="Estilo8">FECHA</td>
        <td align="center" class="Estilo8">Precio</td>
        <td align="center" class="Estilo8">Operacion</td>
        <td align="center" class="Estilo8">TDA</td>
        <td align="center" class="Estilo8">DOC</td>
	</tr >

    <%cont =1
      RS.MOVEFIRST%>
         
    <%do while not rs.eof %>
         <tr  bgcolor="<% if CONT mod 2  = 0 THEN 
                response.write(Application("color1"))
                else
	            response.write(Application("color2"))
	            end IF%>" class="Estilo0" id="fila<%=Trim(Cstr(cont))%>"  style="text-align:left"  onclick="dd('<%=(cont)%>')">

            <td>&nbsp;<%=rs("CODdoc")%></td>
            <td>&nbsp;<%=rs("serie")%></td>
            <td><input id="num" name="num" value="<%=rs("numdoc")%>" class="Estilo5" /></td>
            <td><%=rs("cliente")%></td>
            <td>&nbsp;<%=rs("nombre")%></td>
            <td><%=formatdatetime(rs("Fecdoc"),2)%></td>
             <%if isnull(rs("total"))    then total = "" else total = formatnumber(rs("total"),2,,true) %>
            <td style="text-align:right">&nbsp;<%=total%>&nbsp;</td>
            <td style="text-align:right">&nbsp;<%=rs("operacion")%>&nbsp;</td>
            <td style="text-align:center">&nbsp;<%=rs("tienda")%>&nbsp;</td>
            <td style="text-align:center">&nbsp;<%=rs("coddoc")%>&nbsp;</td>
        </tr>
        
        
           <%rs.movenext%>
           <%cont = cont + 1%>
        <%LOOP%>


</table>
</center>
</body>
<script type="text/jscript" language="jscript">
var rec = '<%=rs.recordcount %>'
    if (rec > 0)
        dd2('1');


function printa(ff) {
    var pos = oldrow
    var t = document.all.TABLA; 
    var cad = ''   
    tip = ltrim(t.rows(pos).cells(6).innerText );
    ope = ltrim(t.rows(pos).cells(4).innerText);
    tda = ltrim(t.rows(pos).cells(5).innerText );
    
    if (trim(tip) == 'FC') 
        cad = 'prnfactura.asp?ope='+ ope + '&tda='+ tda
    else if (trim(tip) == 'BL')
        cad = 'prnboleta.asp?ope=' + ope + '&tda=' + tda
    else if (trim(tip) == 'TR')
        cad = '../NOTASALIDA.asp?ope=' + ope + '&tda=' + tda
    else
        cad = 'prnnota.asp?ope=' + ope + '&tda=' + tda        
window.open (cad)


}

</script>


<%RS.CLOSE %>
</html>
