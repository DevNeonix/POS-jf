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
<!--#include file="../funcionesEfact.asp"-->
<!--#include file="../includes/cnn.inc"-->

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

TBL = "REPRINT_REAL"



cad = " set dateformat dmy; SELECT *,(select TOP 1 ticket from REIMPRIME tx where tx.OPERACION  collate SQL_Latin1_General_CP1_CI_AI= m1.OPERACION order by fecha desc ) ticket FROM "&TBL&" AS M1            " & _
      " full outer JOIN view_CLIENTES_all  AS C1 ON M1.CLIENTE collate SQL_Latin1_General_CP1_CI_AI  = C1.CLIENTE       " & _
      " WHERE m1.FECdoc between '"&INI&"' AND DateAdd(day,1,'"&FIN&"')  " & _
      " and   CODDOC in ('BL','FC','NC','TR') "
if pos<>"TT" then cad = cad&" and TIENDA ='" & POS & "' "
cad = cad&"order by coddoc,serie,numdoc,operacion"
'*************************************************************
'response.write(cad)
'response.end
'response.write(Numlet(cdbl("997.39"))&"<br>")
'response.write(Numlet(cdbl("1096.87"))&"<br>")
'response.write(Numlet(cdbl("1195.59"))&"<br>")
'response.write(Numlet(cdbl("1144.35"))&"<br>")
'response.write(Numlet(cdbl("1084.14"))&"<br>")
'response.write(Numlet(cdbl("1056.62"))&"<br>")
'response.write(Numlet(cdbl("1211.15"))&"<br>")
'response.write(Numlet(cdbl("1194.05"))&"<br>")
'response.write(Numlet(cdbl("991.85"))&"<br>")

'response.write(cad)
rs.open cad,cnn
if rs.recordcount <=0 then RESPONSE.End
%>

<body onload="AGRANDA()">
<center>

<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="TABLA" name="TABLA"  >
	<tr> 
        <td align="center" class="Estilo8">DOCUMENTO</td>
	    <td align="center" class="Estilo8">CLIENTE</td>
        <td align="center" class="Estilo8">FECHA</td>
        <td align="center" class="Estilo8">Precio</td>
        <td align="center" class="Estilo8">Operacion</td>
        <td align="center" class="Estilo8">TDA</td>
        <td align="center" class="Estilo8">DOC</td>
        <!--<td align="center" class="Estilo8">TICKET</td>-->
        <td align="center" class="Estilo8" colspan="3">Operaciones</td>
	</tr >

    <%cont =1
      RS.MOVEFIRST%>
         
    <%do while not rs.eof %>
         <tr  bgcolor="<% if CONT mod 2  = 0 THEN 
                response.write(Application("color1"))
                else
	            response.write(Application("color2"))
	            end IF%>" class="Estilo0" 

                <%  if IsNull(rs("TICKET")) or IsEmpty(rs("TICKET"))  then
                    
                %>
                    ondblclick="printa('<%=cont%>');"
                <%  
                    else
                    %>ondblclick="window.open('/apijf/public/index.php/show?ticket=<%=rs("TICKET")%>&tipo=pdf')"<%
                    end if
                %>

	            id="fila<%=Trim(Cstr(cont))%>"  style="text-align:left;"  onclick="dd('<%=(cont)%>')">

            <td style="padding:12px">&nbsp;<%=rs("CODdoc")&" "&rs("serie")&"-"&rs("numdoc")%></td>
            <td>&nbsp;<%=rs("cliente")&" - "&rs("nombre")%></td>
            <td>&nbsp;<%=formatdatetime(rs("Fecdoc"),2)%>&nbsp;</td>
             <%if isnull(rs("total"))    then total = "" else total = formatnumber(rs("total"),2,,true) %>
            <td style="text-align:right">&nbsp;<%=total%>&nbsp;</td>
            <td style="text-align:right">&nbsp;<%=rs("operacion")%>&nbsp;</td>
            <td style="text-align:center">&nbsp;<%=rs("tienda")%>&nbsp;</td>
            <td style="text-align:center">&nbsp;<%=rs("coddoc")%>&nbsp;</td>
            <!--<td style="text-align:center" style="width: auto;padding: 5px">
                 <% '' if IsNull(rs("TICKET")) or IsEmpty(rs("TICKET")) then
                 %>
                <input type="text" id="tk<%=cont%>" value='<%=rs("TICKET")%>' style="font-size: 12px;width: 100%;    box-sizing: border-box;">
                <button onclick="window.open('../comun/inserticket.asp?ope=<%=trim(rs("operacion"))%>&ticket='+document.getElementById('tk<%=cont%>').value)">Guardar</button>
                <%'else
                   '' response.write(rs("TICKET"))
                'end if%>
              
            </td>-->
            <%  if IsNull(rs("TICKET")) or IsEmpty(rs("TICKET")) then
                    
                    'response.write(reemplazaNull(total,"0"))
                        if reemplazaNull(total,"") = "" and (trim(rs("CODdoc"))="BL" or  trim(rs("CODdoc"))="FC" OR  trim(rs("CODdoc"))="NC") then
                        %><td>No puedes emitir un doc. con precio total 0</td><%
                        else
            %>
                
                
            <%
                        end if
                else
            %>
            <td style="text-align:center">
                <a href="#" onclick="PopupCenter('/apijf/public/index.php/show?ticket=<%=rs("TICKET")%>&tipo=pdf',600,600)"><img style="width:35px" src="../images/print.jpg"/></a>
                <a href="#" onclick="window.open('/apijf/public/index.php/show?ticket=<%=rs("TICKET")%>&tipo=pdf&download=true')"><img style="width:35px" src="../images/disk.gif"/></a>
                <%if trim(Request.Cookies("tienda")("usr")) = "1" or UCASE(trim(Request.Cookies("tienda")("usr"))) = "CSABA" or UCASE(trim(Request.Cookies("tienda")("usr"))) = "MALAGA" then%>
                
                <td style="text-align:center;color:tomato"><a href="#" onclick="PopupCenter('/apijf/public/index.php/show?ticket=<%=rs("TICKET")%>&tipo=cdr&download=true',600,600)">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CDR&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                 <td style="text-align:center;color:tomato"><a href="#" onclick="PopupCenter('/apijf/public/index.php/show?ticket=<%=rs("TICKET")%>&tipo=xml&download=true',600,600)">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;XML&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                
                
                
                <%end if%>
                <%
                'Mensaje Baja
                cbaja="select * from JACINTA..efact_bajas where convert(int,estado) between 200 and 300 and operacion = '"&rs("operacion")&"'"
                Set RSv = Server.CreateObject("ADODB.Recordset")
                RSv.ActiveConnection = Cnn
                RSv.CursorType       = 3 
                RSv.LockType         = 1 
                RSv.CursorLocation   = 3 
                if RSv.state>0 then
                    RSv.close
                end if
                RSv.open cbaja,cnn
                if rsv.recordcount > 0 then
                    RSv.movefirst
            %>

            <td>Este documento ha sido dado de <b>baja</b> <%=rsv("fecbaja")%></td>

            <%
                else
                %>
                <td style="text-align:center;color:tomato"><a href="#" onclick="PopupCenter('/ppbaja.asp?ope=<%=rs("operacion")%>',600,600);ccc()">Anular documento Sunat</a> </td>
                <td></td>
                <%
                end if
            %>
                <td style="text-align:center;color:tomato"><a href="#" onclick="PopupCenter('sendmaildocumento.asp?doc=<%=rs("CODdoc")&" "&rs("serie")&"-"&rs("numdoc")%>&ticket=<%=rs("TICKET")%>',300,300)"><img style="width:35px" src="../images/mail.png"></a> </td>
                <!--<a href="http://intranet:8088/apijf/public/index.php/download?ticket=<%=rs("TICKET")%>&tipo=pdf">DESCARGAR</a>-->
                <!--<a href="http://intranet:8088/apijf/public/index.php/download?ticket=<%=rs("TICKET")%>&tipo=cdr">CDR</a>-->
                <!--<a href="http://intranet:8088/apijf/public/index.php/download?ticket=<%=rs("TICKET")%>&tipo=xml">XML</a>-->
            </td>
            <%
                end if
            %>
            <%if DateValue(rs("Fecdoc")) > DateValue("01-10-2019") and (trim(Request.Cookies("tienda")("usr")) = "1" or UCASE(trim(Request.Cookies("tienda")("usr"))) = "CSABA" or UCASE(trim(Request.Cookies("tienda")("usr"))) = "MALAGA") then%>
            
            
            <%end if%>
           <%
                    ti = "boleta"
                    xti = "03"
                    if left(ucase(trim(rs("coddoc"))),1) = "F" then
                        ti = "factura"
                        xti = "01"
                    elseif ucase(trim(rs("coddoc"))) = "NC" then
                        ti = "NC"
                        xti = "07"
                    elseif ucase(trim(rs("coddoc"))) = "ND" then
                        ti = "ND"
                        xti = "08"
                    end if
                %>
            <%if left(ucase(trim(rs("coddoc"))),1) = "B" or left(ucase(trim(rs("coddoc"))),1) = "F" or left(ucase(trim(rs("coddoc"))),2) = "NC"  or left(ucase(trim(rs("coddoc"))),2) = "ND" then%>
            <td style="text-align:center;color:tomato"><a href="#" onclick="if(confirm('Esta seguro de enviar a Sunat?')){PopupCenter('/pp<%=ti%>.asp?ope=<%=rs("operacion")%>',600,600);ccc()}">Enviar Efact</a> </td>
             <%if trim(Request.Cookies("tienda")("usr")) = "1" or ucase(trim(Request.Cookies("tienda")("usr"))) = "CSABA" or trim(Request.Cookies("tienda")("usr")) = "MALAGA" then%>
                <td style="text-align:center;color:tomato"><a href="#" onclick="PopupCenter('/apijf/public/index.php/grabaticket?documento=<%=trim(rs("serie")&"-"&rs("numdoc"))%>&tipdoc=<%=xti%>&operacion=<%=trim(rs("operacion"))%>',600,600);ccc()">CONSULTA TICKET</a> </td>
            <%end if%>
            <%end if%>

            


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
function ccc(){

    setTimeout(function(){
        this.location.reload();
    }, 16000);
}
function printa(ff) {
    var pos = oldrow
    var t = document.all.TABLA; 
    var cad = ''   
    tip = ltrim(t.rows[pos].cells[6].innerText );
    ope = ltrim(t.rows[pos].cells[4].innerText);
    tda = ltrim(t.rows[pos].cells[5].innerText );
    
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
