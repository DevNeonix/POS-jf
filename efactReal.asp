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
<!--#include file="includes/funcionesVBscript.asp"-->
<!--#include file="includes/cnn.inc"-->
<style>
    .tabla{
        font-size: 12px;
        max-width: 840px;
    }
    .tabla > td{
            border:1px solid #333
        }
    
</style>
<script language="jscript" type="text/jscript">
function AGRANDA() {
    top.parent.window.document.getElementById('body0').height = 680
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

<body onload="AGRANDA()">

<center>
<table width="100%" >
	
    <tr><td align="center" class="Estilo6">Ingrese el Documento del Real</td></tr>
	
</table>
<table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="1"  cellspacing="2"  border="0" align="center" >
    <tr valign="middle">
    	<td class="Estilo1">Doc</td>
        <td class="Estilo1">Serie</td>
        <td class="Estilo1">Numero</td>
    </tr>
    <tr valign="middle" >
        <td class="Estilo12" align="left"  rowspan="2">
        <select id="TIP" name="TIP" class="Estilo12">
            <option value="F">Factura</option>
            <option value="B">Boleta</option>
            <option value="N">Nota de Credito</option>
        </select>
        </td>
        <td class="Estilo12" align="left"  rowspan="2"><input  name="SER" id="SER" size="5"  onkeyup="this.value=(this.value)" maxlength="4"></td>
        <td class="Estilo12" align="left"  rowspan="2"><input  name="NRO" id="NRO" onKeyUp="this.value=toInt(this.value)" maxlength="8" size="10"></td>
        <td><img src="images/ok.gif" onClick="manda()" style="cursor:pointer;"/></td>
    </tr>
</table>

<%
    cad =  "select * from real_efact_documentos where GETDATE() < dateadd(day,30,F5_DFECDOC) order by 2 desc,3,4 desc"
'Response.Write(cad)
rs.open cad,cnn
%>
    <table class="tabla" style="font-family: sans-serif;">
        <tr style="background: #C82F8A;color:#fff;">
            <td>NUMDOC</td>
            <td>CLIENTE</td>
            <td>TOTAL</td>
            <td>OPERACION</td>
            <td>TICKET</td>
            <td>FECEMI</td>
            <td colspan="4">PDF</td>
        </tr>
    
<%


if rs.recordcount > 0 then
    rs.movefirst
    for i=0 to rs.recordcount - 1
        %>

        <tr style="color:#333">
            <td style="padding: 5px"><%=RS("ctd")%>-<%=RS("numser")%>-<%=RS("numdoc")%></td>
            <td style="padding: 5px"><%=RS("nomcli")%></td>
            <td style="padding: 5px"><%=RS("f5_nimport")%></td>
            <td style="padding: 5px"><%=RS("operacion")%></td>
            <td style="padding: 5px"><%=RS("ticket")%></td>
            <td style="padding: 5px"><%=RS("f5_dfecdoc")%></td>
            <td>
                <table>
                    <tr>
                        <td style="padding: 5px" onclick="PopupCenter('./apijf/public/index.php/show?ticket=<%=rs("TICKET")%>&tipo=pdf',600,600)">
                            <a  href="javascript:void()">PDF</a>
                        </td>
                   
                        <td style="padding: 5px" onclick="PopupCenter('./apijf/public/index.php/download?ticket=<%=rs("TICKET")%>&tipo=xml',600,600)">
                            <a  href="javascript:void()">XML</a>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 5px" onclick="PopupCenter('./apijf/public/index.php/download?nom=<%=RS("ctd")%>-<%=RS("numser")%>-<%=RS("numdoc")%>&ticket=<%=rs("TICKET")%>&tipo=pdf',600,600)">
                            <a  href="javascript:void()">DESCARGAR PDF</a>
                        </td>
                  
                    
                    <%
        
                            ti = "boleta"
                            xti = "03"
                            if left(ucase(trim(rs("ctd"))),1) = "F" then
                                ti = "factura"
                                xti = "01"
                            elseif ucase(trim(rs("ctd"))) = "NC" then
                                ti = "NC"
                                xti = "07"
                            elseif ucase(trim(rs("ctd"))) = "ND" then
                                ti = "ND"
                                xti = "08"
                            end if
        
                    %>
                    <td style="text-align:center;color:tomato"><a  href="javascript:void()" onclick="PopupCenter('/apijf/public/index.php/grabaticket?documento=<%=trim(rs("numser")&"-"&rs("numdoc"))%>&tipdoc=<%=xti%>&operacion=<%=trim(rs("operacion"))%>',600,600);">CONSULTA TICKET</a> </td></tr>
                    
                </table>
            </td>
        </tr>
        <tr><td colspan="100"><hr/></td></tr>
        <%

        rs.movenext

    next

end if

%>
</table>

<iframe  width="100%" src="" id="body00" name="body00" scrolling="yes" frameborder="1" height="400" align="middle" style="display:none" ></iframe>
</center>

<script language="jscript" type="text/jscript">
    function manda() {
        var cad = ''
        var tipo = document.getElementById("TIP").value
        var serie = trim(document.getElementById('SER').value)
        var numd = trim(document.getElementById('NRO').value)
        if (serie == '') {
            alert("ingrese la serie por favor")
            return false
        }
        else if (serie.length < 3) {
            alert("La serie tiene que ser de 4 digitos");
            return false; 
        }

        if (numd == '') {
            alert("ingrese el nro de documento por favor")
            return false
        }
        else if (numd.length > 0) {

            document.getElementById('NRO').value = strzero(document.getElementById('NRO').value,7)
        }

        switch (tipo) {
            case "F":
                cad = 'ppfacturaReal.asp?ctd=FT';
                break
            case "B":
                cad = 'ppboletaReal.asp?CTD=BV';
                break;

            case "N":
                cad = 'ppncReal.asp?CTD=NC'
                break;
        }
        cad += '&SER=' + serie + '&doc=' + strzero(document.getElementById('NRO').value, 7)
        //alert(cad)
        window.open(cad,"Real_Doc")    
        setTimeout(function(){
            this.location.reload();
        },15000)
    }
</script>
</body>
</html>
