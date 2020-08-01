<%@Language=VBScript%>
<%Response.Buffer=true%>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="../COMUN/FUNCIONESCOMUNES.ASP"-->
<style type="text/css">
body,td,th {
	font-size: x-small;
	font-family: "Century Gothic";
	font-size: 10px;
	color:#000000;
}
</style>

<!--#include file="../COMUN/COMUNqry.ASP"-->
<!--#include file="../includes/Cnn.inc"-->

<% OPE = REQUEST.QueryString("OPE")
 tda = REQUEST.QueryString("tda")
CAD=    " SELECT M1.CLIENTE, C1.NOMBRE, C1.DIRECCION, M2.CODART,        " & _
        " M1.FECDOC, M1.SERIE, M1.NUMDOC ,M2.SALE, V1.DESCRI,           " & _ 
        " ROUND(M2.PRECIO,2) AS PRECIO, M2.PORDES, m1.pvp as bruto,     " & _
        " (M2.PRECIO - M2.IGV + M2.DESCUENTO) / M2.SALE AS PVT,         " & _
        " M1.TOTAL, M1.DESCUENTO, m1.subtot, m1.igv, AR_CUNIDAD AS UNI, " & _
        "  isnull(m1.iscqty,0) as iscqty, isnull(m1.isc,0) as isc       " & _
        " FROM MOVIMCAB AS M1                                           " & _
        " INNER JOIN CLIENTES AS C1 ON C1.CLIENTE                       " & _
        " collate SQL_Latin1_General_CP1_CI_AI= M1.CLIENTE              " & _
        " INNER JOIN MOVIMDET AS M2 ON M1.OPERACION                     " & _
        " collate  SQL_Latin1_General_CP1_CI_AI= M2.OPERACION           " & _
        " inner join view_ARTICULOS_TIENDA V1 ON M2.CODART              " & _
        " collate  SQL_Latin1_General_CP1_CI_AI = V1.CODIGO             " & _
        " WHERE M1.OPERACION ='"&OPE&"' AND M2.SALE > 0                 " & _
        " AND V1.TIENDA = '"&TDA&"'                                     " & _
        " ORDER BY M2.ITEM                                              "
'response.Write(cad)
RS.OPEN CAD,CNN

RS.MOVEFIRST         
total = round(cdbl(rs("total")),2)
descuento = round(cdbl(rs("descuento")),2)
bruto = round(cdbl(rs("bruto")),2)
IGV = round(cdbl(rs("IGV")),2)
subtot = round(cdbl(rs("subtot")),2)
isc = rs("isc")
iscq = cint(rs("iscqty"))
%>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>

<style type="text/css">
body { font-family:Century Gothic; font-size:10px; color:Black; padding-left:14px; padding-top:150px; width:750px}
body2 { font-family:Century Gothic; font-size:11px;  color:Black; padding-left:14px; padding-top:150px;  }
</style>

</head>
<%DIM AMES(12) 
AMES(1) = "ENERO"
AMES(2) = "FEBRERO"
AMES(3) = "MARZO"
AMES(4) = "ABRIL"
AMES(5) = "MAYO"
AMES(6) = "JUNIO"
AMES(7) = "JULIO"
AMES(8) = "AGOSTO"
AMES(9) = "SETIEMBRE"
AMES(10) = "OCTUBRE"
AMES(11) = "NOVIEMBRE"
AMES(12) = "DICIEMBRE" %>

<body>
<!-- onunload="cierra()" -->
<table width="100%" align="left" border="0">
<tr>
<td>
	<center>
	<img src="../images/print.jpg" border="0" id="prn" name="prn" onClick="printa();" style="display:block; cursor:pointer"/>
	</center>
</td>
</tr>
<tr>
<td>
	<table border="0" align="right" cellpadding="0"cellspacing="0" width="100%">
   	<tr height="05px"  valign="bottom">
    	<td align="right"><%=trim(rs("serie")) %> - <%=trim(rs("numdoc")) %></td>
        <td width="18%">&nbsp; </td>
    </tr>
    <tr height="25px" >
    	<td colspan="2">&nbsp;</td>
    </tr>
    <tr>
    	<td>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=DAY(rs("fecdoc")) %>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <%=AMES(MONTH(RS("FECDOC"))) %>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=YEAR(RS("FECDOC")) %>
		</td>
	</tr>
	</table>
</td>
</tr>
<tr>
<td>
	<table border="0"   align="left" cellpadding="0"cellspacing="0" width="700px">
    <tr  valign="bottom">
        <td width="60px">&nbsp;</td>
        <td  align="left" ><%=left(trim(RS("NOMBRE")),65)%></td>
        <td align="left"><%=RS("CLIENTE")%>&nbsp;</td>
        <td align="right" >&nbsp;</td>
    </tr>
    <tr>
        <td >&nbsp;</td>
        <td align="center">&nbsp;</td>
        <td align="left">&nbsp;</td>
        <td align="right" >&nbsp;</td>
    </tr>
    <tr valign="bottom" height="8px">
        <td >&nbsp;</td>
        <td  align="left"><%=LEFT(TRIM(RS("DIRECCION")),65) %></td>
        <td align="left">&nbsp;</td>
        <td align="center">&nbsp;</td>
    </tr>
	</table>
</td>
</tr>
<tr>
<td>
	<table border="0"   align="left" cellpadding="0"cellspacing="0" width="100%">
    <tr height="60px" valign="bottom">
       	<td>&nbsp;</td>
    </tr>
	</table>
</td>
</tr>
<tr>
<td>
	<table border="0"   align="left" cellpadding="0"cellspacing="0" WIDTH="720px">
    	<%do while not rs.eof %>
      	<tr  valign="bottom">
        	<td width="35px" align="left"><%=UCASE(left(trim(RS("codart")),10)) %></td>
        	<td width="40px" align="center"><%=RS("SALE") %>&nbsp;</td>
        	<td width="450px" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=TRIM(RS("uni"))%>&nbsp;&nbsp; <%=left(trim(RS("DESCRI")),62) %></td>
        	
        	<td width="70px" align="right"><%=formatnumber(round(RS("PVT"),2),2,,,true) %>&nbsp;</td>
            <td width="50px" align="right"><%if  round(RS("PORDES"),0)> 0 then %>-<%=round(RS("PORDES"),0) %>&nbsp;%<%end if %></td>
           	<td width="50px" align="right"><%=formatnumber(rs("PRECIO"),2)%>&nbsp;</td>
        	<%i = i + 1 %>
    	</tr>
	    	<%rs.movenext %>
    	<%loop%>
    	<%for m=i to 14%>
    	<tr  valign="bottom">
        	<td align="left">&nbsp;</td>
        	<td align="right">&nbsp;</td>
        	<td align="left">&nbsp;</td>
        	<td align="right">&nbsp;</td>
        	<td align="right">&nbsp;</td>
    	</tr>
    	<%next %>
        <% if  tienda="16" then %>
                    <tr><td colspan="7" style="font-size:12px"><center>No se acepta cambios ni devoluciones.</center></td></tr>
                        <%else%>
                        <tr><td colspan="7" style="font-size:12px"><center>No se acepta cambios ni devoluciones en ropa interior ni prendas con descuento, pijamas sólo cambios con documento de vta., máx 7 dias.
                <%if cint(iscq) > 0 then 
                    response.write("<b>")
                    response.write("(Se entregaron "& cstr(iscq) & " bolsas plasticas )" )
                    response.write("</b>")
                end if%>
                        
                        
                        </center></td></tr>
                    <% end if %>
	</table>
</td>
</tr>
<tr>
<td>
	<table border="0"   align="left" cellpadding="0"cellspacing="0" width="100%" >
    <tr  valign="bottom" height="60px" >
        <td width="10px">&nbsp;</td>
        <td id="sonn" style="font-weight:800;font-size:14px"></td>
    </tr>
	</table>
</td>
</tr>
<tr>
<td>
	<table border="0" align="left" cellpadding="0" cellspacing="0" width="720px">
    <tr height="9px">
        <td colspan="2" >&nbsp;</td>
    </tr>
    <tr  valign="bottom" height="15px">
        <td width="665px" align="right">S/.</td>
        <td width="55px" align="right" ><%=formatnumber(bruto,2)%>&nbsp;</td>
    </tr>
    <tr  valign="bottom" height="15px">
        <td align="right">S/.</td>
        <td  align="right" ><%=formatnumber(descuento,2)%>&nbsp;</td>
    </tr>
    <tr  valign="middle" height="10px">
        <td align="right">&nbsp;</td>
        <td  align="right" >-----------</td>
    </tr>
    <tr  valign="top" height="20px">
        <td align="right">SUB-TOTAL &nbsp;&nbsp;&nbsp;S/.</td>
        <td align="right" ><%=formatnumber(SUBTOT,2)%>&nbsp;</td>
    </tr>
    <tr  valign="bottom" height="13px">
        <td align="right">S/.</td>
        <td align="right" ><%=formatnumber(IGV,2)%>&nbsp;</td>
    </tr>
     <tr  valign="bottom" height="13px">
        <td align="right">ISC S/.</td>
        <td align="right" ><%=formatnumber(ISC,2)%>&nbsp;</td>
    </tr>
    <tr  valign="bottom" height="13px">
        <td align="right" >S/.</td>
        <td align="right" ><%=formatnumber(TOTAL,2)%>&nbsp;</td>
    </tr>
	</table>
</td>
</tr>
</table>    
  
<script language="jscript" type="text/jscript">
document.getElementById('sonn').innerText = 'Son : ' + FComson('<%=total%>')

function printa() {
    document.all.prn.style.display = 'none'
    window.print()
   
}

</script> 
</body>
</html>
