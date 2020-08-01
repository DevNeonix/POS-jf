<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<% tienda = Request.Cookies("tienda")("pos") %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html>

<head>
<!--#include file="../COMUN/FUNCIONESCOMUNES.ASP"-->
<style type="text/css">
body,td,th {
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
        " ROUND(M2.PRECIO,2) AS PRECIO, M2.PORDES, AR_CUNIDAD AS UNI,   " & _
        " (M2.PRECIO - M2.IGV + M2.DESCUENTO)/M2.SALE AS PVT, M1.TOTAL, " & _
        " isnull(m1.iscqty,0) as iscqty, isnull(m1.isc,0) as isc,       " & _
        " isnull( round(M1.DESCUENTO *(1+ (ar_nigvpor/100)),2),0)       " & _
        "  as DESCUENTO                                                 " & _
        " FROM MOVIMCAB AS M1                                           " & _
        " INNER JOIN CLIENTES AS C1 ON C1.CLIENTE                       " & _
        " collate SQL_Latin1_General_CP1_CI_AI = M1.CLIENTE             " & _
        " INNER JOIN MOVIMDET AS M2 ON M1.OPERACION                     " & _
        " collate SQL_Latin1_General_CP1_CI_AI= M2.OPERACION            " & _
        " inner join view_ARTICULOS_TIENDA V1 ON M2.CODART              " & _
        " collate SQL_Latin1_General_CP1_CI_AI = V1.CODIGO              " & _
        " WHERE M1.OPERACION ='"&OPE&"' AND M2.SALE > 0                 " & _
      " AND V1.TIENDA = '"&tda&"'                                       " & _
      " ORDER BY M2.ITEM                                                "
'response.Write(cad)
'response.end
RS.OPEN CAD,CNN
if rs.recordcount > 0 then 
    RS.MOVEFIRST         
else
    response.write ("Documento sin detalle")
    response.End
end if
total = round(cdbl(rs("total")),2)
descuento = round(cdbl(rs("descuento")),2)
isc = rs("isc")
iscq = cint(rs("iscqty"))
'response.Write(cad)
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>

<!-- <style type="text/css">
-- body { font-family:Century Gothic;  font-size:10px;  color:Black; padding-left:4px;   }
body2 { font-family:Century Gothic;  font-size:11px;  color:Black; padding-left:4px;   } 
</style> -->

<style type="text/css">
body {
	margin-left: 0px;
	margin-top: 0px;
	width: 750px;
	font-family:"Century Gothic";
	font-size:10px;
	color:#000000;
	padding-left:0px;
}
</style>
</head>

<!-- style="margin-left:0px; margin-top:0px; width:750px;" -->

<body>

<table width="100%" align="left" border="0">
	<tr>
    	<td align="right">
        	<img src="../images/print.jpg" border="0" id="prn" name="prn" onClick="printa();" style="display:block; cursor:pointer" />
        </td>
    </tr>
    <tr>
    	<td>
                <table border="0"   align="left" cellpadding="0"cellspacing="0" width="100%">
                    <!--   ESPACIO DE SEPARACION ENTRE EL BORDE DEL PAPEL Y LA CABECERA -->
                  <%if tienda = "07" then %>
                  <tr height="136px" valign="bottom">
                  <%else%>
                    <tr height="154px" valign="bottom">
                  <%end if %>  
                        <td width="70px">&nbsp;</td>
                        <td  align="left" style="font-size:11px"><%=RS("CLIENTE") %></td>
                        <td align="left"  style="font-size:11px"><%=left(trim(RS("NOMBRE")),65) %></td>
                        <td align="right" >&nbsp;</td>
                    </tr>
                    <tr>
                        <td >&nbsp;</td>
                        <td align="center">&nbsp;</td>
                        <td width="480px;" align="left"  style="font-size:11px" ><%=LEFT(TRIM(RS("DIRECCION")),65) %></td>
                        <td align="right" ><%=trim(rs("serie")) %> - <%=trim(rs("numdoc")) %></td>
                    </tr>
                    <tr valign="bottom" height="8px">
                        <td >&nbsp;</td>
                        <td  align="left"  style="font-size:11px" ><%=left(trim(rs("fecdoc")),10) %></td>
                        <td align="left">&nbsp;</td>
                        <td align="center">&nbsp;</td>
                    </tr>
                </table>
      </TD>
  </tr>        
        <tr>
        	<td>        
                <!--   ESPACIO DE SEPARACION ENTRE LA CABECERA Y LAS LINEAS DE DETALLE-->
              <table border="0" align="left" cellpadding="0" cellspacing="0" width="100%">
                    <tr height="20px" valign="bottom">
                        <td>&nbsp;</td>
                    </tr>
                </table>
          </TD>    
        </tr>        
        <tr>
        	<td> 
            	<!-- Detalle Boletas -->             
                <table border="0" align="left" cellpadding="0" cellspacing="0" width="100%" height="210px">
                    <% i = 1
					do while not rs.eof %>
                        <tr valign="bottom">
                          <td width="10px" align="left">	<%=UCASE(left(trim(RS("codart")),10)) %></td>
                          <td width="40px" align="center">	<%=RS("SALE") %>&nbsp;</td>
                          <td width="450px" align="left">	&nbsp;&nbsp;&nbsp;<%=TRIM(RS("uni"))%>&nbsp;&nbsp;<%=left(trim(RS("DESCRI")),62) %></td>
                          <td width="40px" align="left">	<%=round(RS("PVT"),2) %>&nbsp;</td>
                          <td width="50px" align="center">	<%if cint(rs("pordes")) > 0 then %><%=RS("pordes") %>&nbsp;%<%end if %></td>
                          <%'FAC 20130116 if left(right(rs("precio"),3),1) <> "." then precio = rs("precio")&"0" else precio = rs("precio") %>
                          <td width="120px" align="right">	<%=formatnumber(rs("PRECIO"))%>&nbsp;&nbsp;</td>
                          <%i = i + 1 %>
                  </tr>
                   		<%rs.movenext %>
                    <%loop%>
                    
                  <%for m=i to 12%>
                    	<tr valign="bottom">
                       	  <td width="10px"  align="left">&nbsp;</td>
                        	<td width="40px"  align="right">&nbsp;</td>
                        	<td width="450px" align="left">&nbsp;</td>
                        	<td width="40px"  align="right">&nbsp;</td>
                        	<td width="50px"  align="right">&nbsp;</td>
                        	<td width="120px" align="right">&nbsp;</td>
                            
                    	</tr>
                        
                    <%next %>
                    <% if  tienda="16" then %>
                    <tr><td colspan="7" style="font-size:10px"><center>No se acepta cambios ni devoluciones.</center></td></tr>
                        <%else%>
                        <tr><td colspan="7" style="font-size:10px"><center>No se acepta cambios ni devoluciones en ropa interior ni prendas con descuento, pijamas sólo cambios con documento de vta., máx 7 dias.</center></td></tr>
                    <% end if %>
                </table>
          </TD>    
        </tr>        
        <tr>
        	<td>              
                <table border="0" align="left" cellpadding="0" cellspacing="0" width="100%" >
                    <tr  valign="top">
                        <td width="10px">&nbsp;</td>
                        <td id="sonn" style="font-weight:800;font-size:14px"></td>
                    </tr>
                </table>
          </TD>    
        </tr>        
        <tr>
        	<td>               
                <table border="0" align="left" cellpadding="0"cellspacing="0" width="100%" >
                    <tr>
                        <td colspan="4" height="10px" width="460px" align="right">&nbsp;<b>ISC</b></td>
                    </tr>
                <tr valign="middle">
                        <td width="100px" align="center">	S/. &nbsp;&nbsp;&nbsp;&nbsp;<%=formatnumber(cdbl(total)+cdbl(descuento)-cdbl(isc),2)%></td>
                        <td width="120px" align="center">	S/. &nbsp;&nbsp;&nbsp;&nbsp;<%=formatnumber(descuento,2)%></td>
                   	<td width="120px" align="center">	S/. &nbsp;&nbsp;&nbsp;&nbsp;<%=formatnumber(cdbl(total)-cdbl(isc),2)%></td>
                    	<td width="150px" align="right">	ISC S/. &nbsp;&nbsp;&nbsp;&nbsp;<%=formatnumber(isc,2)%></td>
                        <td width="150px" align="right">	S/.&nbsp;&nbsp;<%=formatnumber(total,2)%></td>
                    </tr>
          </TD>    
        </tr>        

</table>


<script language="jscript" type="text/jscript">
    document.getElementById('sonn').innerText = 'Son : ' + FComson('<%=total%>')
    function printa() {
        document.all.prn.style.display = 'none'
        window.print()
        this.window.close()
    } 
</script> 

</body>
</html>
