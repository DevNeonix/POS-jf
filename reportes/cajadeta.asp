<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = true %>
<%Session.LCID=2058%>

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

<script language="jscript" type="text/jscript">
function AGRANDA() {
    top.parent.window.document.getElementById('body0').height = 530
}
</script>

<%tienda = request.QueryString("tda")
fec = request.QueryString("fec")
fec2 = request.QueryString("fec2")

if ucase(trim(tienda)) = "TT" then 
    cad = "exec SP_CAJA_DIA_todAS  '"&fec&"', '"&fec2&"' "
else 
    cad = "exec SP_CAJA_DIA '"&tienda&"', '"&fec&"', '"&fec2&"' "

    'IF ucase(trim(tienda)) = "01" THEN
    ''    cad = "exec SP_CAJA_DIA_REAL '"&tienda&"', '"&fec&"', '"&fec2&"' "
    'END IF

END IF
RESPONSE.WRITE (CAD)
rs.open cad,cnn
if rs.recordcount <=0 then RESPONSE.End
IF  request.QueryString("EXCEL") = "1" THEN
    archivo = "c:\temp\cajaexcel.xls"
    Response.Charset = "UTF-8"
    Response.ContentType = "application/vnd.ms-excel" 
    Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo 
END IF

%>

<body onload="AGRANDA()">
<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="Table2" name="listado"  >
<tr>
<td><input type="button" value="Excel " onclick="REPORTE(1)" /></td>


</tr>
</table>
<center>
<table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="listado" name="listado"  >
	<tr> 
   		<td align="center" class="Estilo8">Operación</td>
        <td align="center" class="Estilo8">RUC</td>
        <td align="center" class="Estilo8">Cliente</td>
        <td align="center" class="Estilo8">Fecha</td>
        <td align="center" class="Estilo8">Hora</td>
        <td align="center" class="Estilo8">Documento</td>
        <td align="center" class="Estilo8">Total<br />Docum.</td>
        <td align="center" class="Estilo8">S/.</td>
        <td align="center" class="Estilo8">$/.</td>
        <td align="center" class="Estilo8">Visa</td>
        <td align="center" class="Estilo8">Mcard</td>
        <td align="center" class="Estilo8">N/C</td>
        <td align="center" class="Estilo8">Total<br>Ingresos</td>
        <td align="center" class="Estilo8"># DOC</td>
        <td align="center" class="Estilo8">USR</td>
	</tr>
	<%I=0 
    tota = 0
    tota2 = 0
    tlndolar=0
    tlnsol=0
	tlnvis=0
	tlnmas=0
	tlnncr=0
	RS.MOVEFIRST
    
    MM=0    %>
	<%do while not rs.eof
    	OPER = RS("OPERACION")
      	MM=0 
		lnsol=0
		lnvis=0
		lnmas=0
		lnncr=0
%>
		<tr bgcolor='<%=application("color2") %>'  class="EstiloT" align="right">
        	<td>&nbsp;<%RESPONSE.WRITE(rs("operacion"))%>&nbsp;</td>
            <td>&nbsp;<%=rs("Cliente")%>&nbsp;</td>
            <td align="left">&nbsp;<%=rs("Nombre")%>&nbsp;</td>
            <td>&nbsp;<%response.write(formatdatetime(rs("fecha"),2))%>&nbsp;</td>
             <td>&nbsp;<%response.write(rs("hora"))%>&nbsp;</td>
            <td>&nbsp;<%response.write(rs("coddoc")&"-"&rs("numdoc"))%>&nbsp;</td>
            <td style="padding-right:10px">
                <%
                if trim(rs("tipo")) = "USD" then
                    response.write(formatnumber(0,2,,true))
                ELSE
                    response.write(formatnumber(rs("total"),2,,true))
                end if

                %>
            </td>
            <!-- nota empieza en blanco por si acaso hayan varios registrospara agar una misma boleta -->
          <% NOTA = "" %>
            <%do while not rs.eof AND RTRIM(LTRIM(RS("OPERACION"))) = RTRIM(LTRIM(OPER))

                if trim(rs("tipo")) = "SOL" then lnsol =  CDBL(rs("SOL"))
                if trim(rs("tipo")) = "USD" then lnDOLAR =  CDBL(rs("DOLAR"))
				if trim(rs("tipo")) = "VIS" then lnvis = lnvis + CDBL(rs("SOL"))
				if trim(rs("tipo")) = "MAS" then lnmas = lnmas + CDBL(rs("SOL"))
				if trim(rs("tipo")) = "NCR" then lnncr = lnncr + CDBL(rs("SOL"))
				if trim(rs("tipo")) = "NCR" then NOTA = rs("NOTA") else  MM = MM + CDBL(RS("SOL"))
                 user = rs("usuario")  

           
                if trim(rs("tipo")) = "USD" then
                    
                ELSE
                    tota = tota + CDBL(RS("total"))
                    
                end if
               

                tota2 = tota2 + CDBL(RS("SOL"))

                tlnsol=tlnsol+lnsol
                tlnDOLAR=tlnDOLAR+lnDOLAR
	            tlnvis=tlnvis+lnvis
	            tlnmas=tlnmas+lnmas
	            tlnncr=tlnncr+lnncr
				rs.movenext
				IF RS.EOF THEN EXIT DO
			loop%>
            <td style="padding-right:10px"><%=FORMATNUMBER(lnsol,2,,,TRUE)%></td>
            <td style="padding-right:10px"><%=FORMATNUMBER(lnDOLAR,2,,,TRUE)%></td>
			<td style="padding-right:10px"><%=FORMATNUMBER(lnvis,2,,,TRUE)%></td>
			<td style="padding-right:10px"><%=FORMATNUMBER(lnmas,2,,,TRUE)%></td>
			<td style="padding-right:10px"><%=FORMATNUMBER(lnncr,2,,,TRUE)%></td>
			<td style="padding-right:10px"><%=FORMATNUMBER(MM,   2,,,true)%></td>
            <td>&nbsp;<%=NOTA%>&nbsp;</td>
            <td>&nbsp;<%=user%>&nbsp;</td>
            <td>&nbsp;<%=tota%>&nbsp;</td>
		</tr> 
        <%I= I + 1%>
 
	
        <%IF RS.EOF THEN EXIT DO %>
	<%loop %>

    
    <tr bgcolor= "<%=application("barra")%>" class="Estilo3" align="right">
        <td colspan="6">&nbsp;Total Ingresos: S/: <%=fec%>  al <%=fec2%> &nbsp;</td>
        <td  style="padding-right:10px"><%=formatnumber(tota,2,,true)%></td>
        <%RS.CLOSE
if ucase(trim(tienda)) = "TT" then     CAD = "exec SP_CAJA_DIA_TOT_TODO  '"&fec&"', '"&fec2&"', 'SOL' " else     CAD = "exec SP_CAJA_DIA_TOT '"&tienda&"', '"&fec&"', '"&fec2&"', 'SOL' " 
'RESPONSE.WRITE(CAD)
' RESPONSE.WRITE(CAD)
    RS.OPEN CAD,CNN
    IF RS.RECORDCOUNT > 0 THEN 
        RS.MOVEFIRST  
        tlnsol = RS("TOT") 
    ELSE 
        tlnsol= 0 
        
    END IF%>

        <td style="padding-right:10px" ><%=formatnumber(tlnsol   ,2,,true) %></td>
        <%RS.CLOSE
        if ucase(trim(tienda)) = "TT" then     CAD = "exec SP_CAJA_DIA_TOT_TODO  '"&fec&"', '"&fec2&"', 'USD' " else     CAD = "exec SP_CAJA_DIA_TOT '"&tienda&"', '"&fec&"', '"&fec2&"', 'USD' " 
        'RESPONSE.WRITE(CAD)
        ' RESPONSE.WRITE(CAD)
            RS.OPEN CAD,CNN
            IF RS.RECORDCOUNT > 0 THEN 
                RS.MOVEFIRST  
                tlnDOLAR = RS("TOT") 
            ELSE 
                tlnDOLAR= 0 
                
            END IF%>
        <td style="padding-right:10px" ><%=formatnumber(tlnDOLAR   ,2,,true) %></td>


         <%RS.CLOSE
  if ucase(trim(tienda)) = "TT" then     CAD = "exec SP_CAJA_DIA_TOT_TODO  '"&fec&"', '"&fec2&"', 'VIS' " else     CAD = "exec SP_CAJA_DIA_TOT '"&tienda&"', '"&fec&"', '"&fec2&"', 'VIS' " 
  'RESPONSE.WRITE(CAD)
    RS.OPEN CAD,CNN
    IF RS.RECORDCOUNT > 0 THEN 
        RS.MOVEFIRST  
        tlnvis = RS("TOT") 
    ELSE 
        tlnvis= 0 
    END IF%>
        <td  style="padding-right:10px"><%=formatnumber(tlnvis   ,2,,true) %></td>
         <%RS.CLOSE
 if ucase(trim(tienda)) = "TT" then     CAD = "exec SP_CAJA_DIA_TOT_TODO  '"&fec&"', '"&fec2&"', 'MAS' " else     CAD = "exec SP_CAJA_DIA_TOT '"&tienda&"', '"&fec&"', '"&fec2&"', 'MAS' " 
    RS.OPEN CAD,CNN
    IF RS.RECORDCOUNT > 0 THEN 
        RS.MOVEFIRST  
        tlnmas = RS("TOT") 
    ELSE 
        tlnmas= 0 
    END IF%>
        <td  style="padding-right:10px"><%=formatnumber(tlnmas   ,2,,true) %></td>
        <%RS.CLOSE
   if ucase(trim(tienda)) = "TT" then     CAD = "exec SP_CAJA_DIA_TOT_TODO  '"&fec&"', '"&fec2&"', 'NCR' " else     CAD = "exec SP_CAJA_DIA_TOT '"&tienda&"', '"&fec&"', '"&fec2&"', 'NCR' " 
    RS.OPEN CAD,CNN
    IF RS.RECORDCOUNT > 0 THEN 
        RS.MOVEFIRST  
        tlnncr = cdbl(RS("TOT") )
    ELSE 
        tlnncr= 0 
    END IF%>
        <td  style="padding-right:10px"><%=formatnumber(tlnncr   ,2,,true) %></td>
        <td  style="padding-right:10px"><%=formatnumber(cdbl(tota2)-cdbl(tlnncr)   ,2,,true) %></td>
        <td>&nbsp;</td>

	</tr>

</table>
<p>&nbsp;</p>
<table align="center" cellpadding="2" cellspacing="0" bordercolor="<%=application("color1")%>"  >
    <tr>
        <!-- tabla de cuenta para comision -->
        <td width="50%">
            
            <table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="Table1" name="listado"  >
               <tr> <td align="center" class="Estilo8" colspan="3">RESUMEN DOCUMENTARIO:</td>   </tr>
               <tr> <td align="center" class="Estilo8">Tipo</td>  
                <td align="center" class="Estilo8">Soles</td>  
                <td align="center" class="Estilo8">Dolares</td>   </tr>

            <tr bgcolor='<%=application("color2") %>' > 
                <td align="left" class="EstiloT">SOLES</td>  
                <td align="right" class="EstiloT" style="padding-right:10px"><%=formatnumber(TLNSOL   ,2,,true) %></td>  
                <td align="center" class="EstiloT">&nbsp;</td>  
            </tr>
            <tr bgcolor='<%=application("color2") %>' > 
                <td align="left" class="EstiloT">DOLARES</td>  
                
                <td align="center" class="EstiloT">0.00</td>  
                <td align="right" class="EstiloT" style="padding-right:10px"><%=formatnumber(TLNDOLAR   ,2,,true) %></td>  
            </tr>
            <tr bgcolor='<%=application("color2") %>' > 
                <td align="left" class="EstiloT">VISA</td>  
                <td align="right" class="EstiloT" style="padding-right:10px"><%=formatnumber(TLNVIS   ,2,,true) %></td>  
                <td align="center" class="EstiloT">&nbsp;</td>  
            </tr>
            <tr bgcolor='<%=application("color2") %>' > 
                <td align="left" class="EstiloT">MASTER</td>  
                <td align="right" class="EstiloT" style="padding-right:10px"><%=formatnumber(TLNMAS   ,2,,true) %></td>  
                <td align="center" class="EstiloT">&nbsp;</td>  
            </tr>
             <tr bgcolor='<%=application("color2") %>' > 
                <td align="left" class="EstiloT">N. CRE.</td>  
                <td align="right" class="EstiloT" style="padding-right:10px"><%=formatnumber((TLNNCR*-1)  ,2,,true) %></td>  
                <td align="center" class="EstiloT">&nbsp;</td>  
            </tr>
            <tr> 
                <td align="left" class="Estilo8">Total S/.</td>  
                <td align="right" class="Estilo8" style="padding-right:10px"><%=formatnumber(cdbl(tota) - cdbl(TLNNCR)    ,2,,true) %></td> 
                
                <td align="center" class="Estilo8"><%=formatnumber((TLNDOLAR)  ,2,,true) %></td>  
            </tr>
            </table>


        </td>
        <!-- tabla de resumen completo -->
        <td width="50%">
            
            <table align="center" cellpadding="2" cellspacing="0" bordercolor='<%=application("color1") %>' border="1" id="Table3" name="listado"  >
               <tr> <td align="center" class="Estilo8" colspan="3">RESUMEN Comisiones:</td>   </tr>
               <tr> <td align="center" class="Estilo8">Tipo</td>  
                <td align="center" class="Estilo8">Soles</td>  
                <td align="center" class="Estilo8">Dolares</td>   </tr>

            <tr bgcolor='<%=application("color2") %>' > 
                <td align="left" class="EstiloT">SOLES</td>  
                <td align="right" class="EstiloT" style="padding-right:10px"><%=formatnumber(TLNSOL   ,2,,true) %></td>  
                <td align="center" class="EstiloT">&nbsp;</td>  
            </tr>
            <tr bgcolor='<%=application("color2") %>' > 
                <td align="left" class="EstiloT">DOLAR</td>  
                
                <td align="center" class="EstiloT">0.00</td>  
                <td align="right" class="EstiloT" style="padding-right:10px"><%=formatnumber(TLNDOLAR   ,2,,true) %></td>  
            </tr>
            <tr bgcolor='<%=application("color2") %>' > 
                <td align="left" class="EstiloT">VISA</td>  
                <td align="right" class="EstiloT" style="padding-right:10px"><%=formatnumber(TLNVIS   ,2,,true) %></td>  
                <td align="center" class="EstiloT">&nbsp;</td>  
            </tr>
            <tr bgcolor='<%=application("color2") %>' > 
                <td align="left" class="EstiloT">MASTER</td>  
                <td align="right" class="EstiloT" style="padding-right:10px"><%=formatnumber(TLNMAS   ,2,,true) %></td>  
                <td align="center" class="EstiloT">&nbsp;</td>  
            </tr>
             <tr bgcolor='<%=application("color2") %>' > 
                <td align="left" class="EstiloT">N. CRE.</td>  
                <td align="right" class="EstiloT" style="padding-right:10px"></td>  
                <td align="center" class="EstiloT">&nbsp;</td>  
            </tr>
            <tr> 
                <td align="left" class="Estilo8">Total S/.</td>  
                <td align="right" class="Estilo8" style="padding-right:10px"><%=formatnumber((cdbl(TLNSOL)+cdbl(TLNVIS)+cdbl(TLNMAS))   ,2,,true) %></td>
                <!--<td align="right" class="Estilo8" style="padding-right:10px"><%=formatnumber(cdbl(tota) - cdbl(TLNNCR)   ,2,,true) %></td>  -->
                <td align="center" class="Estilo8"><%=formatnumber((TLNDOLAR)  ,2,,true) %></td>  
            </tr>
</table>


        </td>
    </tr>
</table>

</center>
</body>
<script language="jscript" type="text/jscript">
    function REPORTE(op) {
        if (op == '1')
            window.open('cajaDETA.asp?pos=' + '<%=tienda %>' + '&fec=' + '<%=fec %>' + '&fec2=' + '<%=fec2%>'+'&tda='+'<%=tienda%>'+'&EXCEL=1')
    }
</script>
</html>
