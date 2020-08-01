<%@ Language=VBScript %>
<%TDA = Request.Cookies("TIENDA")("POS")%>
<%Response.Buffer = TRUE %>
<!--#include file="../includes/Cnn.inc"-->
<!--#include file="../comun/funcionescomunes.asp"-->
<link REL="stylesheet" TYPE="text/css" HREF="../ventas.CSS" >
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<body>

<%
'*********************** OJO **********************
'REVISAR EL LARGO DE LAS COLUMNAS de la tabla
'*********************** OJO **********************
pos = ucase(TRIM(Request.QueryString("pos")))
OP = ucase(TRIM(Request.QueryString("OP")))
marka = 0
    CAD =   " select  CODIGO, DESCRI, ISNULL(STOCK,0) AS STOCK, TIENDA INTO #TEMP from view_articulos_tienda WHERE ESTADO ='A' AND STOCK > 0; "
    'response.write(cad)
    cnn.execute cad
cad =             " select  CODIGO, DESCRI, isnull([07],0)AS  [AS],isnull([08],0) AS PO,isnull([09],0) AS CH," & _
                  "  isnull([10],0) AS SB,isnull([11],0) AS FB, isnull([12],0) as SI, isnull([13],0) as TR,  isnull([14],0) as AR,  isnull([15],0) as T2   from #temp   " & _
            " pivot (SUM(stock) for tienda in ([01],[07],[08],[09],[10],[11],[12],[13],[14],[15])) as STK       " & _
            " where codigo like '"&POS&"%'  " & _
			" order by codigo; "  ' Fac 20121229
           ' response.write(cad)
           ' response.write("<br>")
    RS.OPEN CAD,CNN

    if rs.recordcount > 0 then%>
    <table width="100%">
	<tr><td  align="center" class="estilo6">Artículos <%=ucase(destda) %></td></tr>
	<tr><td><hr /></td></tr>
</table>

<table id="TABLA" align="center"  cellpadding="0" cellspacing="2" bordercolor='<%=application("color2") %>' border="1" >
<%'**************************%>
<%'LINEA DE CABECERA STANDAR %>
<%'**************************%>
<%cont = 1
rs.movefirst %>
    <tr>
          <%for i=0 to rs.fields.count - 1 %>
           <td class="Estilo5" align="center"><%=rs.fields(i).name %></td>
           <%next %>
    </tr>
    <%Dim aCan(10) %>
    <%IF  RS.EOF THEN response.end%>
        <%DO WHILE NOT RS.EOF%>                       
            <%codigo =  rs.fields.item(0) %>
            <%descri =   rs.fields.item(1) %>
            <%do while trim(codigo) = trim(rs.fields.item(0)) and not rs.eof%>
        	        <%if not ISNULL(rs("AS")) then if cint(rs("as")) > 0 then aCan(0) = CINT(rs("AS")) ELSE ACAN(0) = 0 %>
                    <%if not ISNULL(rs("PO")) then if cint(rs("po")) > 0 then aCan(1) = CINT(rs("PO")) ELSE ACAN(1) = 0 %>
                    <%if not ISNULL(rs("CH")) then if cint(rs("ch")) > 0 then aCan(2) = CINT(rs("CH")) ELSE ACAN(2) = 0 %>
                    <%if not ISNULL(rs("SB")) then if cint(rs("sb")) > 0 then aCan(3) = CINT(rs("SB")) ELSE ACAN(3) = 0 %>
                    <%if not ISNULL(rs("FB")) then if cint(rs("fb")) > 0 then aCan(4) = CINT(rs("FB")) ELSE ACAN(4) = 0 %>
                    <%if not ISNULL(rs("si")) then if cint(rs("SI")) > 0 then aCan(5) = CINT(rs("SI")) ELSE ACAN(5) = 0 %>
                    <%if not ISNULL(rs("TR")) then if cint(rs("TR")) > 0 then aCan(6) = CINT(rs("TR")) ELSE ACAN(6) = 0 %>
                    <%if not ISNULL(rs("AR")) then if cint(rs("AR")) > 0 then aCan(7) = CINT(rs("AR")) ELSE ACAN(7) = 0 %>
                    <%if not ISNULL(rs("T2")) then if cint(rs("T2")) > 0 then aCan(8) = CINT(rs("T2")) ELSE ACAN(8) = 0 %>
                    <%rs.movenext %>
                    <%if rs.eof then exit do%>
            <%loop%>
        <%IF CINT(ACAN(0))+CINT(ACAN(1)) +CINT(ACAN(2)) + CINT(ACAN(3)) + CINT(ACAN(4)) + CINT(ACAN(5))+ CINT(ACAN(6))+ CINT(ACAN(7))+ CINT(ACAN(8)) THEN %>
        <tr  bgcolor="<% if CONT mod 2  = 0 THEN 
                response.write(Application("color2"))
                else
	            response.write(Application("color1"))
	            end IF%>"
	                id="fila<%=Trim(Cstr(cont))%>">
                <td class="Estilo5" align="center"><%=CODIGO%></td>
                <td class="Estilo5" align="center"><%=DESCRI%></td>
                <%FOR I=0 TO 8 %>
                 <td class="Estilo5" align="center"><%IF CINT(ACAN(I))>  0 THEN RESPONSE.WRITE(ACAN(i)) ELSE RESPONSE.WRITE("")%></td>
                <%NEXT %>
            <%cont =cont +1 %>
            <%ACAN(0) = 0 %>
            <%ACAN(1) = 0 %>
            <%ACAN(2) = 0 %>
            <%ACAN(3) = 0 %>
            <%ACAN(4) = 0 %>
            <%ACAN(5) = 0 %>
            <%ACAN(6) = 0 %>
            <%ACAN(7) = 0 %>
            <%ACAN(8) = 0 %>
        </tr>              
        <%END IF %>
   
        <%if rs.eof then exit do%>
<%loop%>

</table>
<%END IF%>

</BODY>
</HTML>
