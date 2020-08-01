<%@ Language=VBScript %>
<%TDA = Request.Cookies("TIENDA")("POS")%>
<%Response.Buffer = TRUE %>
<!--#include file="includes/Cnn.inc"-->
<!--#include file="comun/funcionescomunes.asp"-->
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
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
    CAD =   " select  CODIGO, DESCRI, ISNULL(MINIMO,0) AS UME, TIENDA INTO #TEMP from view_articulos_tienda WHERE ESTADO ='A' AND STOCK > 0; "
    'response.write(cad)
    cnn.execute cad
cad =             " select  CODIGO, DESCRI, [07] AS [AS],[08] AS PO, [09] AS CH, [10] AS SB, [11] AS FB, [12] as SI, [13] as TR from #temp   " & _
            " pivot (SUM(UME) for tienda in ([01],[07],[08],[09],[10],[11],[12],[13])) as MIN       " & _
			" order by codigo; "
            'response.write(cad)
           ' response.write("<br>")
    RS.OPEN CAD,CNN
letras = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
'RESPONSE.WRITE(MID(LETRAS,1,1))
    if rs.recordcount > 0 then%>
    <table width="100%">
	<tr><td  align="center" class="estilo6">Art&iacute;culos <%=ucase(destda) %></td></tr>
	<tr><td><hr /></td></tr>
</table>

<table id="TABLA" align="center"  cellpadding="0" cellspacing="2" bordercolor='<%=application("color2") %>' border="1" >
<%'**************************%>
<%'LINEA DE CABECERA STANDAR %>
<%'**************************%>
<%cont = 1
rs.movefirst %>
    <tr> 
        <td class="Estilo5" align="center">Lin</td>
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
        	        <%if not ISNULL(rs("AS")) then if cint(rs("as")) > 0 then aCan(0) = CINT(rs("AS")) ELSE ACAN(0) = 0%>
                    <%if not ISNULL(rs("PO")) then if cint(rs("po")) > 0 then aCan(1) = CINT(rs("PO")) ELSE ACAN(1) = 0%>
                    <%if not ISNULL(rs("CH")) then if cint(rs("ch")) > 0 then aCan(2) = CINT(rs("CH")) ELSE ACAN(2) = 0%>
                    <%if not ISNULL(rs("SB")) then if cint(rs("sb")) > 0 then aCan(3) = CINT(rs("SB")) ELSE ACAN(3) = 0%>
                    <%if not ISNULL(rs("FB")) then if cint(rs("fb")) > 0 then aCan(4) = CINT(rs("FB")) ELSE ACAN(4) = 0%>
                    <%if not ISNULL(rs("SI")) then if cint(rs("SI")) > 0 then aCan(5) = CINT(rs("SI")) ELSE ACAN(5) = 0%>
                    <%if not ISNULL(rs("TR")) then if cint(rs("TR")) > 0 then aCan(6) = CINT(rs("TR")) ELSE ACAN(6) = 0%>
                    <%rs.movenext%>
                    <%if rs.eof then exit do%>
            <%loop%>
        <%IF CINT(ACAN(0))+CINT(ACAN(1)) +CINT(ACAN(2)) + CINT(ACAN(3)) + CINT(ACAN(4)) > 0 THEN %>
        <tr  bgcolor="<% if CONT mod 2  = 0 THEN 
                response.write(Application("color2"))
                else
	            response.write(Application("color1"))
	            end IF%>"
	                id="fila<%=Trim(Cstr(cont))%>">
                <td class="Estilo5" align="center"  width="2%"><%=cont%></td>
                <td class="Estilo5" align="center"  width="5%"><%=CODIGO%></td>
                <td class="Estilo5" align="left"  width="25%"><%=DESCRI%></td>
                <%FOR I=0 TO 6 %>
                 <td class="Estilo5" align="center"  width="5%">
                    <%IF CINT(ACAN(I))> 0 THEN%> 
                      <input id="<%=MID(LETRAS,I+1,1)&CONT%>"   value="<%=aCan(i)%>" class="Estilo1"  />
                    <% ELSE 
                        RESPONSE.WRITE("")
                    end if%>
                 </td>
                <%NEXT %>
            <%cont =cont +1 %>
            <%ACAN(0) = 0 %>
            <%ACAN(1) = 0 %>
            <%ACAN(2) = 0 %>
            <%ACAN(3) = 0 %>
            <%ACAN(4) = 0 %>
            <%ACAN(5) = 0 %>
            <%ACAN(6) = 0 %>
        </tr>              
        <%END IF %>
        <%if rs.eof then exit do%>
<%loop%>

</table>
<%END IF%>
<script language="javascript" type="text/jscript">



    if (!document.getElementById('A1'))
        if (!document.getElementById('B1'))
            if (!document.getElementById('C1'))
                if (!document.getElementById('D1'))
                    if (!document.getElementById('E1'))
                        if (!document.getElementById('F1'))
                            seleccionar(document.getElementById('G1'))
                        else
                            seleccionar(document.getElementById('F1'))
                    else
                        seleccionar(document.getElementById('E1'))
                else
                   seleccionar(document.getElementById('D1'))
            else
                seleccionar(document.getElementById('C1'))
        else
            seleccionar(document.getElementById('B1'))
    else
        seleccionar(document.getElementById('A1'))
  
</script>
</BODY>
</HTML>
