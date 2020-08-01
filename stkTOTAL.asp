<%@ Language=VBScript %>
<%TDA = Request.Cookies("TIENDA")("POS")%>
<%Response.Buffer = FALSE %>
<%IF  request.QueryString("EXCEL") = "1" THEN
  archivo = "c:\temp\stkexcel.xls"
    Response.Charset = "UTF-8"
    Response.ContentType = "application/vnd.ms-excel" 
    Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo %>
<%else %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%END IF%>
<!--#include file="includes/Cnn.inc"-->
<!--#include file="comun/funcionescomunes.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<style>

body {
	/*	background-image: url(imagenes/fondo.jpg);*/
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	color: #C0C0C0;
}

.EstiloT   {font-family:Arial;   font-size:10px;  color:#003366; font-weight:300; border:hidden; background-color:transparent;}
.Estilo0   {font-family:Arial;   font-size:10px;  color:#003366; font-weight:300; border:0px;}
.Estilo1   {font-family:Arial;   font-size:12px;  color:Gray;    font-weight:600}
.Estilo2   {font-family:Arial;   font-size:12px;  color:Teal;    width:100%;}
.Estilo3   {font-family:Tahoma;  font-size:11px;  color:#ffffff; font-weight:200; background-color:#c82f8a; }
.Estilo4   {font-family:Arial;   font-size:12px;  color:Teal;    font-weight:300; border:hidden; padding: 0.4em 0.6em;}
.Estilo5   {font-family:Arial;   font-size:10px;  color:#003366; font-weight:300; border:hidden; padding: 0.1em 0.1em;}
.Estilo6   {font-family:Tahoma;  font-size:15px;  color:#F09;    font-weight:600; background-color:#ffffff;  padding: 0em 0em; }
.Estilo7   {font-family:Tahoma;  font-size:12px;  color:#999;    font-weight:600; background-color:#ffffff; padding: 0em 0em; }
.Estilo8   {font-family:Tahoma;  font-size:12px;  color:#ffffff; font-weight:600; background-color:#c82f8a; }
.Estilo9   {font-family:Tahoma;  font-size:12px;  color:#ffffff; font-weight:600; background-color:#f599c0; }
.Estilo10  {font-family:Arial;   font-size:10px;  color:#003366; font-weight:300; background-color:transparent;}
.Estilo11  {font-family:Tahoma;  font-size:12px;  color:#F09;    font-weight:300; background-color:#ffffff;  padding: 0em 0em; }
.Estilo12  {font-family:Arial;   font-size:12px;  color:Teal;    font-weight:300; padding: 0em 0em;}
.Estilo13  {font-family:Arial;   font-size:12px;  color:Teal;    font-weight:200; border:hidden;}
.Estilo14  {font-family:Tahoma;  font-size:11px;  color:Teal;    font-weight:200; text-align:right; padding: 0em 0.6em; background-color:transparent; }
.Estilo15  {font-family:Tahoma;  font-size:30px;  color:#F09;    font-weight:800; background-color:transparent; text-align:right; padding: 0em 0.6em; border:none }
.Estilo16  {font-family:Tahoma;  font-size:30px;  color:Gray;    font-weight:800; background-color:transparent; text-align:right; padding: 0em 0.6em; border:thin }
.Estilo17  {font-family:Tahoma;  font-size:20px;  color:teal;    font-weight:600; background-color:#ffffff;  }
.Estilo18  {font-family:Tahoma;  font-size:20px;  color:#ffffff; font-weight:600; background-color:#c82f8a; text-align:right; padding: 0em 0em; width:150px; }
.Estilo19  {font-family:Tahoma;  font-size:10px;  color:Teal;    font-weight:300; padding: 0em 0em;}
.Estilo20  {font-family:Tahoma;  font-size:10px;  color:Purple;  font-weight:400; padding:0.0em 0.6em 0.0em 0.0em; text-align:right;}
.Estilo21  {font-family:Tahoma;  font-size:10px;  color:Teal;    font-weight:300; border:hidden; background-color:transparent; text-align:center;}
.Estilo22  {font-family:Tahoma;  font-size:10px;  color:Teal;    font-weight:300; background-color:transparent; text-align:center;}
.Estilo23  {font-family:Arial;   font-size:10px;  color:red;     font-weight:300; padding: 0.2em 0.2em;}
.Estilo24  {font-family:Arial;   font-size:12px;  color:Teal;    font-weight:200; background-color:#F8D3ED;}
.Estilo25  {font-family:Tahoma;  font-size:11px;  color:#000000; font-weight:300; background-color:#f599c0; }         
.Estilo133 {font-family:Arial;   font-size:12px;  color:Teal;    font-weight:200; text-align:right; }
.Estilo555 {font-family:Arial;   font-size:10px;  color:Red;     font-weight:300; border:hidden; padding: 0.1em 0.1em;}
.login     {font-family:Tahoma;  font-size:12px;  color:#999;    font-weight:600; background-color:#ffffff; width:100%; }



</style>
<title><%=titulo%></title>
<link rel="stylesheet" type="text/css" href="ventas.CSS" />
<%TDA= REQUEST.QueryString("TDA") %>
</head>

<body topmargin="0" leftmargin="0" rightmargin="0" border="0">

<%
    CAD = " EXEC SP_STOCKS "
            
    ' response.write(CAD)
    RS.OPEN CAD,CNN
    'response.write(RS.RECORDCOUNT)
if rs.recordcount > 0 then     %>
  
<table id="Table2" align="center"  cellpadding="0" cellspacing="0" 
bordercolor='<%=application("color2") %>' border="1"  width="820px" >

<%'**************************%>
<%'LINEA DE CABECERA STANDAR %>
<%'**************************%>
  <%if request.QueryString("EXCEL") = "1" THEN %>  
  <tr style="display:block">
    <%else%>
    <tr style="display:none">
    <%end if %>
     <td class="Estilo5" align="left" width="70px" >CODIGO</td>
     <td class="Estilo5" align="center" width="300px">DESCRIPCION</td>
     <td class="Estilo5" align="center" width="60px" >AS</td>    
     <td class="Estilo5" align="center" width="60px" >PO</td>
     <td class="Estilo5" align="center" width="60px" >CH</td>
     <td class="Estilo5" align="center" width="60px" >SI</td>
     <td class="Estilo5" align="center" width="60px" >AR</td>
     <td class="Estilo5" align="center" width="60px" >T2</td>
     <td class="Estilo5" align="center" width="60px" >OUT</td>
     <td class="Estilo5" align="center" width="60px" >EM</td>
     <td class="Estilo5" align="center" width="60px" >JF</td>    
    </tr>


<%cont = 1
rs.movefirst %>
 
    <%Dim aCan(10) %>
    <%IF  RS.EOF THEN response.end%>
        <%DO WHILE NOT RS.EOF%>                       
            <%codigo =  rs.fields.item(0) %>
            <%descri =   rs.fields.item(1) %>
            <%do while trim(codigo) = trim(rs.fields.item(0)) and not rs.eof%>
        	        <%if not ISNULL(rs("AS")) then if cint(rs("as")) > 0   then aCan(0) = CINT(rs("AS"))  ELSE ACAN(0) = 0 %>
                    <%if not ISNULL(rs("PO")) then if cint(rs("po")) > 0   then aCan(1) = CINT(rs("PO"))  ELSE ACAN(1) = 0 %>
                    <%if not ISNULL(rs("CH")) then if cint(rs("ch")) > 0   then aCan(2) = CINT(rs("CH"))  ELSE ACAN(2) = 0 %>
                    <%if not ISNULL(rs("si")) then if cint(rs("SI")) > 0   then aCan(3) = CINT(rs("SI"))  ELSE ACAN(3) = 0 %>
                    <%if not ISNULL(rs("AR")) then if cint(rs("AR")) > 0   then aCan(4) = CINT(rs("AR"))  ELSE ACAN(4) = 0 %>
                    <%if not ISNULL(rs("T2")) then if cint(rs("T2")) > 0   then aCan(5) = CINT(rs("T2"))  ELSE ACAN(5) = 0 %>
                    <%if not ISNULL(rs("OUT")) then if cint(rs("OUT")) > 0 then aCan(6) = CINT(rs("OUT")) ELSE ACAN(6) = 0 %>
                    <%if not ISNULL(rs("EM")) then if cint(rs("EM")) > 0   then aCan(7) = CINT(rs("EM"))  ELSE ACAN(7) = 0 %>
                    <%if not ISNULL(rs("JF")) then if cint(rs("JF")) > 0   then aCan(8) = CINT(rs("JF"))  ELSE ACAN(8) = 0 %>
                    <%rs.movenext %>
                    <%if rs.eof then exit do%>
            <%loop%>
            
        <%QTY =  CINT(ACAN(7))+ CINT(ACAN(8))  
         %>
        <tr  bgcolor="<% if CONT mod 2  = 0 THEN 
                response.write(Application("color2"))
                else
	            response.write(Application("color1"))
	            end IF%>"
	                id="fila<%=Trim(Cstr(cont))%>">
                <td class="Estilo5" style="width:50px;" align="center"><%=CODIGO%></td>
                <td class="Estilo5" style="width:250px;" align="left"><%=DESCRI%></td>
                <%FOR I=0 TO 8 %>
                 <%IF  CINT(QTY) > 8  THEN %>
                     <td class="Estilo5" style="width:50px;" align="center">
                 <%else %>
                     <td class="Estilo555" style="width:50px;color:red" align="center">
                 <%end if %> 
                 <%IF CINT(ACAN(I))>  0 THEN RESPONSE.WRITE(ACAN(i)) ELSE RESPONSE.WRITE("")%></td>
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
        <%' END IF %>
   
        <%if rs.eof then exit do%>
<%loop%>

</table>
<%END IF%>

</BODY>
</HTML>
