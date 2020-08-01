<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<% tienda = Request.Cookies("tienda")("pos") %>
<% destda = Request.Cookies("tienda")("tda") %>

<script type="text/jscript" language="jscript">
var oldrow =1
var olddata =''
var chk = ''
</script>

<%' Definir el tamaño de la pagina
Dim pagesize 
pagesize =100
'*********************************************
'     Definir el TITULO de la PAGINA ASP 
'*********************************************
TITULO = Request.QueryString("tit")
pos = Request.QueryString("pos")
%>

<script type="text/jscript" language="jscript">
function dd2(ff)
{	// LLENA TEXTBOX ADICIONALES AL COMUN
	// LOS DEL COMUN SON CODIGO Y DESCRIPCION
	var t = document.all.TABLA;
	var pos = parseInt(ff) ;
	dd(ff,0);
}
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=titulo%></title>
</head>

<body>

<form name="thisForm" id="thisForm" method="post" action="">

<table width="100%">
	<tr><td colspan="3"><hr /></td></tr>
    <tr><td align="center" class="Estilo6"><%=titulo%></td></tr>

</table>
<!--#include file="includes/Cnn.inc"-->
<!--#include file="COMUN/FUNCIONESCOMUNES.ASP"-->
<!--#include file="COMUN/COMUNqry.ASP"-->

<%
marka = 0
if len(trim(pos)) <=0 then 
    POS = Request.Form("pos")
end if

if pos = "" or isnull(pos)  or pos = " " then 	
    pos = ""
end if

if len(pos)>0 then
    cad =   " select docori, numori, tienda from movimcab   " & _
            " where docori = 'GS' and numori = '"&pos&"' ;  "     
    RS.OPEN CAD,CNN
    if rs.recordcount > 0 then 
        marka = 1
        rs.movefirst
        tendero = rs("tienda")
    end if
    rs.close
end if

'response.end
TOTA =  " select sum(c6_ncanten) as total                       " & _
        " from rsfaccar..al0012movd                             " & _
        " where c6_ctd = 'gs'                                   " & _
        " and c6_cnumdoc = '"&pos&"' and c6_ccodigo!='TXT'      " 
rs.open tota,cnn
tot = rs("total")
rs.close

'Fac 20121213
CAD =  "select c6_citem,c6_ccodigo,ar_cdescri as descri,c6_ncanten,ar_npreci6 " & _
       "from rsfaccar..al0012movd inner join RSFACCAR.dbo.AL0012arti" & _
	   " on ar_ccodigo=c6_ccodigo "& _
	   "where c6_ctd = 'GS' and c6_cnumdoc = '"&pos&"' and c6_ccodigo!='TXT'  " '& _
'	   " order by c6_ccodigo"	' Fac 20121229
	   
RS.Open CAD, Cnn
SINDATOS =1
CONT = 1
IF  RS.RECORDCOUNT > 0 THEN 	
	RS.MOVEFIRST
ELSE
	'RESPONSE.Write("NO HAY DATOS DE LA GUIA SOLICITADA")	
	SINDATOS =0
	'RESPONSE.End()
END IF	
%>

<table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="1"  cellspacing="1"  border="1" >
    <tr>
        <td width="25%" valign="top">
            <table id="Table2" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="0"  cellspacing="0"  border="0" width="100%">
                <tr align="center">
                    <td class="Estilo6">Nº de Guia: </td>
                </tr>
                <tr>
                    <td class="Estilo5"><input type="text" id="POS" name="POS" value="<%=pos%>" /></td>
                </tr>
                <tr align="center">
                    <td ><%if  pos = "" or isnull(pos)  or pos = " " then %><input type="submit" id="ENV" name="ENV" value="MOSTRAR" /><%end if %> </td>
                </tr>
	</tr>
    <tr align="center" >
    	<td class="Estilo6" height="40"><hr /></td>
        </tr>
                <tr align="center">
                    <td class="Estilo6"><%if  len(pos)>0 then %>TOTAL PRENDAS:<%end if %> </td>
                </tr>
                <tr align="center">
                    <td ><%if  len(pos)>0 then %><input type="text" id="tota" name="tota" value="<%=tot%>" /><%end if %> </td>
                </tr>
                 <tr align="center" class="Estilo8">
                    <td ><%if  len(pos) > 0 and cint(marka) = 0 then %><input type="button" id="STK" name="STK" value="CARGAR STOCK" onClick="cargar()" />
                    <%else %>
                        <%if cint(marka) = 1 then %>Guia ya registrada<br /> en tienda - <%=tendero%><% end if %>
                    <%end if %> </td>
                </tr>
            </table>
</td>
<td width="75%" rowspan ="8" valign="top">
<%' Nro de columnas regresadas por el objeto RECORDSET	
columnas = rs.Fields.Count %>
<%'*********************************************************************%>
<table id="TABLA" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("barra")%>"  cellpadding="1"  cellspacing="1"  border="1" width="100%">
<%'**************************%>
<%'LINEA DE CABECERA STANDAR %>
<%'**************************%>
<tr>
<%LIMITE =  rs.fields.count - 1%>
    <td align="center" class="Estilo8">IT</td>
	<td align="center" class="Estilo8">Codigo</td>
    <td align="center" class="Estilo8">Descripcion</td>
    <td align="center" class="Estilo8">Cant.</td>
    <td align="center" class="Estilo8">Precio</td>
</tr>
<%'*****************************%>
<%' MUESTRA EL GRid (2 colorES) %>
<%'*****************************%>
<%cont=1 %>
<%IF NOT RS.EOF THEN%>
<%DO WHILE NOT RS.EOF%>
    <tr  bgcolor="<% if CONT mod 2  = 0 THEN 
                response.write(Application("color1"))
                else
	            response.write(Application("color2"))
	            end IF%>"
	            onclick="dd('<%=(cont)%>',0)" id="fila<%=Trim(Cstr(cont))%>">
      
	    <td class="Estilo0" align="center"><%=CINT(rs("c6_citem")) %></td>
	    <td class="Estilo0" align="center"><input type="text" value='<%=rs("c6_ccodigo")%>' class="Estilo0" id="COD<%=CONT%>" name ="COD<%=CONT%>" readonly tabindex="-1" size="12"/></td>
        <td class="Estilo0" width="65%"> <%=rs("descri") %> </td>
        <td class="Estilo0" align="center" width="2%"><input type="text" value='<%=rs("c6_ncanten")%>' class="Estilo0" id="CAN<%=CONT%>" name ="CAN<%=CONT%>" maxlength="3" size="4"/></td>
	    <td class="Estilo0"><input type="text" value='<%=rs("ar_npreci6")%>' class="Estilo0" id="PRE<%=CONT%>" name="PRE<%=CONT%>" maxlength="6" size="4" align="right" readonly/> <%'=formatnumber(rs("ar_npreci6"),2) & ".-  "%></td>
    </tr>  
   	<%RS.MOVENEXT%>
	<%CONT = CONT + 1%>
<%loop%>
</table>
</td>
</tr>
</table>
<%END IF %>

<script type="text/jscript" language="jscript">
rec = parseInt('<%=rs.recordcount%>',10)
if (rec > 0 )
	dd2('1');

function llena()
{ return true }

function cargar() {
//alert("marka ="+'<%=marka %>')
// el recordcount me devuelve el total de lineas empezando en 1...
lin = parseInt('<%=rs.recordcount %>', 10);
//alert(lin);
if (lin > 0) {
	//alert("Procederemos a crearlos...");
	thisForm.action = ('comun/inserguia.asp?lin=' + lin);
	thisForm.submit();
    //window.open('comun/inserguia.asp?lin=' + lin);
	}
else
	alert("No hay nada que ingresar");
}
</script>

<%	RS.Close 
	SET RS  = NOTHING
	Cnn.Close
	SET Cnn = NOTHING
%>

</form>
</body>
</html>
