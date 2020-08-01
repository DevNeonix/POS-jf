<%@ Language=VBScript%>
<%Response.Buffer=True%>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos")%>
<%destda = Request.Cookies("tienda")("tda")%>

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
cradio = Request.QueryString("crad")
marka= 0
%>

<% 'response.Write("entro todito")
'response.End()%>

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
<!--#include file="includes/Cnn.inc"-->
<!--#include file="COMUN/FUNCIONESCOMUNES.ASP"-->
<!--#include file="COMUN/COMUNqry.ASP"-->

<%
if len(trim(pos)) <=0 then 
    POS = Request.Form("pos")
end if
if pos = "" or isnull(pos) or pos = " " then
    pos = ""
end if

if len(trim(cradio))<=0 then	'FAC 20130326
	cradio = Request.Form("Radio1") 'FAC 20130319
end if 
'response.write(cradio)

checked1= ""
checked2= ""
checked3= ""

select case cradio
	case "Rop1"
		chk01="checked"
	case "Rop2"
		chk02="checked"
	case "Rop3"
		chk03="checked"
end select
if cradio="" then chk01="Checked"

if len(pos)>0 then
	xpos=right("00000000000"&clng(pos),11)	'FAC 20130319 solo debe recibir numeros
	'response.Write(xpos&" "&calma)
	' c11 al0012movc y c14 movimcab

	select case cradio 'FAC 20130320 Request.Form("Radio1")	'FAC 20130318
	case "Rop1"
			cwhere = " where docori = 'GS' and numori = '"&pos&"' ; "
	case "Rop2"
			pos=xpos
			cwhere = " where tienda='"&tienda&"' and docori = 'PE' and numori = '"&pos&"' ; "
	case "Rop3"
			pos=xpos
			cwhere = " where tienda='"&tienda&"' and docori = 'PS' and numori = '"&pos&"' ; "
	end select
'		cad = " select docori, numori, tienda from movimcab   " & _
'			" where docori = 'GS' and numori = '"&pos&"' ;  "     
	cad = " select docori, numori, tienda from movimcab " & cwhere
	'response.Write(cad)
    RS.OPEN CAD,CNN
    if rs.recordcount > 0 then 
        marka = 1
        rs.movefirst
        tendero = rs("tienda")
    end if
    rs.close
	'response.Write(marka)
	
'	if cint(tienda)>9 then	'FAC 20130318 para los partes del RealSystem
'		calma="T"&cstr(cint(tienda))
'	else
		calma=right("00T"&cstr(cint(tienda)),4)
'	end if
	'response.Write(calma)
	
	select case cradio	'FAc 20130321 Request.Form("Radio1")	' FAC 20130318
	case "Rop1"
			cwhere = " where c6_ctd = 'GS' and c6_cnumdoc = '"&pos&"' and c6_ccodigo!='TXT'" 
	case "Rop2"
			cwhere = " where c6_calma='"&calma&"' and c6_ctd = 'PE' and c6_cnumdoc = '"&pos&"' and c6_ccodigo!='TXT'" 
	case "Rop3"
			cwhere = " where c6_calma='"&calma&"' and c6_ctd = 'PS' and c6_cnumdoc = '"&pos&"' and c6_ccodigo!='TXT'" 
	end select
	
	TOTA =  " select sum(c6_ncantid) as total            " & _
			" from rsfaccar..al0012movd                             " & _
			cwhere
			' FAC 20130402 lo saco pq en ambos casos si tienen cantidad en id ---sum(c6_ncanten+
'			" where c6_ctd = 'gs'                                   " & _
'			" and c6_cnumdoc = '"&pos&"' and c6_ccodigo!='TXT'      " 
	'response.Write(tota)
	rs.open tota,cnn
	tot = rs("total")
	rs.close

	'Fac 20121213, 20130320 adicionar campo de partes
	CAD = "select c6_citem,c6_ccodigo,ar_cdescri as descri,c6_ncantid as c6_ncanten,ar_npreci6 " & _
		"from rsfaccar..al0012movd inner join RSFACCAR.dbo.AL0012arti on ar_ccodigo=c6_ccodigo "& _
		cwhere
		'c6_ncanten+
		'FAC 20130319 "where c6_ctd = 'GS' and c6_cnumdoc = '"&pos&"' and c6_ccodigo!='TXT'  " '& _
		'	   " order by c6_ccodigo"	' Fac 20121229
	'response.Write(cad)
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

end if
' response.end
%>

<form name="thisForm" id="thisForm" method="post" action="" >
<table width="100%">
	<tr><td colspan="3"><hr /></td></tr>
    <tr><td align="center" class="Estilo6"><%=titulo%></td></tr>
</table>

<table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="1"  cellspacing="1"  border="1" >
<tr>
	<td width="25%" valign="top">
		<table id="Table2" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="0"  cellspacing="0" width="100%">
		<tr align="center">
			<td  class="Estilo6" >
            	<!--crear un input solo para la cadena del where!-->
				<input type="hidden" name="fcwhere" id="fcwhere" value="<%=cwhere%>">
				<input type="Radio" name="Radio1" id="Radio1" value="Rop1" " <%=Chk01%> >Guia<br>
				<input type="Radio" name="Radio1" id="Radio1" value="Rop2" <%=Chk02%> >N.Ingreso
                <input type="Radio" name="Radio1" id="Radio1" value="Rop3" <%=Chk03%> >N.Salida
			</td>
		</tr>
		<tr>
			<td class="Estilo5"><input type="text" id="POS" name="POS" value="<%=pos%>" /></td>
		</tr>
		<tr align="center">
			<td><%'FAC 20130320 if pos="" or marka=1 or sindatos=0 then %>
            	<input type="submit" id="ENV" name="ENV" value="MOSTRAR" title="<%=cad%>"/> 
				<%'end if %>
            </td>
		</tr>

		<tr align="center" >
			<td class="Estilo6" height="40"><hr /></td>
		</tr>
		<tr align="center">
			<td class="Estilo6"><%if pos<>"" then %>TOTAL PRENDAS:<%end if %></td>
		</tr>
		<tr align="center">
			<td><%if len(pos)>0 then %><input type="text" id="tota" name="tota" value="<%=tot%>"/><%end if %></td>
		</tr>
		<tr align="center" class="Estilo8">
			<td><%if marka=0 and sindatos=1 then %><input name="STK" type="button" class="Estilo8" id="STK" onClick="cargar()" value="CREAR TXT" />
				<%else %>
				<%if marka= 1 then %>Documento esta registrado<br /> en tienda - <%=tendero%><% end if %>
				<%end if %>
			</td>
		</tr>
        
		</table>
	</td>
	<td width="75%" rowspan ="8" valign="top">
		<%' Nro de columnas regresadas por el objeto RECORDSET
		if sindatos=0 then
			response.End()
		end if 
		columnas = rs.Fields.Count %>
		<%'*********************************************************************%>
		<table id="TABLA" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="1"  cellspacing="1"  border="1" width="100%">
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
			<td class="ESTILO0" align="center"><%=CINT(rs("c6_citem")) %></td>
			<td class="ESTILO0" align="center"><input type="text" value='<%=rs("c6_ccodigo")%>' class="ESTILO0" id="COD<%=CONT%>" name ="COD<%=CONT%>" readonly tabindex="-1" size="12"/></td>
			<td class="ESTILO0" width="65%">
            <input type="text" value='<%=rs("descri")%>' class="ESTILO0" id="DES<%=CONT%>" name ="DES<%=CONT%>" readonly tabindex="-1" style="width:100%"/></td>
			<td class="ESTILO0" align="right" width="2%"><input type="text" value='<%=rs("c6_ncanten")%>' class="ESTILO0" id="CAN<%=CONT%>" name ="CAN<%=CONT%>" maxlength="3" size="4"/></td>
			<td class="ESTILO0"><input type="text" value='<%=rs("ar_npreci6")%>' class="ESTILO0" id="PRE<%=CONT%>" name="PRE<%=CONT%>" maxlength="6" size="4" align="right" 
            onchange="this.value=toInt(this.value)"/></td>
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
rec = parseInt("<%=rs.recordcount%>",10)
if (rec > 0 )
	dd2('1');

function llena()
{ return true }

function cargar() {
//alert("marka ="+"<%'=marka%>");
// el recordcount me devuelve el total de lineas empezando en 1...
lin = parseInt("<%=rs.recordcount%>", 10);
//alert(lin);
if (lin > 0) {
	//alert("Procederemos a crearlos...");
	thisForm.action = ("comun/CREATXT.asp?lin=" + lin);
	thisForm.submit();
    //window.open("comun/inserguia.asp?lin=" + lin);
	}
else
	alert("No hay nada que ingresar");
}

function inipos() {
// LImpiamos la variable Pos de asp
<%'pos=request.Form("pos")%>
alert("[<%=pos%>]")
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
