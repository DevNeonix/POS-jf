<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%	tienda = Request.Cookies("tienda")("pos") 	%>
<script type="text/jscript" language="jscript">
// SI AUTO ESTA EN cero, SIGNIFICA QUE ES CODIGO MANUAL
// SI auto ESTA EN 1, SIGNIFICA QUE LA LLAVE ES idENTITY
var auto=0;
var url = '../TIPMOV.asp?'
var alias = 'TIPMOV'
var TBL = 'TIPMOV'
var PK  = 'CODIGO'
var DS  = 'DESCRIPCION'
var largo = 2  // es el largo del PK (se usa en el dataentry)
var largo2 = 50  // es el largo del descriptor
var oldrow =1
var olddata =''
var chk = ''
</script>
<%' Definir el tamaño de la pagina
Dim pagesize 
pagesize =20
'****************************************
' Definir el NOMBRE de la Tabla base
Dim ALIAS
alias = "TIPMOV"
'*********************************************
' Definir el NOMBRE de la columna del ORDER BY
Dim indice
indice = "CODIGO"
'*********************************************
' Definir el NOMBRE de la PAGINA ASP de inicio
Dim urlBase
urlBase = "TIPMOV.asp"
'*********************************************
' Definir el nombre del Primary Key
Dim pk
pk = "CODIGO"
'*********************************************
' Definir nombre de la columna descriptor
Dim ds
ds = "DESCRIPCION"
'*********************************************
' Definir el TITULO de la PAGINA ASP 
Dim TITULO
TITULO = "TIPO DE MOVIMIENTOS"
%>
<script type="text/jscript" language="jscript">
// **************************************************************
//  Indicar el nombre de la página donde se realizan los cambios 
// **************************************************************
var funcionalidad = 'comun/INSERTBL.asp?'

function dd2(ff)
{	// LLENA TEXTBOX ADICIONALES AL COMUN
	// LOS DEL COMUN SON CODIGO Y DESCRIPCION
	var t = document.all.TABLA;
	var pos = parseInt(ff) ;
	dd(ff);
}
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=titulo%></title>
</head>
<body>

<form name="thisForm" id="thisForm" method="post" action="TIPMOV.asp">

<table width="100%">
	<tr><td colspan="3"><hr /></td></tr>
    <tr><td align="center" class="Estilo6"><%=titulo%></td></tr>
	
</table>
<%
POS = Request.QueryString("pos")
if pos = "" or isnull(pos)  or pos = " " then
	pos = ""
end if
des = Request.QueryString("des")
if des = "" or isnull(des)  or des = " " then
	des = "" 
end if
'****************************************************
' Texto del Comando (SELECT) a ejecutar (POR DEFAULT)
'****************************************************
CAD =	" SELECT top  "&pagesize&" *  " & _ 
		" from "&alias&"  "& _
        " ORDER BY "& indice &"  " 
		'response.Write(cad)
%>

<!--#include file="includes/Cnn.inc"-->
<!--#include file="COMUN/FUNCIONESCOMUNES.ASP"-->
<!--#include file="COMUN/COMUNtbl.ASP"-->
<%	RS.Open CAD, Cnn
    SINDATOS =1
	CONT = 1
	IF  RS.RECORDCOUNT > 0 THEN 	
		RS.MOVEFIRST
	ELSE
		RESPONSE.Write("TABLA SIN DATOS")	
		SINDATOS =0
		'RESPONSE.End()
	END IF	
%>

<%' Nro de columnas regresadas por el objeto RECORDSET	
columnas = rs.Fields.Count %>


<%'*********************************************************************%>
<table id="TABLA" align="center"  width="100%" bordercolor="#FFFFFF"
	  bgcolor="<%=Application("barra")%>"  cellpadding="2"  cellspacing="1"  border="1" >
<%'**************************%>
<%'LINEA DE CABECERA STANDAR %>
<%'**************************%>
<tr>
<%LIMITE =  rs.fields.count - 1%>
<%for I=0 to LIMITE %>
	<td align="center" class="Estilo8">
		<%=RS.FIELDS(I).NAME%>
	</td>
<%next%>	
</tr>
<%'*****************************%>
<%' MUESTRA EL GRid (2 colorES) %>
<%'*****************************%>
<%IF NOT RS.EOF THEN%>
<%DO WHILE NOT RS.EOF%>

    <tr  bgcolor="<% if CONT mod 2  = 0 THEN 
                response.write(Application("color1"))
                else
	            response.write(Application("color2"))
	            end IF%>"
	            onclick="dd('<%=(cont)%>')" id="fila<%=Trim(Cstr(cont))%>" >
      <%for I=0 to LIMITE %>
		<td>
			<font face="Arial, Helvetica, sans-serif" color="MidnightBlue" size="1">
			<b><%=RS.FIELDS.ITEM(I)%></b>&nbsp;
			</font>
		</td>
	  <%NEXT%>	

	<%RS.MOVENEXT%>
	<%CONT = CONT + 1%>
  </tr>
    
	<%loop%>
</table>
<table border="0" align="center"  cellspacing="5">
	<tr valign="top">
		<td><img src="images/primera.gif" style="cursor:pointer;" onclick="primera()" alt="PRIMERA PAGINA" /></td>
		<td><img src="IMAGES/PREV.GIF" alt="PAGINA ANTERIOR" onclick="atras()" style="CURSOR:pointer" /></td>
		<td><img src= "images/arriba.gif" alt="REGISTRO ANTERIOR" onclick="retrocede()" style="CURSOR:pointer" /></td>
		<td><img src="images/abajo.gif" alt="REGISTRO SIGUIENTE" onclick="avanza()"  style="CURSOR:pointer" /></td>
		<td><img src="images/next.gif" alt="PAGINA SIGUIENTE" onclick="pagina()" style="CURSOR:pointer" /> </td>
		<td><img src= "images/ultima.gif" alt="ULTIMA PAGINA" onclick="ultima()" "cursor:pointer;" /></td>
	<td><img src="images/PRINT.gif" alt="IMPRESION" onclick="imprime()" style="cursor:pointer;" /></td>
    <td><img src="images/SEARCH.gif" onclick="document.all.seeker.style.display='block'" alt="BUSCAR" style="cursor:pointer;" /></td>
	<td id="seeker" name="seeker" style="display:none">
	<table align="center"  width="100%" bordercolor="#FFFFFF"
	  bgcolor="lightgrey"  cellpadding="0"  cellspacing="1"  border="1" >
	  <tr>  
	    <td  bgcolor='<%=Application("COLOR2")%>'><font face="arial" color='<%=Application("Titulo")%>' size="1">
		    <b><%=pk%></b></font></td>
		<td><input id="kod" name="kod" value="" /> </td>    
	  </tr>
	  <tr>  
	    <td  bgcolor='<%=Application("COLOR2")%>'><font face="arial" color='<%=Application("Titulo")%>' size="1">
		    <b><%=DS%></b></font></td>
		<td><input id="ds" name="ds" value="" /> </td>    
	  </tr>
	  <tr>  
	    <td  bgcolor='<%=Application("COLOR2")%>' align="center" style="cursor:pointer" onClick="document.all.seeker.style.display='none'">
	        <font face="arial" color="red" size="1">
		    <b><u>(X) Cerrar</u></b></font></td>
		<td  bgcolor='<%=Application("COLOR2")%>' align="CENTER" style="cursor:pointer" onClick="BUSCA('<%=urlBase%>','<%=alias%>')">
		<font face="arial" color='<%=Application("Titulo")%>' size="1">
		    <b><U>FILTRAR</U></b></font></td>
	  </tr>
	 </table> 
	</td>	
	</tr>
	</table>
<%END IF %>
<iframe frameborder="1" style="visibility:hidden" height="1" width="10" id="ACTIV" name="ACTIV"></iframe>
<table	width="100%" border="0" id="DATAENTRY"  >
   <tr>
   		<td>
            <table	align="center" width="100%" 
            cellpadding="1" cellspacing="0"  bgcolor="WHITE" border="1" >
              <tr valign="middle"> 
                <td width="10%" bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo8">
                  CODIGO :   </td>
                <td bgcolor="WHITE" width="10%" valign="middle"> 
                  <input type="text" id="COD" name="COD" 
                    class="DATOSGRANDE" 
                    tabindex="-1" readonly="readonly" />
                </td>
                <td width="10%" bgcolor="<%=(Application("borde"))%>" align="right" class="Estilo8">DESCRIPCION :</td>
                <td bgcolor="#FFFFFF" colspan="5"> 
                  <input type="text" id="DES" name="DES"  style="width:100%"  />
                </td>
              </tr>
            </table>
         </td>   
        </tr>
        <tr>
        <td>
<table border="0" align="center"  cellspacing="3">
    <tr>
        <td>
            <img src="images/NEW.gif"  alt="REGISTRO EN BLANCO"
                onclick="NUEVO_onclick()" style="cursor:pointer;" />	
        </td>
        <td>		
            <img src="images/DISK.gif" alt="GRABAR"
                onclick="GRABAR_onclick()" 
                style="cursor:pointer;" />	
        </td>		
        <td>		
            <img src="images/DELETE.gif" alt="ELIMIAR REGISTRO"
                onclick="elimina()" style="cursor:pointer;" />	
        </td>	
        	
    </tr>
</table>
</td>
</tr>
</table>
<script type="text/jscript" language="jscript">
rec = parseInt('<%=rs.recordcount%>',10)
if (rec > 0 )
dd2('1');

thisForm.COD.maxLength=largo
thisForm.DES.maxLength=largo2
if (rec <= 0)
{  NUEVO_onclick()
	SS = trim('<%=SINDATOS%>')
	if (SS == "1")
	{thisForm.kod.maxLength=largo
	thisForm.ds.maxLength=largo2
	}
}    
</script>

<%	RS.Close 
	SET RS  = NOTHING
	Cnn.Close
	SET Cnn = NOTHING %>

</form>
</body>
</html>
