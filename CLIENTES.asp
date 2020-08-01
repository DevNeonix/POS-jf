<%@ Language=VBScript %>
<%Session.LCID=2057%>
<% Response.Buffer = true %>
<%tienda = Request.Cookies("tienda")("pos") 	%>
<script type="text/jscript" language="jscript">
// SI AUTO ESTA EN cero, SIGNIFICA QUE ES CODIGO MANUAL
// SI auto ESTA EN 1, SIGNIFICA QUE LA LLAVE ES idENTITY
var auto=0;
var url = '../CLIENTES.asp'
var alias = 'CLIENTES'
var TBL = 'CLIENTES'
var PK  = 'CLIENTE'
var DS  = 'NOMBRE'
var largo = 11  // es el largo del PK (se usa en el dataentry)
var largo2 = 100  // es el largo del descriptor
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
alias = "CLIENTES"
'*********************************************
' Definir el NOMBRE de la columna del ORDER BY
Dim indice
indice = "CLIENTE"
'*********************************************
' Definir el NOMBRE de la PAGINA ASP de inicio
Dim urlBase
urlBase = "CLIENTES.asp"
'*********************************************
' Definir el nombre del Primary Key
Dim pk
pk = "CLIENTE"
'*********************************************
' Definir nombre de la columna descriptor
Dim ds
ds = "NOMBRE"
'*********************************************
' Definir el TITULO de la PAGINA ASP 
Dim TITULO
TITULO = "CLIENTES"
%>
<script type="text/jscript" language="jscript">
// **************************************************************
//  Indicar el nombre de la página donde se realizan los cambios 
// **************************************************************
var funcionalidad = 'comun/INSERCLI.asp?'

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
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" lang="es-pe" />
<title><%=titulo%></title>
</head>
<body>

<form name="thisForm" id="thisForm" method="post" action="clientes.asp">

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
CAD =	" SELECT top  "&pagesize&" CLIENTE,     " & _ 
        " NOMBRE, DIRECCION, TIPOCLI,MAIL,      " & _
        " FONO, DCTO1, DCTO2                    " & _
		" from CLIENTES                         " & _
        " WHERE CLIENTE >= '"&POS&"'            " & _
        " and estado ='A' and                   "
IF DES > "" THEN
	CAD = CAD + " nombre >= '"&DES&"' ORDER BY 2 "
ELSE
	CAD = CAD + " cliente >= '"&POS&"' ORDER BY 1 "
END IF
      
'	response.Write(cad)
%>
<!--#include file="includes/Cnn.inc"-->
<!--#include file="COMUN/FUNCIONESCOMUNES.ASP"-->
<!--#include file="COMUN/COMUNgeneral.ASP"-->
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

    Set RS2 = Server.CreateObject("ADODB.Recordset")
	RS2.ActiveConnection = Cnn
	
	RS2.CursorType       = 3 'CONST adOpenStatic = 3
	RS2.LockType         = 1 'CONST adReadOnly = 1
	RS2.CursorLocation   = 3 'CONST adUseClient = 3
%>

<%' Nro de columnas regresadas por el objeto RECORDSET	
columnas = rs.Fields.Count %>


<%'*********************************************************************%>
<table id="TABLA" align="center"  bordercolor="#FFFFFF" bgcolor="<%=Application("barra")%>" cellpadding="2" cellspacing="1" border="1" >
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
		<td><img src="images/primera.gif" style="cursor:pointer;" onClick="primera('<%=urlBase%>')" alt="PRIMERA PAGINA" /></td>
		<td><img src="IMAGES/PREV.GIF" alt="PAGINA ANTERIOR" onClick="atras(alias, '<%=indice%>')" style="CURSOR:pointer" /></td>
		<td><img src= "images/arriba.gif" alt="REGISTRO ANTERIOR" onClick="retrocede()" style="CURSOR:pointer" /></td>
		<td><img src="images/abajo.gif" alt="REGISTRO SIGUIENTE" onClick="avanza()"  style="CURSOR:pointer" /></td>
		<td><img src="images/next.gif" alt="PAGINA SIGUIENTE" onClick="pagina('<%=urlBase%>'+'?pos=')" style="CURSOR:pointer" /></td>
		<%  ' PARA LA FUNCION ULTIMA : 
			' enviar el nombre de la página de retorno
			' el nombre de la tabla 
			' el nombre de la columna de primary key%>
		<td><img src= "images/ultima.gif" alt="ULTIMA PAGINA" onClick="ultima()" style="cursor:pointer;" /></td>
	<td><img src="images/PRINT.gif" alt="IMPRESION" onClick="imprime()"	style="cursor:pointer;" /></td>
    <td><img src="images/SEARCH.gif" onClick="document.all.seeker.style.display='block'" alt="BUSCAR" style="cursor:pointer;" /></td>
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
<table	align="center" border="0" id="DATAENTRY"  >
   <tr>
   		<td>
            <table	align="center" cellpadding="1" cellspacing="0"  bgcolor="WHITE" border="1" >
              <tr valign="middle"> 
                <td bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  CODIGO : &nbsp;&nbsp;  </td>
                <td bgcolor="WHITE" valign="middle"> 
                  <input type="text" id="COD" name="COD"  class="Estilo2" tabindex="-1" readonly />
                </td>
                <td bgcolor="<%=(Application("borde"))%>" align="right" class="Estilo3">NOMBRE :&nbsp;&nbsp;</td>
                <td bgcolor="#FFFFFF" colspan="4"> 
                  <input type="text" id="DES" name="DES" class="Estilo2" style="width:100%"  />
                </td>
              </tr>
              <tr valign="middle"> 
                <td bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  DIRECCION :&nbsp;&nbsp;  </td>
                <td bgcolor="WHITE"valign="middle" colspan="4"> 
                  <input type="text" id="DIR" name="DIR" class="Estilo2" maxlength="200" />
                </td>
            </tr>
            <tr><td width="10%" bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  CORREO :  &nbsp;&nbsp; </td>
                <td bgcolor="WHITE" valign="middle" colspan="4"> 
                  <input type="text" id="COR" name="COR"  class="Estilo2" maxlength="50"  />
                </td>
            </tr>
            <tr>
                <td bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  TELEFONO : &nbsp;&nbsp;  </td>
                <td bgcolor="WHITE" valign="middle"> 
                  <input type="text" id="TEL" name="TEL"  class="Estilo2" maxlength="12" />
                </td>
                <td  bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  TIPO CLIENTE :&nbsp;&nbsp;   </td>
                <td bgcolor="WHITE"  valign="middle" colspan="2"> 
                  <select type="text" id="TIP" name="TIP" class="Estilo2">
                       <option></option> 
                       <%CAD =	" select tg_cclave as CODIGO,   " & _
                                " TG_CDESCRI AS DESCRIPCION     " & _
                                " from RSFACCAR..AL0012TABL     " & _
                                " where tg_ccod = '67'          " & _
                                " order by 2                    " 
                       
                       rs2.open CAD ,cnn
                       rs2.movefirst
                       do while not rs2.eof%>
                       <option value="<%=rs2("codigo") %>"><%=rs2("descripcion") %></option>
                        <%rs2.movenext
                       loop%>
                  </select>
                  <%rs2.close %>                  
                </td>
              </tr>
              <tr valign="middle" style="display:none"> 
                <td bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  DCTO. 1 :&nbsp;&nbsp;   </td>
                <td bgcolor="WHITE" valign="middle"> 
                  <input type="text" id="DT1" name="DT1"  maxlength="5" class="Estilo2" onKeyUp="this.value=toInt(this.value)"  value="0" />
                </td>
                <td bgcolor="<%=(Application("borde"))%>" align="right" class="Estilo3">DCTO. 2 :&nbsp;&nbsp;</td>
                <td bgcolor="#FFFFFF" colspan="4"> 
                  <input type="text" id="DT2" name="DT2"  maxlength="5"  value="0" class="Estilo2" onKeyUp="this.value=toInt(this.value)" style="width:100%" />
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

function llena(pos) {
    var t = document.all.TABLA;
    thisForm.DIR.value = ltrim(t.rows(pos).cells(2).innerText);
    thisForm.TEL.value = ltrim(t.rows(pos).cells(5).innerText);
    thisForm.COR.value = ltrim(t.rows(pos).cells(4).innerText);
    var subcadena = trim(t.rows(pos).cells(3).innerText);
    var elemento = thisForm.TIP;
    thisForm.TIP.selectedIndex = seleindice(subcadena, elemento);
    thisForm.DT1.value = ltrim(t.rows(pos).cells(6).innerText);
    thisForm.DT2.value = ltrim(t.rows(pos).cells(7).innerText);
return true
}

function NUEVO_onclick() {
    chk = "0"
    if (auto == 1) {
        thisForm.COD.readOnly = false;
        thisForm.COD.value = 'AUTO';
        thisForm.COD.readOnly = true;
        thisForm.DES.focus();
    }
    else {
        thisForm.COD.readOnly = false;
        thisForm.COD.value = '';
        thisForm.COD.focus();
    }
    thisForm.DES.value = '';
    thisForm.DIR.value = '';
    thisForm.TEL.value = '';
    thisForm.COR.value = '';
    thisForm.DT1.value = 0;
    thisForm.DT2.value = 0;
    thisForm.TIP.selectedIndex = -1;
}
function VALIDA() {
    if (window.thisForm.COD.value == "" && largo > 0 && window.thisForm.COD.value != 'AUTO') {
        alert('Debe Informar el Código');
        window.thisForm.COD.focus();
        return false;
    }
    else {
        cad = window.thisForm.COD.value;
        ff = trim(window.thisForm.COD.value);
        if (ff.length > largo && largo > 0 && window.thisForm.COD.value != 'AUTO') {
            alert("El código no puede tener más de " + largo + " dígitos")
            return false;
        }
    }
    if (trim(window.thisForm.DES.value) == "" && largo2 > 0) {
        alert('Debe Informar EL NOMBRE DEL CLIENTE');
        window.thisForm.DES.focus();
        return false;
    }
    else {
        thisForm.DES.value = Left(ltrim(rtrim(thisForm.DES.value)), largo2)
    }
    if (isEmail(trim(window.document.all.COR.value), 1) == false)
        return false;

    /*if (trim(document.all.COR.value) == '') {
        alert("Debe informar el coreo por Orden SUNAT")
        return false;
    }*/


/* FAC 20121207 Sra. Saba indica no es inpedimento para crear cliente    if (isInt(trim(window.thisForm.TEL.value)) == false) {
        window.thisForm.TEL.focus();
        return false;
    }
*/
    if (trim(window.thisForm.TIP.value) == "" ) {
        alert('Debe Informar EL TIPO DE CLIENTE');
        window.thisForm.TIP.focus();
        return false;
    }

    return true;
}
function completa() {
    var cad
    cad = 'DIR=' + toAlpha(ltrim(thisForm.DIR.value));
    cad += '&TEL=' + trim(thisForm.TEL.value);
    cad += '&cor=' + trim(thisForm.COR.value);
    cad += '&tip=' + trim(thisForm.TIP.value);
    if (trim(thisForm.DT1.value) == '')
        thisForm.DT1.value=0
    cad += '&DT1=' + trim(thisForm.DT1.value);
    if (trim(thisForm.DT2.value) == '')
        thisForm.DT2.value = 0
    cad += '&DT2=' + trim(thisForm.DT2.value);
    //alert(cad)
    return cad
}
</script>

<%	RS.Close 
	SET RS  = NOTHING
	Cnn.Close
	SET Cnn = NOTHING %>

</form>
</body>
</html>
