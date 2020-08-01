<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos")%>

<script type="text/jscript" language="jscript">
// SI AUTO ESTA EN cero, SIGNIFICA QUE ES CODIGO MANUAL
// SI auto ESTA EN 1, SIGNIFICA QUE LA LLAVE ES idENTITY
var auto=0;
var url = '../articulos.asp?'
var alias = 'VIEW_ARTICULOS_TIENDA'
var TBL = 'VIEW_ARTICULOS_TIENDA'
var PK  = 'CODIGO'
var DS  = 'DESCRI'
var largo = 10  // es el largo del PK (se usa en el dataentry)
var largo2 = 100  // es el largo del descriptor
var oldrow =1
var olddata =''
var chk = ''
var nuevo = 0
</script>

<%' Definir el tamaño de la pagina
Dim pagesize 
pagesize =10
'****************************************
' Definir el NOMBRE de la Tabla base
Dim ALIAS
alias = "VIEW_ARTICULOS_TIENDA"
'*********************************************
' Definir el NOMBRE de la columna del ORDER BY
Dim indice
indice = "CODIGO"
'*********************************************
' Definir el NOMBRE de la PAGINA ASP de inicio
Dim urlBase
urlBase = "ARTICULOS.asp"
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
TITULO = "MAESTRO DE ARTICULOS"
%>

<script type="text/jscript" language="jscript">
// **************************************************************
//  Indicar el nombre de la página donde se realizan los cambios 
// **************************************************************
var funcionalidad = 'comun/INSERart.asp?'

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
<body onkeypress="veri()">

<form name="thisForm" id="thisForm" method="post" action="articulos.asp">

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
CAD =	" SELECT top  "&pagesize&"              " & _
        " CODIGO, DESCRI, AR_CUNIDAD AS UNI,    " & _
        " AR_CMONVTA AS MON, STOCK, AR_NIGVPOR  " & _
        " AS IGV, LISTA1, LISTA2, LISTA3,       " & _
        " MINIMO, PLANILLA AS PLA              " & _
		" from VIEW_ARTICULOS_TIENDA            " & _
        " WHERE ESTADO = 'A' AND                " & _
        " tienda = '"&tienda&"'  AND            " 

IF DES > "" THEN
	CAD = CAD + " DESCRI >= '"&DES&"' ORDER BY 2 "
ELSE
	CAD = CAD + " codigo >= '"&POS&"' ORDER BY 1 "
END IF
'	response.Write(cad)
%>

<!--#include file="includes/Cnn.inc"-->
<!--#include file="COMUN/FUNCIONESCOMUNES.ASP"-->
<!--#include file="COMUN/COMUNgeneralJFF.ASP"-->

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

<table	align="center" border="0" id="DATAENTRY"   width="100%">
<tr>
	<td>
		<table	align="center" cellpadding="1" cellspacing="0"  bgcolor="WHITE" border="1">
		<tr valign="middle"> 
    		<td bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
            	CODIGO :  &nbsp;</td>
            <td bgcolor="WHITE" valign="middle" width="90px"> 
            	<input type="text" id="COD" name="COD"  class="Estilo2" tabindex="-1" readonly tabindex="-1" onChange="CAMBIA(this)" maxlength="10" size="80px"/>
            </td>
            <td bgcolor="<%=(Application("borde"))%>" align="right" class="Estilo3">
                Descripción : &nbsp;</td>
            <td bgcolor="#FFFFFF" colspan="6" width="300px"> 
                <input type="text" id="DES" name="DES" class="Estilo2" style="width:100%" tabindex="-1" readonly tabindex="-1" /></td>
        </tr>
        <tr valign="middle"> 
            <td  bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
            	Moneda : &nbsp;  </td>
            <td bgcolor="WHITE" valign="middle" > 
                <select type="text" id="MON" name="MON" class="Estilo2" tabindex="-1">
                <option></option> 
                <%CAD =	" select tg_cclave as CODIGO,   " & _
                        " TG_CDESCRI AS DESCRIPCION     " & _
                        " from RSFACCAR..AL0012TABL     " & _
                        " where tg_ccod = '03'          " & _
                        " order by 2                    " 
                       
                rs2.open CAD ,cnn
                rs2.movefirst
                do while not rs2.eof%>
                   <option value="<%=rs2("codigo") %>"><%=rs2("descripcion") %></option>
                   <%rs2.movenext
                loop%>
                </select>
                <%RS2.CLOSE %>
            </td>
            <td bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
            	P.V.P : &nbsp;   </td>
            <td bgcolor="WHITE"valign="middle" > 
            	<input type="text" id="L1" name="L1" class="Estilo2" /></td>
            <td bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                Lista 2 : &nbsp;  </td>
                <td bgcolor="WHITE" valign="middle" > 
                  <input type="text" id="L2" name="L2"  class="Estilo2" maxlength="50"  />
                </td>
            
                <td bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  Lista 3 :  &nbsp;  </td>
                <td bgcolor="WHITE" valign="middle"> 
                  <input type="text" id="L3" name="L3"  class="Estilo2" maxlength="12" />
                </td>
            </tr>
            <tr valign="middle"> 
               
                 <td bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  Unidad : &nbsp;   </td>
                <td bgcolor="WHITE"valign="middle" > 
                   <select type="text" id="UNI" name="UNI" class="Estilo2" tabindex="-1">
                       <option></option> 
                       <%CAD =	" select tg_cclave as CODIGO,   " & _
                                " TG_CDESCRI AS DESCRIPCION     " & _
                                " from RSFACCAR..AL0012TABL     " & _
                                " where tg_ccod = '05'          " & _
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
                <td bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  Stock :  &nbsp;  </td>
                <td bgcolor="WHITE" valign="middle"> 
                  <input type="text" id="STK" name="STK"  class="Estilo2"  onfocus="seleccionar(this)" maxlength="12" <%if tienda <> "06" then %>readonly tabindex="-1" disabled <%end if%>/>
                </td>
                <td bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  Minimo :  &nbsp;  </td>
                <td bgcolor="WHITE" valign="middle"> 
                  <input type="text" id="MNN" name="MNN" onkeyup="this.value=toInt(this.value)"  onfocus="seleccionar(this)" class="Estilo2" maxlength="1" />
                </td>
                <td bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  PLANTILLA ?  &nbsp;  </td>
                <td bgcolor="WHITE" valign="middle"> 
                  <select  id="PLA" name="PLA"  class="Estilo2" >
                    <option></option>
                    <option value="1">SI</option>
                    <option value="0">NO</option>
                  </select>
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
function completa() {
    l1 = thisForm.L1.value
    l2 = thisForm.L2.value 
    l3 = thisForm.L3.value
    mnn = thisForm.MNN.value
    cad = 'l1=' + l1
    cad += '&l2=' + l2
    cad += '&l3=' + l3
    cad += '&mnn=' + mnn
    cad += '&PLA=' + thisForm.PLA.value
    cad += '&STK=' + thisForm.STK.value
    return cad
     }
function llena(pos) {
    var t = document.all.TABLA;
    thisForm.DES.value = ltrim(t.rows(pos).cells(1).innerText);
    thisForm.L1.value = ltrim(t.rows(pos).cells(6).innerText);
    thisForm.L2.value = ltrim(t.rows(pos).cells(7).innerText);
    thisForm.L3.value = ltrim(t.rows(pos).cells(8).innerText);
    thisForm.STK.value = ltrim(t.rows(pos).cells(4).innerText);
    thisForm.MNN.value = ltrim(t.rows(pos).cells(9).innerText);
    var subcadena = trim(t.rows(pos).cells(2).innerText);
    var elemento = thisForm.UNI;
    thisForm.UNI.selectedIndex = seleindice(subcadena, elemento);
    var subcadena = trim(t.rows(pos).cells(3).innerText);
    var elemento = thisForm.MON;
    thisForm.MON.selectedIndex = seleindice(subcadena, elemento);
    var subcadena = trim(t.rows(pos).cells(10).innerText);
    var elemento = thisForm.PLA;
    thisForm.PLA.selectedIndex = seleindice(subcadena, elemento); 
    nuevo = 0
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
    thisForm.L1.value = '';
    thisForm.L2.value = '';
    thisForm.L3.value = '';
    thisForm.STK.value = 0;
    thisForm.MNN.value = '';
    thisForm.UNI.selectedIndex = -1;
    thisForm.MON.selectedIndex = -1;
    thisForm.PLA.selectedIndex = -1;
    nuevo = 1
}

function veri() {
    if (event.keyCode == 13) {
        CAMBIA(document.all.COD)
    }
}
function CAMBIA(objj) {
    dato = objj.value
 /*   document.all.ACTIV.style.visibility = 'visible'
    document.all.ACTIV.height = "150"
    document.all.ACTIV.width = "100%"
    alert(dato)
   */
    if (nuevo == 1) {
        cad = 'bake/bakeprendas.asp?pos=' + dato
        document.all.ACTIV.src = cad;
    }



return true
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
    if (window.thisForm.PLA.selectedIndex == -1) {
        alert("Por favor indicar si va a formar parte de los articulos que\ngeneran reposicion automatica")
        window.thisForm.PLA.focus()
        return false
    }
    return true;
}

</script>

<%	RS.Close 
	SET RS  = NOTHING
	Cnn.Close
	SET Cnn = NOTHING %>

</form>
</body>
</html>
