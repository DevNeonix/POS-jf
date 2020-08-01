<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<% tienda = Request.Cookies("tienda")("pos") %>

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


function MUESTRA() {
    if (trim(thisForm.TT.value) == '')
       thisForm.TT.value == '<%=tienda%>'
    cad = 'SKU_RIPLEY.asp?INI=' + trim(thisForm.ini.value) + '&FIN=' + trim(thisForm.fin.value) 
	/*alert(cad)*/
	parent.window.frames[1].window.location.replace(cad)
}

</script>
<%aMes = array("" ,"ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SET","OCT","NOV","DIC") %>
<body>
<form id ="thisForm" name= "thisForm" >
<table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="2"  cellspacing="4"  border="0">
    <tr valign="middle" >

		<td class="Estilo11"align = left  VALIGN=MIDDLE>Acumular desde :A&Ntilde;0/mes </td> 
        <td><%cad = "select distinct anio+mes  as ini from ripley order by 1"
        rs.open cad,cnn
        IF RS.RECORDCOUNT > 0 THEN RS.MOVEFIRST%>
		<select id="ini" name="ini" class="inputs">    
			<option value=''></option>
		<%DO WHILE NOT RS.EOF%>
			<option value='<%=TRIM(rs("ini"))%>'><%=aMes(cint(right(trim(rs("ini")),2)))%> - <%=left(trim(rs("ini")),4)%></option>
			<%RS.MOVENEXT%>
		<%LOOP%>
		</select>
        </td>
        <td class="Estilo11"align = left  VALIGN=MIDDLE>hasta :  </td> 
        <td>
        <select id="fin" name="fin" class="inputs">    
			<option value=''></option>
            <%rs.movefirst%>
		<%DO WHILE NOT RS.EOF%>
            <option value='<%=TRIM(rs("ini"))%>'><%=aMes(cint(right(trim(rs("ini")),2)))%> - <%=left(trim(rs("ini")),4)%></option>
			<%RS.MOVENEXT%>
		<%LOOP%>
		</select></td>

        <td class="Estilo11"align = left VALIGN=MIDDLE onclick="MUESTRA()"><img src="../images/ok.gif" /></td> 
	</tr>
</table>
<iframe width="100%" onLoad="calcHeight();" src="" id="body0" name="body0"  scrolling="yes" frameborder="3" height="450" align="middle">
</iframe>

</form>

</body>
<script language="jscript" type="text/jscript">



    function calcHeight() {
        document.getElementById('body0').height = 425 

    }
  function MUESTRA() {
    if (trim(thisForm.ini.value) == '') {
        alert("Favor indicar el período inicial")
        return false;
    }
    if (trim(thisForm.fin.value) == '') {
        alert("Favor indicar el período Final")
        return false;
    }
    
    ii = parseInt(trim(thisForm.ini.value), 10)
    ff = parseInt(trim(thisForm.fin.value), 10)
    if ((ff - ii) < 0) {
        alert("el período inicial no puede ser mayor que el final")
        return false
    }
    cad = 'sku_ripley.asp?ini=' + trim(thisForm.ini.value) + '&fin=' + trim(thisForm.fin.value)
document.all.body0.src = cad
    return true

}

</script> 
</html>
