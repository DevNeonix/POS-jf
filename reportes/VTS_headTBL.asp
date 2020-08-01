<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = true%>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos")%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<link REL="stylesheet" TYPE="text/css" HREF="..\ventas.CSS" >
<!--#include file="../comun/funcionescomunes.asp"-->
<!--#include file="../includes/funcionesVBscript.asp"-->
<!--#include file="../includes/cnn.inc"-->

<SCRIPT language="javascript" src="../includes/cal.js"></SCRIPT>

<script type="text/jscript" language="jscript">
	function calcHeight() {
		//find the height of the internal page
		var the_height = document.getElementById('mirada').contentWindow.document.body.scrollHeight;
		//change the height of the iframe
		document.getElementById('mirada').height = the_height + 20;
	}
</script>

<script type="text/jscript" language="jscript">
    function strSQL() {
    
    cad =  "VTS_detaTBL.asp?TDA=" + document.all.tienda.value + '&ini=' + document.all.ini.value + '&fin=' + document.all.fin.value
    cad += "&tem=" + document.all.tempo.value
    //alert(cad)
    parent.window.frames[1].window.location.replace(cad)
    return true
}
</script>

<script language="jscript" type="text/jscript">
    addCalendar("Calendar1", "Elija una fecha", "ini", "thisForm")
    addCalendar("Calendar2", "Elija una fecha", "fin", "thisForm")
</script>

<% CAD =   " SELECT * FROM TIENDAS WHERE ESTADO ='A' order by descripcion "
          '  response.write(cad)
          '  response.write("<br>")
    RS.OPEN CAD,CNN
    IF rs.recordcount > 0 THEN rs.movefirst
%>

<body onload="top.parent.window.document.getElementById('body0').height = 480">

<form id ="thisForm" name= "thisForm" >

<table width="100%">
    <tr><td align="center" class="Estilo6">Ventas  Articulo :</td></tr>
</table>

<table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="2"  cellspacing="1"  border="0">
	<tr valign="middle" >
    	<td class="Estilo11" valign="middle" align="right">Tiendas:&nbsp;</td> 
        <td  class="Estilo12" align="left">
            <select  name="tienda" id="tienda">
                <option value = "TT" selected>TODAS</option>
                <%do while not rs.eof %>
                    <option value="<%=TRIM(RS("CODIGO"))%>"><%=TRIM(RS("DESCRIPCION")) %></option>
                    <%rs.movenext %>
                <%loop %>
                <%rs.close %>
            </select>
        </td>
        <td width="15px;"></td>
        <td width="15px">&nbsp;</td>
        <td class="Estilo11" valign="middle" align="right">Temporada:&nbsp;</td> 
        <td  class="Estilo12" align="left">
            <select  name="tempo" id="tempo">
                <option value = "" selected>TODAS</option>
                <% CAD =   " SELECT * FROM temporadas WHERE ESTADO ='A' order by descripcion "
          '  response.write(cad)
          '  response.write("<br>")
    RS.OPEN CAD,CNN
    IF rs.recordcount > 0 THEN rs.movefirst
                
                do while not rs.eof %>
                    <option value=" <%=TRIM(RS("CODIGO"))%>"><%=TRIM(RS("DESCRIPCION")) %></option>
                    <%rs.movenext %>
                <%loop %>
                <%rs.close %>
            </select>
        </td>
        <td width="15px;"></td>
        <td width="15px">&nbsp;</td>
        <td class="Estilo11" align = left  VALIGN=MIDDLE>Inicio : </td> 
        <td class="Estilo11" align = left  VALIGN=MIDDLE>
        	<A href="javascript:showCal('Calendar1')"><img height=16 src="../images/cal.gif" width=16 border=0></A>
        </td>
        <td>
			<INPUT ID="ini" NAME="ini" VALUE ="<%=date()%>" tabindex="-1" readonly class="Estilo21" style="width:70px">
		</td>
        
		<td class="Estilo11" align = left  VALIGN=MIDDLE>Fin : </td> 
        <td class="Estilo11" align = left  VALIGN=MIDDLE>
        	<A href="javascript:showCal('Calendar2')"><img height=16 src="../images/cal.gif" width=16 border=0></A>
        </td>
        <td>
			<INPUT ID="fin" NAME="fin" VALUE ="<%=trim(date())%>" tabindex="-1" readonly class="Estilo21" style="width:70px">
		</td>		
        
        <td><img src="../images/ok.gif" onclick="strSQL()" style="cursor:pointer;"/></td>            
    </tr>
   
</table>

<iframe src="" id="mirada" name="mirada" style="display:none" width="100%"></iframe> 
</form>
</body>
<%set rs=nothing %>
</html>
