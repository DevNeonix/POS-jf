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

<SCRIPT language="javascript" src="../includes/cal.js"></SCRIPT>
<script language="jscript" type="text/jscript">
addCalendar("Calendar1", "Elija una fecha", "inicio", "thisForm")
addCalendar("Calendar2", "Elija una fecha", "final", "thisForm")

function MUESTRA() {
    if (trim(thisForm.TT.value) == '')
       thisForm.TT.value == '<%=tienda%>'
    cad = 'cajadeta.asp?fec=' + trim(thisForm.inicio.value) + '&fec2=' + trim(thisForm.final.value) + '&tda=' + trim(thisForm.TT.value)
	/*alert(cad)*/
	parent.window.frames[1].window.location.replace(cad)
}

</script>

<body>
<form id ="thisForm" name= "thisForm" >
<table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="2"  cellspacing="4"  border="0">
    <tr valign="middle" >
      <% CAD =   " SELECT * FROM TIENDAS WHERE ESTADO ='A' order by descripcion "
                  '  response.write(cad)
                  '  response.write("<br>")
            RS.OPEN CAD,CNN
            IF rs.recordcount > 0 THEN rs.movefirst%>
       	        <td class="Estilo11" valign="middle" align="right" rowspan="2">
                    <label for="Radio">Tiendas:&nbsp;</label></td> 
                <td  class="Estilo12" align="left"  rowspan="2">
                    <select  name="TT" id="TT" >
                        <option value = "TT" selected>TODAS</option>
                        <%do while not rs.eof %>
                            <option value="<%=TRIM(RS("CODIGO"))%>"><%=TRIM(RS("DESCRIPCION")) %></option>
                            <%rs.movenext %>
                        <%loop %>
                    </select>
                </td>
    
        <%RS.CLOSE %>


		<td class="Estilo11"align = left  VALIGN=MIDDLE>Ingrese fecha a Procesar : </td> 
        <td class="Estilo11"align = left     VALIGN=MIDDLE>
			<A href="javascript:showCal('Calendar1')"><img height=16 src="../images/cal.gif" width=16 border=0></A>
        </td>
        <td>
			<INPUT ID=inicio NAME=inicio  READONLY VALUE ="<%=date()%>" tabindex="-1" width=70>
		</td>		
        <td class="Estilo11"align = left     VALIGN=MIDDLE>
			<A href="javascript:showCal('Calendar2')"><img height=16 src="../images/cal.gif" width=16 border=0></A>
        </td>
        <td>
			<INPUT ID=final NAME=final READONLY VALUE ="<%=date()%>" tabindex="-1" width=70>
		</td>		

        <td class="Estilo11"align = left VALIGN=MIDDLE onclick="MUESTRA()"><img src="../images/ok.gif" /></td> 
	</tr>
</table>
</form>
</body>

</html>
