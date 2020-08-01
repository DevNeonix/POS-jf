<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.Buffer = false %>
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
<SCRIPT language="javascript" src="../includes/jquery.js"></SCRIPT>
<SCRIPT language="javascript" src="../includes/cal.js"></SCRIPT>
<script language="jscript" type="text/jscript">
addCalendar("Calendar1", "Elija una fecha", "inicio", "thisForm")
addCalendar("Calendar2", "Elija una fecha", "final", "thisForm")

</script>
<body>
<form id ="thisForm" name= "thisForm" >
<table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="2"  cellspacing="4"  border="0">
    <tr valign="middle" >
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
		<td>
			<input type="checkbox" id="agrupar">Ocutar fecha y Agrupar
		</td>		
        <td class="Estilo11"align = left VALIGN=MIDDLE onclick="MUESTRA()"><img src="../images/ok.gif" /></td> 
        <td class="Estilo11"align = left VALIGN=MIDDLE onclick="excel()"><img style="width: 28px" src="../images/xl1.png" /></td> 
	</tr>
</table>
</form>
<div id="contenedor" style="width:100%;border:none;"></div>
<script type="text/javascript">
	var iInicio;
	var iFinal;
	function MUESTRA(){
		$("#contenedor").html("cargando...")
		iInicio = document.getElementById("inicio").value;
		iFinal = document.getElementById("final").value;
		
		
		var agru = "false";
		if(document.getElementById("agrupar").checked){
			agru="true"
		}
		var cad = './reporteripleydeta.asp?ini='+iInicio+'&agru='+agru+'&fin='+iFinal;
		$.ajax({url:cad+'&excel=0',type:'get',cache:false,success:function(res){
			$("#contenedor").html(res)
		}});
	}
	function excel(){
		$("#contenedor").html("")
		iInicio = document.getElementById("inicio").value;
		iFinal = document.getElementById("final").value;
	
		var agru = "false";
		if(document.getElementById("agrupar").checked){
			agru="true"
		}
		var cad = './reporteripleydeta.asp?ini='+iInicio+'&fin='+iFinal+'&excel=1'+'&agru='+agru
		window.open(cad)
	}
</script>
</body>

</html>
