<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = true %>
<%Session.LCID=2058%>
<%tienda = Request.Cookies("tienda")("pos") %>
<%'**************************************************************************
' CAMBIA DATOS DE CABECERAS COMO CLIENTE, NUMERO DE DOCUMENTO Y FECHA 
'**************************************************************************** %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<style>
.botonimagen{
  background-image:url('images/check.png');
  background-repeat: no-repeat;
  background-position: left center;
  cursor:pointer;
  border:none;
  height:35px;
  width:35px;
}
</style>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>


<LINK REL="stylesheet" TYPE="text/css" HREF="VENTAS.CSS">
<!--#include file="comun/funcionescomunes.asp"-->
<!--#include file="includes/funcionesVBscript.asp"-->
<!--#include file="includes/cnn.inc"-->
<!--#include file="comun/comunQRY.asp"-->
<script language="javascript" src="includes/cal.js"></script>
<script language="jscript" type="text/jscript">
function AGRANDA() {
    top.parent.window.document.getElementById('body0').height = 580
}
var oldrow = 1

function dd2(ff) {	// LLENA TEXTBOX ADICIONALES AL COMUN
    // LOS DEL COMUN SON CODIGO Y DESCRIPCION
    var t = document.all.TABLA;
    var pos = parseInt(ff);
    dd(ff, 0);
}
</script>
<%
TDA = request.QueryString("TDA")
ope = request.QueryString("oper")

    cad =   " select * from movimdet where operacion =  '"&ope&"' and flag = '*' order by item "

'RESPONSE.WRITE(cAD)
'response.end
rs.open cad,cnn

if rs.recordcount <=0 then 
    response.Write("<br><br><br>")
    response.write("Operacion no Tiene lineas asignadas a NC ")
    RESPONSE.End
end if
RS.MOVEFIRsT
tda = rs("tienda")
Set RS2 = Server.CreateObject("ADODB.Recordset")
	RS2.ActiveConnection = Cnn
	RS2.CursorType       = 3 'CONST adOpenStatic = 3
	RS2.LockType         = 1 'CONST adReadOnly = 1
	RS2.CursorLocation   = 3 'CONST adUseClient = 3

	RS2.Open "select * from tiendas where codigo = '"&tda&"'", Cnn

%>

<body onload="AGRANDA()">
<center>
<form id ="thisForm" name= "thisForm" >
<table align="center" cellpadding="3" cellspacing="1" bordercolor='<%=application("color2") %>' border="1" id="TABLA" name="TABLA" 
bordercolorlight="<%=application("color2") %>"  bordercolordark="<%=application("color1") %>" >
    <tr>
        <td class="Estilo13">Operacion :&nbsp; </td>
        <td class="Estilo1" align="left"> <%=OPE%></td>
       
        <td class="Estilo13" align="right" >Tienda :&nbsp; </td>
        <td class="Estilo1" align="left" colspan="2"> <%=tda%>-<%=rs2("descripcion") %></td>
    </tr>
    <%rs.movefirst%>
    <%rs2.close%>
    <tr> 
       
        <td class="Estilo13">It </td>
        <td class="Estilo13">COD</td>
        <td class="Estilo13">Descripcion</td>
        <td class="Estilo13">Cant</td>
        <td class="Estilo13">Flag</td>
    </tr>
    <%con = 1 %>
    <%do while not rs.eof%>
    <tr> 
        <td class="Estilo1" align="center"> <input  type="text" readonly="readonly" value="<%=RS("item")%>" id="txt"/>	</td> 
        <td class="Estilo1" align="center"> <%=RS("codart")%></td> 
        <%cod = rs("codart")
        rs2.open "select top 1 descri  from View_ARTICULOS_TIENDA where codigo = '"&cod&"'" , cnn
        des = ""   
        if rs2.recordcount > 0 then 
            rs2.movefirst
            des = rs2("descri")
        end if    
            rs2.close
        %>
        <td class="Estilo1" align="left"> <%=trim(ucase(des))%></td> 
        <td class="Estilo1" align="center"> <%=RS("sale")%></td> 
        <td class="Estilo1" align="center"><input type="checkbox" value="<%=RS("flag")%>" checked id='ch<%=con%>' name='ch<%=con%>'/></td> 
       </tr>
       <%con = con + 1 %>
        <%rs.movenext%>     
    <%loop%>   
</table>
<center>
<input type="button" class="botonimagen" value="&nbsp;" onclick="graba()" /> 
</center>
<iframe  width="100%" src="" id="mirada" name="mirada" scrolling="yes" frameborder="1" height="40" align="middle" style="display:none" ></iframe>
</form>

<script language="jscript" type="text/jscript">
   


function graba() {
    ss = '<%=ope%>'
    tope = '<%=rs.recordcount %>'
    var aLin = Array()
    for (i = 1; i <= tope; i++) {
        if (eval("document.all.ch" + i + ".checked") == false)

        aLin[aLin.length] =  i.toString()
    }
    if (aLin.length == 0) {
        alert("No ha desmarcado ninguna Linea")
        return false;
    }
	aLin=document.getElementById("txt").value;
    cad = "comun/modiflag.asp?ope=" + trim(ss)
    cad += '&lin=' + aLin
	//alert(cad)

    //document.all.mirada.style.display='block'
    document.all.mirada.src = cad


}
</script>
</center>
</body>
</html>
