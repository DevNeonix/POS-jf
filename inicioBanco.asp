<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<!--#include file="includes/Cnn.inc"-->
<!--#include file="COMUN/FUNCIONESCOMUNES.ASP"-->
<body onkeyup="veri()">
<%tda = request.Form("tienda")
        
        rs.open "select codigo, descripcion from tiendas where codigo = '"&tda&"' ", cnn
         if rs.recordcount>0 then
            rs.movefirst
            ppp = RS("CODIGO")
            tda = rs("descripcion")
            RESPONSE.COOKIES("tienda")("bcp") = ppp
            RESPONSE.COOKIES("tienda")("bcpDES") = tda
        end if
  
'response.Write(cad)
rs.close
%>

 <script language="jscript" type="text/jscript">
    // top.window.document.all.titu.innerHTML = '' 
 </script>

<form action="" method="post" target="_self" id="frmIngreso" name="frmIngreso" >
<table align="center" border="0" cellpadding="0" cellspacing="2">
<tr><td colspan="2" style="height:50px"></td></tr>
<tr>
    <td align="center" class="Estilo8">Tienda :</td>
    <td align="center">
        <select id="tienda" name="tienda" class="login" onchange="cambia()">
        <option></option>
            <%CAD = "SELECT CODIGO,DESCRIPCION FROM TIENDAS where estado = 'a' ORDER BY DESCRIPCION"
            RS.OPEN CAD, CNN
            IF RS.RECORDCOUNT >0 THEN %>
                <%RS.MOVEFIRST
                DO WHILE NOT RS.EOF %>
                    <option value='<%=RS("CODIGO")%>' <%if request.form("tienda") = rs("codigo") then %>selected<%end if %>><%=RS("DESCRIPCION") %></option>
                <%RS.MOVENEXT %>
                <%LOOP %>
            <%END IF %>
        </select>
    </td>
</tr>
<tr><td colspan="2" style="height:50px"></td></tr>
<tr>
	<td colspan ="2" height="30" align="center" valign="middle" id="button1" name="button1"><img src="images/aceptar.jpg" width="67" onclick="Ingresar()" style="cursor:hand;" onmouseover="this.src='images/aceptar2.JPG'" onmouseout="this.src='images/aceptar.jpg'" height="19" />  
    </td>
</tr>
</table>

</form>

<script language="javascript">
   
//Poner foco al elemnto inicial
    document.getElementById("tienda").focus()
    function veri() {
        if (event.keyCode == 13)
            Ingresar()
    }
function Ingresar(){
    
    if (document.all.frmIngreso.tienda.selectedIndex <= 0) {
        //alert("Debe seleccionar la tienda")
        //document.all.button1.style.display='none'
        document.getElementById("tienda").selectedIndex = -1
        document.getElementById("tienda").focus()
        return false
    }
	else
        if (top.window.document.all.titu.innerHTML == '')
            frmIngreso.submit()
        else
           { top.window.document.all.titu.innerHTML = 'Verificador de Tarjetas BCP'
           window.location.replace('bcp.asp')
        }
	return true
}

function cambia() {

tt=document.getElementById("tienda").options(document.getElementById("tienda").selectedIndex).text

    top.window.document.all.tiend.innerHTML = document.getElementById("tienda").value
    top.window.document.all.tiend.innerHTML += ' - '
    top.window.document.all.tiend.innerHTML += tt
}


</script>
</body>
</html>
