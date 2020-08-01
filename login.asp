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

<%  RESPONSE.COOKIES("tienda")("pos")    = ""
    RESPONSE.COOKIES("tienda")("TDA")    = ""
    RESPONSE.COOKIES("tienda")("USR")    = ""
    RESPONSE.COOKIES("tienda")("PERFIL") = ""

    usr = request.Form("txtusr")
  pwd = request.Form("txtpwd")
  tda = request.Form("tienda")
  msg = ""
  cad = "select * from usuarios where usuario = '"&usr&"' and clave = '"&pwd&"'"
  rs.open cad,cnn
  if rs.recordcount>0 then
        rs.movefirst
        usuario = rs("nombres")
        PERFIL = rs("PERFIL")
        rs.close
        rs.open "select codigo, descripcion from tiendas where codigo = '"&tda&"' ", cnn
         if rs.recordcount>0 then
            rs.movefirst
            tda = rs("descripcion")
            ppp = RS("CODIGO")
            RESPONSE.COOKIES("tienda")("pos") = ppp
            RESPONSE.COOKIES("tienda")("TDA") = TDA
            RESPONSE.COOKIES("tienda")("USR") = USR
            RESPONSE.COOKIES("tienda")("PERFIL") = PERFIL
        else
            RESPONSE.COOKIES("tienda")("USR") = ""
        end if
  else
    
    if len(trim(usr)) > 0 and len(trim(pwd)) then
        rs.close
        cad = "select * from usuarios where usuario = '"&usr&"'"
        rs.open cad,cnn
        if rs.recordcount<=0 then
            msg = "Usuario no Registrado"
        else
            msg = "Password Incorrecto"
        end if
    end if
  end if
'response.Write(cad)
rs.close
CAD = "EXEC SP_TIPODECAMBIO"
RS.Open CAD, Cnn

    if rs.recordcount < 1 then%>
      <script language="jscript" type="text/jscript">
          alert("No hay tipo de cambio ingresado\nAvise a Sistemas")
      </script>
    <%response.end
    end if


rs.close
 %>

 <script language="jscript" type="text/jscript">
 msg = '<%=trim(msg) %>'
 if (msg.length >0)
    alert(msg)
 </script>

<form action="" method="post" target="_self" id="frmIngreso" name="frmIngreso" >
<table align="center" border="0" cellpadding="0" cellspacing="2">
<tr><td colspan ="2" height="60"></td></tr>
<tr>
    <td align="center" class="Estilo8">Usuario :</td>
    <td align="center" width="100">
        <input id="txtusr" name="txtusr" type="text" maxlength="10" tabindex="0" class="login" value='<%=request.form("txtusr") %>'/>
    </td>
</tr>
<tr>
    <td align="center" class="Estilo8">Password :</td>
    <td align="center">
        <input id="txtpwd" name="txtpwd" type="password" maxlength="10" class="login" value='<%=request.form("txtpwd") %>'/>
    </td>
</tr>
<tr>
    <td align="center" class="Estilo8">Tienda :</td>
    <td align="center">
        <select id="tienda" name="tienda" class="login">
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

<tr>
	<td colspan ="2" height="30" align="center" valign="middle" id="button1" name="button1"><img src="images/aceptar.jpg" width="67" onclick="Ingresar()" style="cursor:hand;" onmouseover="this.src='images/aceptar2.JPG'" onmouseout="this.src='images/aceptar.jpg'" height="19" />  
    </td>
</tr>
</table>

</form>

<script language="javascript">
    uus = '<%=trim(usuario)%>' 
    if (uus.length > 0) 
    {   tta = '<%=cstr(ppp) %>' + '-'+ '<%=tda%>'
        top.window.document.all.usuariu.innerHTML = uus
        top.window.document.all.tiend.style.color = '#FF0000'
     //   top.window.document.all.usuar.innerHTML = '<%=ucase(request.Form("txtusr"))%>'
        top.window.document.all.usero.value = '<%=ucase(request.Form("txtusr"))%>'
        top.window.document.all.tiend.innerHTML = tta
        


        if (parseInt('<%=PERFIL%>',10)== 0)
            top.window.document.all.perfidia.style.display = 'block'
                  
        if (parseInt('<%=PERFIL%>', 10) == 0 || parseInt('<%=PERFIL%>', 10) == 4)
            top.window.document.all.amante.style.display = 'block'
       
       //  muestra el menu
        top.window.document.all.menu.style.display = 'table'
        // restringe opciones a las venddoras....
        var str = '1 CSABA SISTEMAS MALAGA JFARAH';
        var n = str.indexOf(trim(parent.document.getElementById('usero').value));

        if (n >= 0) {
            parent.document.getElementById('sho0').style.display = 'block'
            parent.document.getElementById('sho1').style.display = 'block'
        }
        else {
            parent.document.getElementById('sho0').style.display = 'none'
            parent.document.getElementById('sho1').style.display = 'none'
        }


        window.location.replace("blanco.htm")
    }

//Poner foco al elemnto inicial
    document.getElementById("txtusr").focus()
    function veri() {
        if (event.keyCode == 13)
            Ingresar()
    }


function Ingresar(){
    var _user = trim(document.all.frmIngreso.txtusr.value)
    var _pwd = trim(document.all.frmIngreso.txtpwd.value)
	document.all.button1.style.display='block'
	
	if(_user.length==0){
		alert("Debe ingresar su Usuario")
		//document.all.button1.style.display='none'
		document.getElementById("txtusr").focus()
		return false
	}
	else if(_pwd.length==0){
		//alert("Debe ingresar su Password")
		//document.all.button1.style.display='none'
		document.getElementById("txtpwd").focus()
		return false
    }
    else if (document.all.frmIngreso.tienda.selectedIndex <= 0) {
        //alert("Debe seleccionar la tienda")
        //document.all.button1.style.display='none'
        document.getElementById("tienda").selectedIndex = -1
        document.getElementById("tienda").focus()
        return false
    }
    else
      frmIngreso.submit()

	return true

}
</script>
</body>
</html>
