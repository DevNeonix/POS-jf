<!DOCTYPE html>
<html>
<head>
	<!--#include file="../includes/cnn.inc"-->
	<meta charset="utf-8">
	<title></title>
	<style type="text/css">
		*{
			font-family: sans-serif;
		}
		label{
			display: block;
			color:magenta;
			padding: 5px
		}
		input{
			display: block;

			width: 100%
		}
		img{
			transition: all .2s;
			box-sizing: border-box;
		}
		img:hover{
			border: 1px solid magenta;
		}
	</style>
	<script type="text/javascript">
		function trim(checkString)
{ // QUITA TODOS LOS BLANCOS DE LA CADENA
var	newString = "";
if (checkString == "")
    return newString;
    // loop through string character by character
  for (i = 0; i < checkString.length; i++) 
  {	ch = checkString.substring(i, i+1);
    // quita blancos
    if (ch != ' ' ) 
    {  newString += ch;      }
  }
 
  return newString;
}
		function isEmail(ele, op)
{  // valida campos tipo e-mail
	ele = trim(ele)
   // return false if e-mail field is blank.
   // Si op == 1, el mail puede ir en blanco
   if (ele == "" && op != 1 ) 
   {
      alert("\n El campo E-MAIL está en blanco\n\nIngrese la dirección e-mail.")
      return false; 
   }
   if (ele != '')
   // return false if e-mail field does not contain a '@' and '.' .
   if (ele.indexOf ('@',0) == -1 || 
       ele.indexOf ('.',0) == -1)
   {    alert("El E-MAIL  require  una  \"@\"y un \".\"necesariamente.\n\nRegistre adecuadamente el e-mail.")
      return false;
  }
  if (ele.length < 7) {
      alert("le agradeceremos un mail valido");
      return false;
      }
return true;      
}
	</script>
</head>
<body>
	<%

	documento = request.querystring("doc")
	ticket    = request.querystring("ticket")
	email    = request.querystring("email")
	envio    = request.querystring("envio")
	if trim(envio) = "1" then
		body = "<a href='http://jf.elmodelador.com.pe/apijf/public/index.php/show?ticket="&ticket&"&tipo=pdf'>Ver Documento</a><iframe src='http://jf.elmodelador.com.pe/apijf/public/index.php/show?ticket="&ticket&"&tipo=pdf' frameborder='0' width='655' height='550' marginheight='0' marginwidth='0' id='pdf'></iframe>"
		'response.write(body)
		cnn.execute "exec [sp_envia_mail] 'JACINTA FERNANDEZ, Su factura "&documento&" fue enviada ','"&replace(body,"'","''")&"','"&email&"' "
		response.write("Correo en cola por sistemas@elmodelador.com.pe, revise por favor su bandeja de entrada. ")
		%>
		<script type="text/javascript">
			setTimeout(function(){window.close()},5000)
		</script>
		<%
		response.end
	end if
	%>
	<label>
		Documento
		<input type="text" readonly="readonly" style="color:#eee;background: #aaa" value="<%=trim(documento)%>">
	</label>
	<label>
		Ticket
		<input type="text" readonly="readonly" style="color:#eee;background: #aaa" value="<%=ticket%>">
	</label>
	<label>
		Email
		<input type="text" value="<%=email%>" id="txtemail">
	</label>
	<label>
		<img src="../images/mail.png" onclick="envia()" style="display: block;margin:0 auto;cursor: pointer;">
	</label>
	<script type="text/javascript">
		function envia(){
			if(valida()){
				window.location.href=window.location.href+"&envio=1&email="+txtemail.value
			}
		}
		function valida(){
			var txtemail = document.getElementById("txtemail")
			if(!isEmail(txtemail.value)){
				return false;
			}
			return true
		}
	</script>
</body>
</html>