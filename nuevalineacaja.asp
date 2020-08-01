<%
ope = request.QueryString("ope")
tda = request.QueryString("tda")
%>
<!DOCTYPE HTML>
<html>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
</HEAD>
<body>
<table>
	<tr> 
        <td class="Estilo12" width="10%">Moneda</td>
        <td class="Estilo12" width="20%">Tipo</td>
        <td class="Estilo12" width="10%">Monto</td>
        <td class="Estilo12" width="40%">Doc Ref</td>
    </tr>
	<tr> 
        <td class="Estilo1" align="left">
            <select  class="INPUTS" id="monuevo" name="monuevo"  value="" style="text-align:center;width:100%">
                <option value=""></option>
                <option value="MN">MN</option>
                <option value="US">US</option>
           </select> 
       </td>
       <td class="Estilo1" align="left">
            <select  class="INPUTS" id="tipnuevo" name="tipnuevo"  value="" style="text-align:center;width:100%">
                <option value=""></option>
                <option value="SOL">SOLES</option>
                <option value="DOL">DOLARES</option>
                <option value="MAS">MASTERCARD</option>
                <option value="VIS">VISA</option>
                <option value="NCR">NOTA DE CREDITO</option>
           </select> 
       </td> 
       <td class="Estilo1" align="right"><input id="solnuevo" name="solnuevo"  class="INPUTS" style="width:100%" value='' onfocus="" /></td> 
       <td class="Estilo1" align="left" ><input id="obsnuevo" name="obsnuevo"  class="INPUTS" style="width:100%" value='' onfocus="" /></td> 
        <td><img src="../images/check.png" onclick="g()" title="guardar" tag="guardar"/></td>
    </tr>
</table>
<script>
function g() {
    cad  = "comun/insertlineacaja.asp?ope=<%=ope%>&tda=<%=tda%>"
    cad += '&mon=' + document.all.monuevo.value
    cad += '&tipo=' + document.all.tipnuevo.value
    cad += '&monto=' + document.all.solnuevo.value
    cad += '&obs=' + document.all.obsnuevo.value
    window.open(cad);
	setTimeout(function(){parent.location.reload();},1500);
}
</script>
</body>
</html>