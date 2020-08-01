<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
    <style type="text/css">
        .style1
        {
            font-family: Tahoma;
            font-size: 12px;
            color: #ffffff;
            font-weight: 600;
            height: 118px;
        }
        .style2
        {
            height: 118px;
        }
    </style>
<script src="SpryAssets/SpryMenuBar.js" type="text/javascript"></script>
<link href="SpryAssets/SpryMenuBarHorizontal.css" rel="stylesheet" type="text/css" />
<script type="text/jscript" language="jscript">
function calcHeight()
{
  //find the height of the internal page
  var the_height=
    document.getElementById('body0').contentWindow.
      document.body.scrollHeight;

  //change the height of the iframe
  document.getElementById('body0').height=
      the_height+20;
}
function centro(opc)
{	var cad
    
	if (opc==11)
	{	cad= "../../agetex/index.asp"
		
	}
	else
	{	cad="blanco.htm"
		}
	document.all.body0.src =cad
}


</script>
</head>
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<body align="center">

<table align="center" border="0" cellpadding="0" cellspacing="0">
<tr><td align="center">
<img src="images/logo_tit.jpg" width="212" height="70" align="center" />
</td>
</tr>
</table>

<table align="center" border="0" cellpadding="0" cellspacing="0" id="menu" style="display:none">
<tr><td align="center">
<ul id="MenuBar1" class="MenuBarHorizontal">
  <li><a class="MenuBarItemSubmenu" href="#">Mantenimientos</a>
    <ul>
      <li><a href="javascript:centro(11)">Monedas</a></li>
      <li><a href="#">Parametros</a>
          <ul>
              <li><a href="javascript:centro(121)">Impuesto</a></li>
              <li><a href="javascript:centro(122)">Tipo de Cambio</a></li>
           </ul>
      </li>
      <li><a href="#">Auxiliares</a>
      		<ul>
              <li><a href="javascript:centro(131)">Tipo de Cliente</a></li>
              <li><a href="javascript:centro(132)">Tipo Documento</a></li>
              <li><a href="javascript:centro(133)">Tipo Movimiento</a></li>
              <li><a href="javascript:centro(134)">Tipo de Pago</a></li>
            </ul>
      </li>
      <li><a href="javascript:centro(15)">Clientes</a></li>          
      <li><a href="javascript:centro(16)">Proveedores</a></li>      
      <li><a href="javascript:centro(17)">Artículos</a></li> 
      <li><a href="javascript:centro(18)">Genéricos</a></li>           
    </ul>
  </li>
  <li><a href="javascript:centro(21)">Movimientos</a></li>
  <li><a class="MenuBarItemSubmenu" href="#">Consultas</a>
    <ul>
      <li><a class="MenuBarItemSubmenu" href="javascript:centro(31)">Item 3.1</a>
        <ul>
          <li><a href="javascript:centro(311)">Item 3.1.1</a></li>
          <li><a href="javascript:centro(312)">Item 3.1.2</a></li>
        </ul>
      </li>
      <li><a href="javascript:centro(32)">Item 3.2</a></li>
      <li><a href="javascript:centro(33)">Item 3.3</a></li>
    </ul>
  </li>
  <li><a href="#">Procesos</a></li>
  <li><a href="#">Sistemas</a></li>
 
</ul>
</td>
</tr>
</table>



<center>
<iframe width="90%" onLoad="calcHeight();" src="login.asp" id="body0" name="body0"  scrolling="NO" frameborder="0" height="1" align="middle">
</iframe>
</center>
<script type="text/javascript">
var MenuBar1 = new Spry.Widget.MenuBar("MenuBar1", {imgDown:"SpryAssets/SpryMenuBarDownHover.gif", imgRight:"SpryAssets/SpryMenuBarRightHover.gif"});
</script>
</body>


</html>
