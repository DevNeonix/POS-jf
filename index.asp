<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Buffer = false%>
<%Session.LCID=2057%>
<%tienda  = Request.Cookies("tienda")("pos") %>
<%Usuario = Request.Cookies("tienda")("usr")%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

 <!-- Favicon -->
  <link href="imgAGES/favicon.png" rel="icon" type="image/png">
<title>Jacinta Fernandez - POS</title>
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
<!--#include file="includes/Cnn.inc"-->
<!--#include file="comun/funcionescomunes.asp"-->
<script src="includes/jquery.js"></script>
<script type="text/jscript" language="jscript">
    var opc = "directories=no,height=400,width=600, ";
    opc = opc + "hotkeys=no,location=no,";
    opc = opc + "menubar=no,resizable=no,";
    opc = opc + "left=0,top=0,scrollbars=no,";
    opc = opc + "status=no,titlebar=no,toolbar=no,";

function bcp() { 
    window.open('bcp.asp','',opc)   
    }

function tbl() {
    window.open('http://servsql_bkp/Reports_SQL/Pages/Report.aspx?ItemPath=%2fJF_Stk_Vts%2fTabla_Dinamica')
}
function calcHeight() {
    aa = Right(document.getElementById('body0').src, 9)
    bb = Right(document.getElementById('body0').src, 11)

    if (aa != 'venta.asp' && bb!='precios.asp' ) 
    { //find the height of the internal page
        var the_height =
        document.getElementById('body0').contentWindow.
          document.body.scrollHeight;

        //change the height of the iframe
        document.getElementById('body0').height =
          the_height + 120;
          $("#body0").height($(document).height()-($("#menu").offset().top+$("#brand").height()));
    }
    else 
    {   document.getElementById('body0').height = 425    }

}
function centro(opc)
{	var cad

    if (opc == 11)                // parametros
    { cad = "PARAMETROS.ASP" }
    else if (opc == 121)         // TIPOS DE CLIENTE
    { cad = "tabla.asp?pos=67&TIT=TIPOS DE CLIENTE" }
    else if (opc == 122)         // TIPOS DE DOCUMENTO
    { cad = "documentos.asp" }
    else if (opc == 123)        // clase DE MOVIMIENTOS
    { cad = "tIPMOV.asp?pos=" }
    else if (opc == 124)     // TIPO DE PAGO
    { cad = "TIPPAG.ASP" }
    else if (opc == 125)     // TIPOS DE TARJETA DE CREDITO
    { cad = "tabla.asp?pos=25&tit=TIPOS DE TARJETA DE CREDITO" }
    else if (opc == 126)                // MONEDAS
    { cad = "tabla.asp?pos=03&tit=MONEDAS" }
    else if (opc == 127)     // TIPOS NOTA DE CREDITO
    { cad = "tabla.asp?pos=76&tit=CLASES DE NOTA DE CREDITO" }
    else if (opc == 128)     // motivo transferencia
    { cad = "MOTITRANS.asp" }
    else if (opc == 129)     // temporadas
    { cad = "temporadas.asp" }
    else if (opc == 130)     // temporadas
    { cad = "CLASES.asp" }
    else if (opc == 131)     // temporadas
    { cad = "lineas.asp" }

    else if (opc == 13)     // clientes
    { cad = "clientes.asp" }
    else if (opc == 14)     // ARTICULOS
    { cad = "ARTICULOS.asp" }
    else if (opc == 15)     // Actualizacion de stock
    { cad = "indexbcp.asp" }

    else if (opc == 21)     // ventas
    { cad = "venta.asp" }
    else if (opc == 22)     // devoluciones - notas de credito
    { cad = "notacredito.asp" }
    else if (opc == 23)     // anulacion de documentos
    { cad = "anulaciones.asp" }
    else if (opc == 24)     // ingresos con guia
    { cad = "ingresos.asp" }
    else if (opc == 25)     // TRANSFERENCIAS
    { cad = "transferencias.asp" }
    else if (opc == 26)     // notas de credito espaciales
    { cad = "espaciales.asp" }
    else if (opc == 27)        // Documentos Electronicos
    { cad = "efactReal.asp" }
    else if (opc == 31)     // Reporte de caja del día
    { cad = "REPORTES/cajadia.asp" }
    else if (opc == 32)     // Reporte de ventas por producto
    { cad = "REPORTES/resumenventas.asp" }
    else if (opc == 33)     // solicitud de reposicion de articulos con stock minimo
    { cad = "REPORTES/REPOSICIONES.asp" }
    else if (opc == 34)     // Stock de tienda
    { cad = "REPORTES/stocks.asp" }
    else if (opc == 35)     // Stock de tienda
    { cad = "REPORTES/stocks2.asp" }
    else if (opc == 36)     // Reimpresión de documentos
    { cad = "REPORTES/reimpresiones.asp" }
    else if (opc == 37)     // movimientos por artículo
    { cad = "REPORTES/KARDEX.asp" }
    else if (opc == 38)     // VISUALIZACION DE DOCUMENTOS POR OPERACION
    { cad = "REPORTES/OPERACIONES.asp" }
    else if (opc == 39)     // VENTAS POR ARTICULO
    { cad = "REPORTES/VTS_CODIGO.asp" }
    else if (opc == 30)     // ventas por grupo 5 caracteres
    { cad = "REPORTES/VTS_GRUPO.asp" }
    else if (opc == 301)     // ventas por grupo 5 caracteres
    { cad = "REPORTES/VTS_TABLA.asp" }
    else if (opc == 302)     // REPORTE DETALLADO PARA EL COMERCIO
    { cad = "REPORTES/COMERCIO.asp" }
    else if (opc == 303)     // REPORTE comparativo de ventas con años anteriores
    { cad = "REPORTES/comparativo.asp" }
    else if (opc == 304)     // Reporte de caja del día
    { cad = "REPORTES/VENTASdia.asp" }

    else if (opc == 41)     // manejar las listas de precio
    { cad = "precios.asp" }
    else if (opc == 42)     // mantenimiento de usuarios
    { cad = "usuarios.asp" }
    else if (opc == 43)     // mantenimiento de tiendas
    { cad = "tiendas.asp" }
    else if (opc == 44)     // mantenimientos de descuentos y porcentajes
    { cad = "descuentos.asp" }
    else if (opc == 45)     // ajuste de inventario
    { cad = "ajustes.asp" }
    else if (opc == 46)     // interfase de ventas al softcom
    { cad = "CONTABLE.asp" }
    else if (opc == 47)     // interfase de ventas al softcom
    { cad = "STOCK_PLANTILLA.asp" }
    else if (opc == 48)     // mantenimientos de usuarios autorizados por tienda
    { cad = "reportes/sku.asp" }
    else if (opc == 49)     // interfase de ventas al softcom
    { cad = "STOCK_PLANTILLA_tot.asp" }
    else if (opc == 490)     // edicion de documentos
    {
        USS = trim(top.window.document.all.usero.value)
        //alert(USS)
        if (USS == 'CSABA' || USS == 'SISTEMAS' || USS == 'JFARAH' || USS == '1')
        { cad = "ediciones.asp" }
        else
        { cad = "blanco.htm" }
    }
    else if (opc == 491)     // ingresa costos
    {
        USS = trim(top.window.document.all.usero.value)
        //alert(USS)
        if (USS == 'CSABA' || USS == 'SISTEMAS' || USS == 'JFARAH' || USS == '1')
        { cad = "costos.asp" }
        else
        { cad = "blanco.htm" }
    }
    else if (opc == 492)     // reporte de costos
    {
        USS = trim(top.window.document.all.usero.value)
        //alert(USS)
        if (USS == 'CSABA' || USS == 'SISTEMAS' || USS == 'JFARAH' || USS == '1')
        { cad = "reportes/prncostos.asp" }
        else
        { cad = "blanco.htm" }
    }
    else if (opc == 493)     // exporta para tiendas
    { cad = "exportaGUIA.asp" }
    else if (opc == 494)     // Actualiza minimos a todas las tiendas
    { cad = "MINIMOS_TOT.asp" }
    else if (opc == 51)     // Actualiza minimos a todas las tiendas
    {
        USS = trim(top.window.document.all.usero.value)
        //alert(USS)
        if (USS == 'CSABA' || USS == 'SISTEMAS' || USS == 'JFARAH' || USS == '1')
        { cad = "oc/jfoc.asp" }
        else
        { cad = "blanco.htm" }
    } else if (opc == 520)     // Actualiza minimos a todas las tiendas
    {
        USS = trim(top.window.document.all.usero.value)
        //alert(USS)
        if (USS == 'CSABA' || USS == 'SISTEMAS' || USS == 'JFARAH' || USS == '1')
        { cad = "actualizacostos.asp" }
        else
        { cad = "blanco.htm" }
    } else if (opc == 701)     // Actualiza minimos a todas las tiendas
    {
        USS = trim(top.window.document.all.usero.value)
        //alert(USS)
        if (USS == 'CSABA' || USS == 'SISTEMAS' || USS == 'JFARAH' || USS == '1')
        { cad = "reportes/reporteripleyhead.asp" }
        else
        { cad = "blanco.htm" }
    } else if (opc == 702)     // Actualiza minimos a todas las tiendas
    {
        USS = trim(top.window.document.all.usero.value)
        //alert(USS)
        if (USS == 'CSABA' || USS == 'SISTEMAS' || USS == 'JFARAH' || USS == '1')
        { cad = "reportes/reporteripleyheadtrial.asp" }
        else
        { cad = "blanco.htm" }
    }else if (opc == 703)     // Actualiza minimos a todas las tiendas
    {
        USS = trim(top.window.document.all.usero.value)
        //alert(USS)
        cad = "contingencia.asp"
       
    }
    else
    { cad = "blanco.htm" }
    
    //document.all.body0.focus();  
	document.all.body0.src =cad
}
</script>

</head>
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<body align="center">
<% CAD = "EXEC SP_TIPODECAMBIO"
RS.Open CAD, Cnn%>
<center>
<table align="center" id="brand" border="0" cellpadding="0" cellspacing="5" width="95%">
<tr>
   	<td align="left" class="Estilo7" width="25%" valign="bottom" rowspan="1" colspan="2"><%=DATE()%></td>
	<td align="center" width="50%" rowspan="2">
		<img src="images/logo_tit.jpg" width="212" height="70" align="center" />
	</td>
	<td align="right" class="Estilo7" width="10%" valign="bottom"><FONT COLOR=RED>Tienda : </FONT></td>
	<td align="right" class="Estilo7" width="15%" id="tiend" valign="bottom" colspan="2"></td>
</tr>
<tr>
	<td align="left" class="Estilo7" width="5%" valign="top" rowspan="1">D&oacute;lar :</Td>
	<td valign="top"  align="left"><input id="cambio" name="cambio" value='<%=RS("bajo")%>' size="3" class="Estilo7" style="border:0; text-align:left" /></td>
	<td align="right" class="Estilo7" width="15%" valign="top">Usuario : </td>
    <td align="right" class="Estilo7" width="1%" id="usuar" valign="top" style="display:none;color:red"><input id="usero" name="usero" style = "display:none;color:green" ></td>
	<td align="right" class="Estilo7" width="5%" id="usuariu" valign="top"></td>
</tr>

</table>
<%permitidos = "1 CSABA SISTEMAS"%>
<center>
    <table align="center" border="0" cellpadding="0" cellspacing="5" id="menu" style="display:none">
<tr>
<td align="center">
    <ul id="MenuBar1" class="MenuBarHorizontal">
       
      
        <li id="sho0" name="sho0" style="display:block" ><a  href="aa.asp">NUEVO</a></li>
        <li id="sho1" name="sho1" style="display:block"><a class="MenuBarItemSubmenu" href="#">Mantenimientos</a>
            <ul>
                <li><a href="javascript:centro(11)">Parametros</a>   </li>
                <li><a href="#">Auxiliares</a>
                    <ul>
                        <li><a href="javascript:centro(121)">Tipo de Cliente</a></li>
                        <li><a href="javascript:centro(122)">Documentos</a></li>
                        <li><a href="javascript:centro(123)">Clase Movimiento</a></li>
                        <li><a href="javascript:centro(124)">Tipo de Pago</a></li>
                        <li><a href="javascript:centro(125)">Tarjeta Credito</a></li>
                        <li><a href="javascript:centro(126)">Monedas</a></li>
                        <li><a href="javascript:centro(127)">Clase Nota Credito</a></li>
                        <li><a href="javascript:centro(128)">Motivo Transferencia</a></li>
                        <li><a href="javascript:centro(129)">Temporadas</a></li>
                        <li><a href="javascript:centro(130)">Clases</a></li>
                        <li><a href="javascript:centro(131)">Lineas</a></li>
               
                    </ul>
                </li>
               
                <li><a href="javascript:centro(14)">Artículos</a></li>
                <li><a href="javascript:bcp()">Oferta BCP</a></li>
            </ul>
        </li>
      
        
        <li><a class="MenuBarItemSubmenu" href="#">Movimientos</a>
            <ul>  
                <li><a href="javascript:centro(13)">Clientes</a></li>
                <li><a href="javascript:centro(21)">Ventas</a></li>
                <li><a href="javascript:centro(22)">Nota de Credito</a></li>            
                <li><a href="javascript:centro(23)">Anulaciones</a></li>
                <li><a href="javascript:centro(24)">Ingresos GS</a></li>    
                <li><a href="javascript:centro(25)">Transferencias - Salida</a></li>    
                <li><a href="javascript:centro(26)">Vales Especiales</a></li> 
                <li><a href="javascript:centro(27)">Doc. Real eFact</a></li> 
                <li><a href="javascript:centro(703)">Contingencia / Manual</a></li>
            </ul>
        </li>
        
        <li><a class="MenuBarItemSubmenu" href="#">Reportes</a>
        <ul>  
            <li><a href="javascript:centro(31)">Caja Diario</a></li>
            <li><a href="javascript:centro(32)">Resumen Ventas/Doc</a></li>            
            <li><a href="javascript:centro(33)">Reposiciones</a></li>
            <li><a href="javascript:centro(34)">Stock Total</a></li>   
            <li><a href="javascript:centro(35)">Con Stock</a></li>    
            <li><a href="javascript:centro(36)">Reimpresiones</a></li> 
            <li><a href="javascript:centro(37)">Deta Articulos</a></li>  
            <li><a href="javascript:centro(38)">Operaciones</a></li>                 
            <li><a href="javascript:centro(39)">Vtas. Articulo</a></li> 
            <li><a href="javascript:centro(30)">Vtas. Grupo</a></li>  
            <li><a href="javascript:centro(301)">Vtas. TIENDAS</a></li>       
            <li><a href="javascript:centro(302)">Detalle El Comercio</a></li> 
            <li><a href="javascript:centro(303)">Comparativos</a></li>     
            <li><a href="javascript:centro(304)">VENTAS Dia</a></li>                
        </ul>
        </li>
        
        <li id="perfidia" style="display:none"><a href="#">Utilitarios</a>
        <ul>  
            <li><a href="javascript:centro(41)">Precios</a></li>
            <li><a href="javascript:centro(42)">Usuarios</a></li>
            <li><a href="javascript:centro(43)">Tiendas</a></li>         
            <li><a href="javascript:centro(44)">Descuentos</a></li>
            <li><a href="javascript:centro(45)">Ajustes - Ingreso</a></li>  
            <li><a href="javascript:centro(46)">Contabilidad</a></li> 
            <li><a href="javascript:centro(47)">Elimina Plantilla</a></li>
            <li><a href="javascript:centro(49)">Elimina Plantilla Todos</a></li>
            <li><a href="javascript:centro(48)">Reporte Ripley</a></li> 
        
                <li><a  class="MenuBarItemSubmenu" href="#">Gerencia</a>
                    <ul>
                        <li><a href="javascript:centro(490)">Ediciones</a></li> 
                        <li><a href="javascript:centro(491)">Edita Costos</a></li> 
                        <li><a href="javascript:centro(492)">Reporte Costos</a></li> 
                    </ul>
                </li> 
            <li><a href="javascript:centro(493)">Exporta Franquicia</a></li> 
            <li><a href="javascript:centro(494)">Minimos Tiendas</a></li> 
            <li><a href="javascript:centro(520)">Actualiza Costos</a></li>
            <li><a href="javascript:centro(701)">Reporte Ripley</a></li>
            <li><a href="javascript:centro(702)">Reporte Trial</a></li>
            <li><a href="javascript:centro(703)">Contingencia / Manual</a></li>
        </ul>
        </li>
  <li id="amante" style="display:none"><a href="#" class="MenuBarItemSubmenu">Logistica</a>
        <ul>  
            <li><a href="javascript:centro(51)">Orden de Compra</a></li>
            
        </ul>
        </li>
 
    </ul>

     
   
    </ul>
</td>
</tr>
</table>
</center>


<iframe width="100%" onLoad="calcHeight();" src="login.asp" id="body0" name="body0"  scrolling="yes" frameborder="1" height="1" align="middle">
</iframe>
</center>
<script type="text/javascript">
    var MenuBar1 = new Spry.Widget.MenuBar("MenuBar1", { imgDown: "SpryAssets/SpryMenuBarDownHover.gif", imgRight: "SpryAssets/SpryMenuBarRightHover.gif" });
</script>

</body>


</html>
