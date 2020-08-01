<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<% tienda = Request.Cookies("tienda")("pos") %>
<!--#include file="includes/Cnn.inc"-->
<!--#include file="COMUN/FUNCIONESCOMUNES.ASP"-->
<!--#include file="COMUN/COMUNqry.ASP"-->
<script type="text/jscript" language="jscript">
function calcHeight()
{
  //find the height of the internal page
  var the_height=
    document.getElementById('mirada').contentWindow.
      document.body.scrollHeight;

  //change the height of the iframe
  document.getElementById('mirada').height=
      the_height+20;
}

</script>
<% 'CAD =   " SELECT * FROM documento WHERE   " & _
   '         " cia = '"&TIENDA&"' AND codigo in ('BL','FC') "
   '       '  response.write(cad)
   '       '  response.write("<br>")
   ' RS.OPEN CAD,CNN
   ' bol = "000-00000000"
   ' fac = "111-11111111"
   ' IF rs.recordcount > 0 THEN 
   '     rs.movefirst
   '     ss = cdbl (rs("correl"))   +1
   '     ss = RIGHT("0000000"+TRIM(CSTR(SS)),7)
   '     if rs("codigo")="BL" then
   '         BOL = RS("SERIE")&"-"& SS
   '     else
   '         fac =  RS("SERIE")&"-"& SS
   '     end if
   '     RS.MOVENEXT
   '     ss = cdbl (rs("correl"))   +1
   '     ss = RIGHT("0000000"+TRIM(CSTR(SS)),7)
   '     if rs("codigo")="BL" then
   '         BOL = RS("SERIE")&"-"& SS
   '     else
   '         fac =  RS("SERIE")&"-"& SS
   '     end if    
   ' end if
   ' RS.CLOSE
   FAC = "-"
   BOL = "-"
     %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
    
</head>

<body onkeyup="enter()" onload="document.all.CLI.focus();">
<table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="0"  cellspacing="0"  border="0" align="center" width="100%" >
    <tr valign="middle"><td colspan="40"><h2 style="text-align: center;color:#F09;font-family: sans-serif;">Registro de Documentos de Contingencia</h2></td></tr>
    <tr valign="middle" >
        <td class="Estilo11" valign="middle" align="right" rowspan="1"><label for="Radio">Documento:&nbsp;</label></td> 
        <td  class="Estilo12" align="left"  rowspan="1"><input type="Radio" name="miRadio" id="miRadio" value="BL" checked onclick="DOCUM()">Boleta(Manual)
        <br /><input type="Radio" name="miRadio" id="miRadio" value="FC"  onclick="DOCUM()">Factura(Manual)</td>

        <%
            aBol = split(bol,"-")
            aFac = split(fac,"-")
        %>

        <td>  
            <div id="DBOL" style="display: inline;">
                <input id="BOL" name="BOL" maxlength="4" class="Estilo12" value="<%=aBol(0)%>" size="20" onchange="if(this.value.length <3 ){if(confirm('La serie no puede tener menos de 3 caracteres \n desea completarlo con ceros?')){this.value =  strzero(trim(this.value),4)}}" style=" padding:0.2em 0.2em;width: 35px"/>
                <span>-</span>
                <input id="BOLDOC" name="BOLDOC" maxlength="7" class="Estilo12" value="<%=aBol(1)%>" onkeyup="" onchange="this.value = Left(strzero(trim(toInt(this.value)),7),7)" size="20" style=" padding:0.2em 0.2em "/>
            </div>
            <br />
            <div id="DFAC" style="display: inline;">
                <input id="FAC" name="FAC" maxlength="4" class="Estilo12" value="<%=aFac(0)%>" size="20" onchange="if(this.value.length <3 ){if(confirm('La serie no puede tener menos de 3 caracteres \n desea completarlo con ceros?')){this.value =  strzero(trim(this.value),4)}}" style=" padding:0.2em 0.2em;width: 35px"/>
                <span>-</span>
                <input id="FACDOC" name="FACDOC" maxlength="7" class="Estilo12" value="<%=aFac(1)%>" onkeyup="" onchange="this.value = Left(strzero(trim(toInt(this.value)),7),7)" size="20" style=" padding:0.2em 0.2em "/>
            </div>
        </td>

        <td class="Estilo11" valign="middle" align="right"  rowspan="2">Cliente :&nbsp;</td> 
        <td align="left"><input type="text" name="CLI" id="CLI" value="" class="Estilo12" onchange="cliente(this.value)"  size="20" maxlength="11">
        <input type="text" name="DES" id="DES" value="" maxlength="50" size="50" class="Estilo12" readonly tabindex="-1">
        <br /><input type="text" name="DIR" id="DIR" value="" maxlength="100" size="80" class="Estilo12"  readonly tabindex="-1"></td>

    <td align="center" rowspan="1" valign="middle" style="border:0; background-image:images/lupa2.JPG"><input id="sik" name="sik" maxlength="10" class="Estilo14"   /><br />
    Busca Stock</td>
    <td align="left" rowspan="1" valign="middle" style="border:0"><img src="images/search.JPG" border="0" style="cursor:pointer" onclick="LOOK()"></td>
 

        <td class="Estilo11" valign="middle" align="right" rowspan="2"><label for="Radio2">Moneda:&nbsp;&nbsp;</label></td> 
        <td  class="Estilo12" align="left"  rowspan="2"><input type="Radio" name="Radio2" id="Radio2" value="MN" checked>Soles
        <br /><input type="Radio" name="Radio2" id="Radio2" value="US" disabled><font color="gainsboro"> D&oacute;lares</font></td>
       
    </tr>
   
</table>
<iframe src="" allowScriptAccess='always'  id="mirada" name="mirada" style="display:none"></iframe>

</body>
<script type="text/jscript" language="jscript">
    document.all.CLI.focus();
   
    if (document.all.miRadio[0].checked == true){
        document.all.DFAC.style.display = 'none'
    }

function DOCUM() {
    parent.window.frames[1].window.TOTALES()    
    if (document.all.miRadio[0].checked == true) {
        document.all.DFAC.style.display = 'none'
        document.all.DBOL.style.display = 'block'


        parent.window.frames[2].window.document.all.ig.style.display = 'none'
        parent.window.frames[2].window.document.all.iig.style.display = 'none'
        parent.window.frames[2].window.document.all.st.style.display = 'none'
        parent.window.frames[2].window.document.all.sst.style.display = 'none'
        parent.window.frames[2].window.document.all.Table1.style.display = 'none'
        parent.window.frames[2].window.document.all.Table2.style.display = 'block'
    }
    else {
        document.all.DBOL.style.display = 'none'
        document.all.DFAC.style.display = 'block'
        parent.window.frames[2].window.document.all.ig.style.display = 'block'
        parent.window.frames[2].window.document.all.iig.style.display = 'block'
        parent.window.frames[2].window.document.all.st.style.display = 'block'
        parent.window.frames[2].window.document.all.sst.style.display = 'block'
        parent.window.frames[2].window.document.all.Table1.style.display = 'block'
        parent.window.frames[2].window.document.all.Table2.style.display = 'none'
    }
    }
    function cliente(dato) {

      ss = document.all.CLI.value
    cad = "bake/bakecliente.asp?pos=" + trim(dato)
 //   document.all.mirada.style.display='block'
    document.all.mirada.src = cad

    if (ss.length < 8) {
        document.all.CLI.value = strzero(trim(document.all.CLI.value), 11)
        dato = strzero(dato, 11)
    }

}
function enter() {
    if (event.keyCode == 13) {
        ss = document.all.CLI.value
        if (ss.length < 8) 
            document.all.CLI.value = trim(strzero(document.all.CLI.value,11))
        cliente(document.all.CLI.value)
    }
}

function valida() {
    
    if (trim(document.all.CLI.value) == '') {
        alert("Debe ingresar un cliente para poder emitir documento")
        return false
    }
    else 
    {  // alert(parent.window.frames.length)
        parent.window.frames[1].window.location.replace("ventadeta.asp")
    }

}
function NEWCLI() {
     var opc = "directories=no,height=600,width=600, ";
    opc = opc + "hotkeys=no,location=no,";
    opc = opc + "menubar=no,resizable=yes,";
    opc = opc + "left=0,top=0,scrollbars=no,";
    opc = opc + "status=no,titlebar=no,toolbar=no,";

    cad = "help/newcliente.asp?COD=" +  document.all.CLI.value
    //window.parent.document.all.CLI.value = ''
    //alert(cad)
    window.open(cad, "", opc)
}

function LOOK(){

    var opc = "directories=no,height=600,width=600, ";
    opc = opc + "hotkeys=no,location=no,";
    opc = opc + "menubar=no,resizable=yes,";
    opc = opc + "left=20,top=20,scrollbars=no,";
    opc = opc + "status=no,titlebar=no,toolbar=no,";

cad = "help/hlpprendastotal.asp?pos=" + trim(document.all.sik.value)
    window.open(cad, "", opc)
}


</script>

</html>
