<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<% tienda = Request.Cookies("tienda")("pos") %>
<!--#include file="../includes/Cnn.inc"-->
<!--#include file="../comun/funcionescomunes.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link REL="stylesheet" TYPE="text/css" HREF="../ventas.CSS" >
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Nuevo Cliente</title>
</head>
<script type="text/jscript" language="jscript">
    function modifica() {
        window.resizeTo(500, 250)
        window.moveTo(0, 0)
        this.window.focus()
    }
</script>
<%cod = request.querystring("cod") 
cnn.execute "delete clientes where cliente = '"&cod&"' and direccion = '' "%>

<body onLoad="modifica()" >
<table	align="center" cellpadding="1" cellspacing="0"  bgcolor="WHITE" border="0" >
<tr valign="middle">
<td colspan="4" class="Estilo3" align="center" valign="middle" style="height:25px">CLIENTE NUEVO</td> </tr>
<td colspan="4"  align="center" valign="middle" style="height:25px">&nbsp;</td> </tr>
              <tr valign="middle"> 
                <td bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  CODIGO : &nbsp;</td>
                <td bgcolor="WHITE" valign="middle"> 
                  <input type="text" id="COD" name="COD"  class="Estilo12" maxlength="11" size="12" value="<%=request.querystring("cod") %>"/>
                </td>
                <td bgcolor="<%=(Application("borde"))%>" align="right" class="Estilo3">NOMBRE :&nbsp;</td>
                <td bgcolor="#FFFFFF" > 
                  <input type="text" id="DES" name="DES" class="Estilo12"  size="40"  />
                </td>
              </tr>
              <tr valign="middle"> 
                <td bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  DIRECCION :&nbsp;&nbsp;  </td>
                <td bgcolor="WHITE"valign="middle" colspan="3"> 
                  <input type="text" id="DIR" name="DIR" class="Estilo12" maxlength="200" size="93" style="width:100%" />
                </td>
            </tr>
            <tr style="display:block"><td width="10%" bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  CORREO :  &nbsp;</td>
                <td bgcolor="WHITE" valign="middle" colspan="3"> 
                  <input type="text" id="COR" name="COR"  class="Estilo12" maxlength="50" size="93" style="width:100%" />
                </td>
            </tr>
            <tr style="display:block">
                <td bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  TELEFONO : &nbsp;</td>
                <td bgcolor="WHITE" valign="middle"> 
                  <input type="text" id="TEL" name="TEL"  class="Estilo12" maxlength="12"  size="12" style="width:100%" onKeyUp="this.value=toInt(this.value)"  />
                </td>
                <td  bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  TIPO CLIENTE :&nbsp; </td>
                <td bgcolor="WHITE"  valign="middle"> 
                  <select type="text" id="TIP" name="TIP" class="Estilo12" style="width:100%"  >
                       <option value=""></option> 
                       <%CAD =	" select tg_cclave as CODIGO,   " & _
                                " TG_CDESCRI AS DESCRIPCION     " & _
                                " from RSFACCAR..AL0012TABL     " & _
                                " where tg_ccod = '67'          " & _
                                " order by 2                    " 
                       
                       rs.open CAD ,cnn
                       rs.movefirst
                       do while not rs.eof%>
                       <option value="<%=rs("codigo") %>"><%=rs("codigo") %>-<%=rs("descripcion") %></option>
                        <%rs.movenext
                       loop%>
                  </select>
                  <%rs.close %>                  
                </td>
              </tr>
              <tr valign="middle" style="display:none"> 
                <td bgcolor="<%=(Application("barra"))%>" align= "right" class="Estilo3">
                  DCTO. 1 :&nbsp;</td>
                <td bgcolor="WHITE" valign="middle"> 
                  <input type="text" id="DT1" name="DT1"  maxlength="5" class="Estilo12" value="0"
                  onkeyup="this.value=toInt(this.value)"   size="12" style="width:100%"  />
                </td>
                <td bgcolor="<%=(Application("borde"))%>" align="right" class="Estilo3">DCTO. 2 :&nbsp;</td>
                <td bgcolor="#FFFFFF"> 
                  <input type="text" id="DT2" name="DT2"  maxlength="5" value="0" class="Estilo12" onKeyUp="this.value=toInt(this.value)" style="width:100%" size="15" />
                </td>
              </tr>
              <tr>
                  <td colspan="4" align="center" valign="middle" >
                    <img src="../images/ok.gif" border="0" style="cursor:pointer" onClick="valida()"/>
                  </td> 
             </tr>
</table>

</body>
<script type="text/jscript" language="jscript">

function valida() {
    if (document.all.COD.value == "") 
    {   alert('Debe Informar el Código');
        document.all.COD.focus();
        return false;
    }
    
    /*if (trim(document.all.COR.value) == '') {
        alert("Debe informar el coreo por Orden SUNAT")
        return false;
    }*/

    cli = trim(document.all.COD.value)
    if (cli.length < 8) {
        cli = strzero(cli, 11)
        document.all.COD.value = cli
    }
    if (trim(document.all.DES.value) == "")
    {   alert('Debe Informar EL NOMBRE DEL CLIENTE');
    document.all.DES.focus();
        return false;
    }
    else
        document.all.DES.value = Left(ltrim(rtrim(document.all.DES.value)), 100)

    if (trim(document.all.TIP.value) == "") {
        tip = 'C'
    }
    else
        tip = trim(document.all.TIP.value)
    /*if (isEmail(trim(document.all.COR.value), 1) == false)
    { return false; }*/
    
    var cad
    cad  = "../comun/insernewcli.asp?cod=" + Left(trim(document.all.COD.value),11);
    cad += '&des=' + document.all.DES.value;
    cad += '&DIR=' + toAlpha(ltrim(document.all.DIR.value));
    cad += '&TEL=' + trim(document.all.TEL.value);
    cad += '&cor=' + trim(document.all.COR.value);
    cad += '&tip=' + tip;
    if (trim(document.all.DT1.value) == '')
        document.all.DT1.value = 0
    cad += '&DT1=' + trim(document.all.DT1.value);
    if (trim(document.all.DT2.value) == '')
        document.all.DT2.value = 0
    cad += '&DT2=' + trim(document.all.DT2.value);
    cad += '&opcion=vta'


   window.location.replace(cad)
}
</script>

</html>
