<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<% tienda = Request.Cookies("tienda")("pos") %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="COMUN/FUNCIONESCOMUNES.ASP"-->
<!--#include file="COMUN/COMUNqry.ASP"-->
<!--#include file="includes/Cnn.inc"-->
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>
<%pos = trim(request.QueryString("pos") )
TDA = left(pos,2)
' la primera vez que se muestra la pantalla no tiene querystring y se deb mostratr en blanco la pantalla
if len(trim(pos)) <=0 then  response.End  

tipo = right(pos,2)
pos = left(pos,4)
cad = "select 1"
if  pos = "TTPG" then
        cad =   " SELECT DISTINCT  ' TODAS ' AS TIENDA, A1.GRUPO,   " & _
                " A1.DESCRI, A2.LISTA1                              " & _
                " FROM  VIEW_GRUPOS AS A1                           " & _
                " INNER JOIN ARTICULOS A2 ON A1.CODIGO = A2.CODIGO  " & _  
                " GROUP BY A1.GRUPO, A1.DESCRI, A2.LISTA1           " & _
                " ORDER BY 2                                        "
elseif pos =  "TTPP" then
        cad =   " SELECT DISTINCT ' TODAS ' AS TIENDA, A1.CODIGO,   " & _
                " A3.DESCRI, A2.LISTA1                              " & _
                " FROM  VIEW_GRUPOS AS A1                           " & _
                " INNER JOIN ARTICULOS A2 ON A1.CODIGO = A2.CODIGO  " & _  
                " INNER JOIN VIEW_ARTICULOS_TIENDA A3               " & _
                " ON A1.CODIGO = A3.CODIGO                          " & _     
                " GROUP BY  A1.CODIGO, A3.DESCRI, A2.LISTA1         " & _
                " ORDER BY 2                                       "
ELSE
    if right(pos,2)= "PG"  then
         cad =  " SELECT DISTINCT '"&TDA&"' as TIENDA, A1.GRUPO, A1.DESCRI,       " & _
                " isnull(A2.LISTA1,0) AS LISTA1 FROM  VIEW_GRUPOS AS A1                     " & _
                " INNER JOIN ARTICULOS A2 ON A1.CODIGO = A2.CODIGO      " & _
                " WHERE A1.TIENDA = '"&TDA&"' and A2.TIENDA = '"&TDA&"'" & _
                " GROUP BY A1.GRUPO, A1.DESCRI, A2.LISTA1               " & _
                " ORDER BY 2                                            " 
    else
        cad =   " SELECT DISTINCT A1.TIENDA, A1.CODIGO, A3.DESCRI,      " & _
                " ISNULL(A2.LISTA1,0) AS LISTA1 FROM  VIEW_GRUPOS AS A1                     " & _
                " INNER JOIN ARTICULOS A2 ON A1.CODIGO = A2.CODIGO      " & _
                " INNER JOIN VIEW_ARTICULOS_TIENDA A3                   " & _
                " ON A1.CODIGO = A3.CODIGO                              " & _
                " WHERE A1.TIENDA = '"&tda&"' and A2.TIENDA = '"&TDA&"' " & _
                " GROUP BY A1.TIENDA, A1.CODIGO, A3.DESCRI, A2.LISTA1   " & _
                " ORDER BY 2                                            "
    end if
end if

' response.write(cad)
RS.OPEN CAD,CNN
If rs.recordcount<=0 then    
  response.End  
else     
  rs.movefirst
end if
%>
<body onload="agranda()">
<iframe width="100%" src="" id="mirada" name="mirada"  scrolling="yes" frameborder="3" height="100" align="middle" style="display:none">
</iframe>

<table align="center" cellpadding="0" cellspacing="1" bordercolor='<%=application("color2") %>' border="0" id="listado" name="listado" style="display:none" >
   <tr> <td align="center" class="Estilo8">IT</td>
        <td align="center" class="Estilo8">&nbsp;<%=trim(rs.fields(0).name) %>&nbsp;</td>
	    <td align="center" class="Estilo8"><%=rs.fields(1).name %></td>
        <td align="center" class="Estilo8"><%=rs.fields(2).name %></td>
        <td align="center" class="Estilo8"><%=rs.fields(3).name %></td>
   </tr>
    <%I=0 %>
    <%do while not rs.eof %>
        <tr bgcolor="<%IF I MOD 2 = 0 THEN RESPONSE.WRITE(application("color1")) ELSE RESPONSE.WRITE(application("color2"))%>" class="Estilo19">
            <td>&nbsp;<%=I%>&nbsp;</td>
            <td align="center"><input id="TD<%=i%>" value="<%=trim(rs.fields.ITEM(0)) %>" class="Estilo21" size="5" readonly tabindex="-1"/></td>
            <td><input id="CD<%=i%>" value="<%=trim(rs.fields.ITEM(1)) %>" class="Estilo21" size="11"  readonly tabindex="-1"/></td>
            <td><%=rs.fields.ITEM(2) %></td>
            <%IF ISNULL(RS.FIELDS.ITEM(3)) THEN CCC= "0.00" ELSE CCC= formatnumber(rs.fields.ITEM(3),2,,,true)%>
            <td><input id ="LS<%=I%>" value="<%=CCC%>" class="Estilo20" onkeyup=(this.value=toInt(this.value))  onblur="actualiza(this.value, '<%=I%>')" size="10"/></td>
       </tr> 
       <%I= I + 1%>
       <%rs.movenext %>
   <%loop %>
</table>

<table align="center" cellpadding="0" cellspacing="1" bordercolor='<%=application("color2") %>' border="0" id="individual" name="individual" style="display:none" >
   <tr> <td align="center" class="Estilo8">TDA</td>
        <td align="center" class="Estilo8"><%if right(pos,2)= "PG" then%>GRUPO<%ELSE%>CODIGO<%END IF %></td>
	    <td align="center" class="Estilo8">Descripcion</td>
        <td align="center" class="Estilo8">Precio</td>
   </tr>
   <tr>
        <td><input id ="TD9999" value="<%=TDA%>" class="Estilo20" size="5" readonly tabindex="-1"/></td>
        <td><input id ="CD9999" value="" class="Estilo20" size="10" onchange="bakeCODIGO(this.value, '9999')"/></td>
        <td><input id ="DS9999" value="" class="Estilo20" readonly tabindex="-1" /></td>
        <td><input id ="LS9999" value="" class="Estilo20"  size="10" onkeyup=(this.value=toInt(this.value)) maxlength="6"/></td>
   </tr>
   <tr><td colspan="3"><img src="images/ok.gif" onclick="actualiza(document.all.LS9999.value, '9999')" style="cursor:pointer;"/></td></tr>
</table>

</body>


<script language="jscript" type="text/jscript">
//alert()
function agranda() {
    top.parent.window.document.getElementById('body0').height = 480
}

pantalla = '<%=trim(tipo) %>'
if (pantalla == 'TL')
    document.all.listado.style.display = 'block'
else {
    document.all.individual.style.display = 'block';
    document.all.CD9999.focus();
}

function bakeCODIGO(cod, op) {

    cad = "bake/bakeprecios.asp?cod=" + cod + "&op=" + op
    //document.all.mirada.style.display = 'block'
    document.all.mirada.src = cad

return true
}

function actualiza(pre, opc) {

    op = parseInt(opc,10)
    tda = eval("document.all.TD" + op + ".value")
    cod = eval("document.all.CD" + op + ".value")
    
    cad = 'comun/updateprecio.asp?tda=' + tda
    cad += '&cod=' + cod
    cad += '&pre=' + pre
    cad += '&opc=' + op
    //document.all.mirada.style.display = 'block' 
    document.all.mirada.src= cad
}

</script>

</html>
