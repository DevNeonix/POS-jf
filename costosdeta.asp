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
<%
    cad =   " SELECT * FROM  VIEW_costos order by 1"

' response.write(cad)
RS.OPEN CAD,CNN

Set RS2 = Server.CreateObject("ADODB.Recordset")
	RS2.CursorLocation   = 3
	RS2.CursorType       = 3    
	RS2.LockType         = 1 	




If rs.recordcount<=0 then    
  response.End  
else     
  rs.movefirst
end if
%>
<body onload="agranda()">
<iframe width="100%" src="" id="mirada" name="mirada"  scrolling="yes" frameborder="3" height="100" align="middle" style="display:none">
</iframe>

<table align="center" cellpadding="0" cellspacing="1" bordercolor='<%=application("color2") %>' border="1" id="listado" name="listado" >
   <tr> <td align="center" class="Estilo8">IT</td>
        <td align="center" class="Estilo8">&nbsp;<%=trim(rs.fields(0).name) %>&nbsp;</td>
	    <td align="center" class="Estilo8"><%=rs.fields(2).name %></td>
        <td align="center" class="Estilo8"><%=rs.fields(3).name %></td>
        <td align="center" class="Estilo8"><%=rs.fields(1).name %></td>
        <td align="center" class="Estilo8"><%=rs.fields(6).name %></td>
        <td align="center" class="Estilo8"><%=rs.fields(7).name %></td>
   </tr>
    <%I=1%>
    <%do while not rs.eof %>
        <tr bgcolor="<%IF I MOD 2 = 0 THEN RESPONSE.WRITE(application("color1")) ELSE RESPONSE.WRITE(application("color2"))%>" class="Estilo19" >
            <td>&nbsp;<%=I%>&nbsp;</td>
            <td align="center" width="15%"><input id ="cd<%=I%>" value="<%=trim(rs.fields.ITEM(0)) %>" class="Estilo21" style=" border-bottom-style:hidden; border-color:transparent; border-bottom-width:0" readonly tabindex="-1" /></td>
            <td width="45%"><%=trim(rs.fields.ITEM(2))%></td>
            <td><%=rs.fields.ITEM(3) %></td>
            <td width="10%"><input id ="LS<%=I%>" value="<%=formatnumber(rs.fields.ITEM(1),2,,,true)%>" class="Estilo20" onkeyup=(this.value=toInt(this.value)) onblur="actualiza('<%=I%>')"  size="10"/></td>
            <td><select id="ln<%=i%>" name="ln<%=i%>" value="" onchange="actualiza('<%=I%>')">
                    <option value=""></option>
                    <%rs2.open "select codigo , descripcion from lineas where estado = 'a'", cnn %>
                    <%do while not rs2.eof%>
                        <option value="<%=rs2.fields.ITEM(0)%>" <%if rs2.fields.ITEM(0) = rs.fields.ITEM(4) then %>selected <%end if%>><%=trim(ucase(rs2.fields.ITEM(1)))%></option>
                        <%rs2.movenext %>
                    <%loop %>
             
                    <%rs2.close%>
                </select>
            </td>
            <td><select id="cl<%=i%>" name="cl<%=i%>" value="" onchange="actualiza('<%=I%>')">
                    <option value=""></option>
                    <%rs2.open "select codigo , descripcion from clases where estado = 'a'", cnn %>
                    <%do while not rs2.eof%>
                        <option value="<%=rs2.fields.ITEM(0)%>" <%if rs2.fields.ITEM(0) = rs.fields.ITEM(5) then %>selected <%end if%>><%=trim(ucase(rs2.fields.ITEM(1)))%></option>
                        <%rs2.movenext %>
                    <%loop %>
             
                    <%rs2.close%>
                </select>
            </td>
       </tr> 
       <%I= I + 1%>
       <%rs.movenext %>
       <%Response.flush() %>
   <%loop %>
</table>
<script language="jscript" type="text/jscript">
    parent.window.document.getElementById('wait').style.display='none'
</script>

</body>


<script language="jscript" type="text/jscript">
function agranda() {
    top.parent.window.document.getElementById('body0').height = 480
}

function actualiza( opc) {

    op= parseInt(opc,10)
    COD = eval("document.all.cd" + op + ".value")
    pre = eval("document.all.LS" + op + ".value")
    ln = eval("document.all.ln" + op + ".value")
    cl = eval("document.all.cl" + op + ".value")
    cad = 'comun/updatecosto.asp?cod=' + COD
    cad += '&pre=' + pre
    cad += '&opc=' + op
    cad += '&lin=' + ln
    cad += '&cla=' + cl
    //alert(cad)
    document.all.mirada.src= cad

    
}

</script>

</html>
