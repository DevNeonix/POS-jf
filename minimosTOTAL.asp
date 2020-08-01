<%@ Language=VBScript %>
<%TDA = Request.Cookies("TIENDA")("POS")%>
<%Response.Buffer = TRUE %>
<!--#include file="includes/Cnn.inc"-->
<!--#include file="comun/funcionescomunes.asp"-->
<link REL="stylesheet" TYPE="text/css" HREF="ventas.CSS" >
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

<title><%=titulo%></title>
<link rel="stylesheet" type="text/css" href="ventas.CSS" />
<%tem= REQUEST.QueryString("Tem") %>
</head>
<script type="text/jscript" language="jscript">
function calcHeight() {
        document.getElementById("tit").style.height = document.body.offsetHeight - ((document.getElementById("Table1").offsetHeight + document.getElementById("Table2").offsetHeight) * 1.4) - 20
    }
</script>
<body topmargin="0" leftmargin="0" rightmargin="0" border="0">


 <table id="Table1" align="center"  bordercolor="#FFFFFF"  bgcolor="<%=Application("color2")%>"  cellpadding="2"  cellspacing="4"  border="0">
	       
            <tr valign="middle" >
                 <td class="Estilo11" valign="middle" align="right" rowspan="2">
                    <label for="Radio">Excel:&nbsp;</label></td> 
                    <td><input id="excel" type="checkbox"  name="excel" /></td>
                    
    	       
                   <td class="Estilo11" valign="middle" align="right" style="display:none">Temporada:&nbsp;</td> 
        <td  class="Estilo12" align="left" style="display:none">
                <select  name="tempo" id="tempo">
                    <option value = "" selected>TODAS</option>
                    <%CAD =   " SELECT * FROM temporadas WHERE ESTADO ='A' order by descripcion "
              '  response.write(cad)
              '  response.write("<br>")
                RS.OPEN CAD,CNN
                    IF rs.recordcount > 0 THEN rs.movefirst
                    do while not rs.eof %>
                        <option value=" <%=TRIM(RS("CODIGO"))%>"><%=TRIM(RS("DESCRIPCION")) %></option>
                        <%rs.movenext %>
                    <%loop %>
                
                </select>
            </td>
            
            <td><img src="images/ok.gif" onClick="reemplaza()" style="cursor:pointer;"/></td>       
        </tr>
        
    
        <%RS.CLOSE %>
    </table>

<table id="Table2" align="center"  cellpadding="0" cellspacing="1" bordercolor='<%=application("color2") %>' border="1"  width="800px" >
 <tr>
     <td class="Estilo5" align="left" width="70px" >CODIGO</td>
     <td class="Estilo5" align="center" width="300px">DESCRIPCION</td>
     <td class="Estilo5" align="center" width="60px" >AS</td>    
     <td class="Estilo5" align="center" width="60px" >PO</td>
     <td class="Estilo5" align="center" width="60px" >CH</td>
     <td class="Estilo5" align="center" width="60px" >SI</td>
     <td class="Estilo5" align="center" width="60px" >AR</td>
     <td class="Estilo5" align="center" width="60px" >T2</td>
     <td class="Estilo5" align="center" width="60px" >OUT</td>
     <td class="Estilo5" align="center" width="60px" >EM</td>
     <td class="Estilo5" align="center" width="60px" >JF</td>    
    </tr>
</table>
<p align="center">
<iframe id="tit" name="tit" src="" onload="calcHeight()" width="850px" scrolling="yes"  frameborder=1></iframe>
</p>
<script language="jscript" type="text/jscript">
    kad  = 'stkTOTAL.asp?tem=' + '<%=request.querystring("tem")%>'
    kad += '&excel=' + '<%=request.querystring("excel")%>'  
    
    if (trim('<%=request.querystring("excel")%>')=='')
        document.getElementById('tit').src = kad
    else
        window.open(kad)
function reemplaza() {
    cad = 'minimosTOTAL.asp?tem=' + (document.all.tempo.value)
    if (document.all.excel.checked == true)
        cad += '&excel=1'
    else
        cad += '&excel=' 
    window.location.replace(cad)
}
</script>
</BODY>
</HTML>
