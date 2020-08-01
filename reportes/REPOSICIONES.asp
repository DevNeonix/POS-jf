<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<% tienda = Request.Cookies("tienda")("pos") %>
<% destda = Request.Cookies("tienda")("tda") %>
<!--#include file="../includes/Cnn.inc"-->


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=titulo%></title>
<link rel="stylesheet" type="text/css" href="../ventas.CSS" />

</head>
 <% IF  request.QueryString("EXCEL") = "1" THEN
  archivo = "c:\temp\cajaexcel.xls"
    Response.Charset = "UTF-8"
    Response.ContentType = "application/vnd.ms-excel" 
    Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo 
  END IF

   tem = request.QueryString("tem") 
   
   tda = request.QueryString("tda")%>
  
<body topmargin="0" leftmargin="0" rightmargin="0" border="0">

<table width="100%" align="center">
<%CAD =   " SELECT * FROM TIENDAS WHERE codigo= '"&tda&"'"
                   ' response.write(cad)
                  '  response.write("<br>")
            RS.OPEN CAD,CNN
            if rs.recordcount> 0 then 
                rs.movefirst
                destda =ucase(rs("descripcion"))
            else
                destda = "NINGUNA SELECCIONADA"
            end if
 %>
	<tr><td  align="center" class="estilo6" colspan="3">Lista de Artículos que requieren reposición tienda: <%=ucase(destda) %></td></tr>
	<tr><td colspan="3"><hr /></td></tr>
    <tr align="right" >
         <% CAD =   " SELECT * FROM TIENDAS WHERE ESTADO ='A' order by descripcion "
                  '  response.write(cad)
                  '  response.write("<br>")
                  rs.close
            RS.OPEN CAD,CNN
            IF rs.recordcount > 0 THEN rs.movefirst%>
        
         <td class="Estilo11" valign="middle" align="right"><label for="Radio">Tiendas:&nbsp;</label></td> 
         <td  class="Estilo12" align="left">
            <select  name="TT" id="TT">
                <option value = "" selected></option>
                <%do while not rs.eof %>
                    <option value="<%=TRIM(RS("CODIGO"))%>"><%=TRIM(RS("DESCRIPCION")) %></option>
                    <%rs.movenext %>
                <%loop %>
                <%rs.close %>
            </select>
        </td>
        
        <td align="right"><input type="checkbox" checked id="chk" name="chk"/>&nbsp;Excel</td>

        <td  class="estilo11">Temporada : &nbsp;&nbsp;</td>
        <td  align="left"><select  name="tempo" id="tempo">
                <option value = "0"></option>
                <option value = "">TODAS</option>
                <% CAD =   " SELECT * FROM temporadas WHERE ESTADO ='A' order by descripcion "
                  '  response.write(cad)
                  '  response.write("<br>")
                RS.OPEN CAD,CNN
                IF rs.recordcount > 0 THEN rs.movefirst

                do while not rs.eof %>
                    <option value=" <%=TRIM(RS("CODIGO"))%>" <%if ucase(trim(rs("codigo"))= ucase(trim(tem))) then %>selected<%end if%>><%=TRIM(RS("DESCRIPCION")) %></option>
                    <%rs.movenext %>
                <%loop %>
                <%rs.close %>
            </select>
        </td>
        <td align="left"><img src="../images/ok.gif" style="cursor:pointer" onclick="carga()" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
</table>


<%if trim(REQUEST.QueryString("op")) = "1" then 
        CAD =	" select DISTINCT CODIGO, DESCRI, stock as STK, (MINIMO-STOCK) AS CANT, " & _
                " MIN(SK_NSKDIS) AS SK_NSKDIS, CAST(round(LISTA1*1.18, 1) AS DECIMAL(8,2)) as pvp from  view_reponer_fabrica_2   " & _                
                " where stock <  minimo                                                                       " & _           
                " and tienda = '"&tda&"'                                                                                       " 
       if trim(ucase(tem)) <> "" then    cad = cad +    " and descri like '%"&tem&"%'                                          "
        
        cad = cad + " GROUP BY CODIGO, DESCRI, stock, (MINIMO-STOCK), round(LISTA1*1.18, 1)  order by codigo"		
       ' response.write(cad)
        RS.OPEN CAD ,Cnn
		
        If rs.eof or rs.bof then
	        Response.Write("Stock OK, no se requiere Reposición")		
	       ' Response.End
        else	
            rs.movefirst
        end if
%>



<table id="TABLA" align="center"  cellpadding="0" cellspacing="2" bordercolor='<%=application("color2") %>' border="0" >
<%'**************************%>
<%'LINEA DE CABECERA STANDAR %>
<%'**************************%>
<%cont = 1%>
<tr>
    <td align="center" class="Estilo8">IT</td>
	<td align="center" class="Estilo8">Codigo</td>
    <td align="center" class="Estilo8">Descripcion</td>
    <td align="center" class="Estilo8">Stk.</td>
    <td align="center" class="Estilo8">Cant.</td>
    <td align="center" class="Estilo8">Fab</td>
    <td align="center" class="Estilo8">P.V.P.</td>
</tr>
<%IF NOT RS.EOF THEN%>
    <%DO WHILE NOT RS.EOF%>
        <tr  bgcolor="<% if CONT mod 2  = 0 THEN 
                response.write(Application("color2"))
                else
	            response.write(Application("color1"))
	            end IF%>"
	             id="fila<%=Trim(Cstr(cont))%>">
            <td class="Estilo5" align="center" style="padding.right:6px;"><%=Cont%></td>
	        <td class="Estilo5" align="center" style="padding.right:6px;"><%=rs("codigo") %></td>
            <td class="Estilo5" align="center" style="padding.right:6px;"><%=ucase(rs("descri")) %></td>
            <td class="Estilo5" align="center" style="padding.right:6px;"><%=rs("STK") %></td>
            <td class="Estilo5" align="center" style="padding.right:6px;"><%=rs("cant") %></td>
            <td class="Estilo5" align="center" style="padding.right:6px;"><%=rs("SK_NSKDIS") %></td>
            <td class="Estilo5" align="center" style="padding.right:6px;"><%=formatnumber(cdbl(rs("pvp")),2,,,true) %></td>
            <%cont =cont +1 %>
        </tr>
        <%rs.movenext %>
    <%loop%>
<%end if %>
</table>
<center>
<form id="reponete" name="reponete" action="" method="post">
    <%if rs.recordcount >0 then %>
        <input type="submit" id="ENV" name="ENV" value="ENVIAR" onclick="correo()" />
    <%end if %>
</form>
</center>

<%end if %>
<script language="jscript" type="text/jscript">
function correo() {

    cad = "../comun/enviarepone.asp?tem=" + document.getElementById('tempo').value
    document.getElementById('reponete').action = cad
    document.getElementById('reponete').submit()
    }

    function carga() {
        cad = 'reposiciones.asp?tem=' + document.getElementById('tempo').value + '&OP=1&tda=' + document.all.TT.value 
        if (document.getElementById('chk').checked== true)
            cad += '&excel=1' 
 //  alert(cad)
    window.location.replace(cad)
}

</script>

</body>

</html>
