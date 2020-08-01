<%@ Language=VBScript %>
<% Response.Buffer = true %>
<%Session.LCID=2058%>
<!--#include file="includes/Cnn.inc"-->
<%
   
if request.QueryString("ejecuta") = "si" then
cad =""    
cad = cad +"SET DATEFORMAT DMY;                                                                                                                               "
cad = cad +"DROP TABLE TMPCTO;SET ANSI_WARNINGs OFF;"
cad = cad +"select left(CODART,5) as CODGRU, DESgru as DESCRI, SUM (SALE) AS CANT, isnull(v.COSTO,0) as costo, (isnull(V.COSTO,0)*SUM (SALE)) AS CTO_TOT,     "
cad = cad +"isnull(SUM(A.PVP-dct-IGV),0) AS TOT_DOC, ((isnull(V.COSTO,0)*SUM (SALE)) - SUM(A.PVP-dct-IGV) )* -1 AS DIF , FECHA, TIENDA                        "
cad = cad +"INTO TMPCTO                                                                                                                                       "
cad = cad +"FROM VIEW_VENTAS_ARTICULO A full outer JOIN VIEW_COSTOS V ON left(ltrim(rtrim(a.CODART)),5) COLLATE Modern_Spanish_CI_AI = V.CODIGO               "
cad = cad +"WHERE SALE > 0 and isnull(A.pvp,0) > 0                                                                                                            "
cad = cad +"GROUP BY left(codart,5), A.FECHA, DESgru, V.COSTO, TIENDA                                                                                         "
    response.Write(cad)
    cnn.execute(cad)
    %>
<script type="text/javascript">
    alert("datos actualizados")
    this.location.replace = "./actualizacostos.asp?ejecuta=no"
</script>
    <%
end if

 %>
<!DOCTYPE html/>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <style type="text/css">
        *
        {
            font-family:Sans-Serif
            }
    </style>
</head>
<body>
    Actualiza Costos &nbsp; <button style="background:url(images/check.png);width:30px;height:30px;padding:0px;" onclick="ejecutar()"></button>
    <script type="text/javascript">
        function ejecutar() {
            this.location.replace="./actualizacostos.asp?ejecuta=si"
        }
    </script>
</body>
</html>
