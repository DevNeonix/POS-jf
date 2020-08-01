<p><!--#include file="../includes/Cnn.inc"--></p>
<!DOCTYPE html>
<html>
<head>
	<title>Informe</title>
	<link rel="stylesheet" type="text/css" href="style.css">
	<meta charset="utf-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
</head>
<body>
	
	<div class="chart-container">
  <h2 class="h4">Reporte de ventas por mes de Jacinta Fernandez</h2>
    <canvas id="chart"></canvas>
</div>
<%

rs.open "SELECT MONTH(FECDOC) AS MES, SUM(total) AS TOTAL FROM JACINTA.dbo.movimcab WHERE YEAR(FECDOC) = 2019 AND CODDOC IN ('BL','FC') GROUP BY MONTH(FECDOC) ORDER BY 1;",cnn
response.write(rs.recordCount)
'cnn.close
'response.write("<h2>Registros guardados</h2>")
if rs.recordCount > 0 then
	rs.movefirst
end if

%>
<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.9.3/Chart.js"></script>
<script type="text/javascript">

	var data = {
  labels: ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"],
  datasets: [{
    label: "Ventas",
    backgroundColor: "rgba(228, 22, 90, 1)",
    borderColor: "rgba(255,99,132,1)",
    borderWidth: 2,
    hoverBackgroundColor: "rgba(255,99,132,0.4)",
    hoverBorderColor: "rgba(255,99,132,1)",
    data: [
      <%for i=0 to rs.recordcount-1%>
        <%=replace(rs("total"),",","")&","%>
        <%rs.movenext%>
      <%next%>
      
    ],
  }]
};

var option = {
  scales: {
    yAxes: [{
      stacked: true,
      gridLines: {
        display: true,
        color: "rgba(255,99,132,0.2)"
      }
    }],
    xAxes: [{
      gridLines: {
        display: false
      }
    }]
  }
};

Chart.Bar('chart', {
  options: option,
  data: data
});
</script>
</body>
</html>
