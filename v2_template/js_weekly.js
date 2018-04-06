<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js">

</script><script src="https://code.jquery.com/jquery-3.3.1.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/1.10.16/js/dataTables.bootstrap.min.js"></script>
<link rel="stylesheet" type="text/css" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
<link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.10.16/css/dataTables.bootstrap.min.css">

<script type="text/javascript" language="javascript">
    <!--position_data_chart-->

    google.charts.load('current', {'packages':['bar', 'corechart', 'line']});
    google.charts.setOnLoadCallback(showDailyReport);

    function showDailyReport(){
        document.getElementById("loading-progress").style.display = "none";
        document.getElementById("main-page").style.display = "block";
        drawChart('div-chart-bar-week', dataChartBar);
        drawChartPie();
        if(type_report == 'weeklyReport'){
            drawChartBar();
            drawChartLineWeekly()
        }else{
            drawChartLineMonthly()
        }
    }

    $(document).ready(function() {
        if (type_report == 'weeklyReport'){
            $('#div-table-weekly-rp').DataTable({
            "searching": false,
            "paging": false,
            data: dataTgReport,
            columns: [
               {title: 'TG'},
               {title: 'Total\n(Total/Resolved/Open)'},
               {title: 'Urgent\nTotal/Resolved/Open)'}
            ]
            });
        }
        else if(type_report == 'monthlyReport'){
            $("#table-summary-issue").hide();
            document.getElementById('title-chart-bar').innerText = MonthReport + " Issue Summary";
            document.getElementById('title-line-chart').innerText = "Monthly Total Issues";
            document.getElementById('sub-title-weekly-report').innerText = "Last month issue statistics";
            document.getElementById('title-chart-pie').innerText = "Last Month Issue by TG";
            $('#div-table-weekly-rp').DataTable({
                "searching": false,
                "paging": false,
                data: dataTgReport,
                columns: [
                   {title: 'TG'},
                   {title: 'Total\n(Total/Resolved/Open)'}
                ]
            });
        }
    });

    function drawChartPie(){
        var data = google.visualization.arrayToDataTable(dataChartPie);
        var options = {isStacked:true};
        var chart = new google.visualization.PieChart(document.getElementById('div-total-urgent-pie'));
        chart.draw(data, options);
    }

    function drawChart(id_type_chart, data_chart) {
        var data = new google.visualization.arrayToDataTable(data_chart);
        var options = {
            width: 500,
            height: 350,
            legend: { position: 'none' },
            bars: 'horizontal', // Required for Material Bar Charts.
            bar: { groupWidth: "90%" }
        };
        var chart = new google.charts.Bar(document.getElementById(id_type_chart));
        chart.draw(data, options);
    }

    function drawChartLineWeekly(){
        <!--data_for_line_chart_weekly-->
        var data = new google.visualization.DataTable();
        data.addColumn('number', 'Week');
        for(i=0; i < titleChartLine.length; i++){
            data.addColumn('number', titleChartLine[i]);
        }
        data.addRows(dataChartLine);
        var options = {
            width: 500,
            height: 350
        };
        var chart = new google.charts.Line(document.getElementById('div-chart-line-week'));
        chart.draw(data, google.charts.Line.convertOptions(options));
    }

    function drawChartLineMonthly(){
        <!--data_for_line_chart_monthly-->
        var data = new google.visualization.DataTable();
        data.addColumn('date', 'month');

        for(i=0; i < titleChartLine.length; i++){
            data.addColumn('number', titleChartLine[i]);
        }
        for(i=0; i < dataChartLine.length; i++){
            itemValue = dataChartLine[i];
            year = itemValue[0] % 10000;
            month = Math.floor(itemValue[0] / 10000) - 1;
            itemValue[0] = new Date(year, month);
            console.log(itemValue);
            dataChartLine[i] = itemValue;
        }
        var options = {
            width: 500,
            height: 350,
            hAxis: {
                format: 'MMM-yy'
            }
        };
        data.addRows(dataChartLine);
        var chart = new google.charts.Line(document.getElementById('div-chart-line-week'));
        chart.draw(data, google.charts.Line.convertOptions(options));
    }

    function drawChartBar(){
        <!--data_issue_team_by_prj-->
        var data = google.visualization.arrayToDataTable(dataChartStack);
        var chartAreaHeight = data.getNumberOfRows() * 10;
        var chartHeight = chartAreaHeight + 40;
        var options = {
                isStacked:true,
                height: chartHeight,
                chartArea: {
                    height: chartAreaHeight
                }
                };
        // Instantiate and draw the chart.
        var chart = new google.visualization.BarChart(document.getElementById('weekly-prj-chart-summary'));
        chart.draw(data, options);
    }
</script>