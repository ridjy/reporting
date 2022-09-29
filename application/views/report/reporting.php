<!DOCTYPE html >
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <title><?php echo $titre; ?></title>
  <meta name="Robots" content="all"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link rel="shortcut icon" href="<?php //echo img_url('logo.png'); ?>">
  <link rel="stylesheet" media="screen" type="text/css" title="Design" href="<?php echo css_url('bouton_style'); ?>" /> 
  <link href="<?php echo css_url('bootstrap.min')?>" rel="stylesheet" type="text/css">
    <style>
  
  </style>
    <script type="text/javascript">
       /*var tableToExcel = (function () {
           var uri = 'data:application/vnd.ms-excel;base64,'
   , template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:Name>"temp.xls"</x:Name><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>'
   , base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) }
   , format = function (s, c,x) { return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; }) }
           return function (table, name) {
               if (!table.nodeType) table = document.getElementById(table)
               var ctx = { worksheet: name || 'Worksheet', table: table.innerHTML }
               window.location.href = uri + base64(format(template, ctx));

           }
       })()*/
   
   var tableToExcel = (function () {
           var uri = 'data:application/vnd.ms-excel;base64,'
   , template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:Name>"Statistique_globale.xls"</x:Name><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>'
   , base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) }
   , format = function (s, c,x) { return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; }) }
           return function (table, name, filename) {
               if (!table.nodeType) table = document.getElementById(table)
               var ctx = { worksheet: name || 'Worksheet', table: table.innerHTML }
               /*window.location.href = uri + base64(format(template, ctx))*/
               document.getElementById("dlink").href = uri + base64(format(template, ctx));
              document.getElementById("dlink").download = filename;
              document.getElementById("dlink").click();

           }
       })()

   </script>
</head>
<body >
  <div class="content">
            <div class="container-fluid"> 
                <div class="row">
                  <div class="col-md-3">
                     <a href="<?php echo site_url('reporting/index') ; ?>"><button class="btn btn-warning">Retour </button></a>
                  </div> 
                  <div class="col-md-4">
                    <p><?php echo "Heure actuelle: ".date('H:i'); ?></p>
                  </div>
                </div>
                <br/>
                <div class="row">
                    <section class="col-md-12 table-responsive">
                        <table id="report" style="text-align:center" class="table table-bordered table-striped table-condensed">
                                <tr style="height:110px;">
                                    <td colspan="2"><?php echo $date;//echo date('d-m-Y');?></td>
                                    <td colspan="8"><img src="<?php echo img_url('Oceancall Group.png'); ?>" ></td>
                                    <td colspan="8"><img src="<?php echo img_url('ZEOP-LOGO.jpg'); ?>" ></td>
                                </tr>  
                                <tr style="background-color:#006699; color:#FFFFFF">
                                   <td>Agents</td>
                                   <td>Total Appel emis</td>
                                   <td>Total Appel recus</td>
                                   <td>Date d'engagement</td>
                                   <td>%</td>
                                   <td>Deja Engage</td>
                                   <td>%</td>
                                   <td>Refus Argu</td>
                                   <td>%</td>
                                   <td>Refus Intro</td>
                                   <td>%</td>
                                   <td>Repondeur</td>
                                   <td>%</td>
                                   <td>Relance</td>
                                   <td>%</td>
                                   <td>Vente</td>
                                   <td>%</td>
                                   <td>Temps de com / vente</td>
                                   <td>Temps de com moyen global</td>
                                </tr>  
                                    
                                    
                                     