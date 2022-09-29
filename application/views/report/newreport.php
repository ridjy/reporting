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
       var tableToExcel = (function () {
           var uri = 'data:application/vnd.ms-excel;base64,'
   , template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:Name>"Statistique_globale.xls"</x:Name><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>'
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
  <br/>
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
                        <table id="report" style="text-align:center" class="table table-bordered">
                                <tr style="height:110px;">
                                    <td ><?php echo $date;//echo date('d-m-Y');?></td>
                                    <td width="200px"><img src="<?php echo img_url('Oceancall Group.png'); ?>" ></td>
                                    <td colspan="2" width="150px"><img src="<?php echo img_url('ZEOP-LOGO.jpg'); ?>" ></td>
                                </tr>  
                                <tr>
                                   <!--ici pr modifier la taille du fichier excel--> 
                                   <td style="width:150px;"></td>
                                   <td style="width:200px;"></td>
                                   <td style="width:100px; background-color:#F7FE2E; font-weight: bold; vertical-align: middle;" >JOUR</td>
                                   <td style="width:200px; background-color:#F7FE2E; font-weight: bold;" >CUMMUL DEBUT DU MOIS</td>
                                
                                </tr>  
                                <tr >
                                  <td style="background-color:#58ACFA; font-weight: bold; vertical-align: middle;" rowspan="4">MADA</td>
                                  <td style="background-color:#58ACFA; font-weight: bold;">Total Appel émis</td>
                                  <td><?php echo $report['TOTAL_EMIS']; ?></td>
                                  <td><?php echo $cumultotal['TOTAL_EMIS']; ?></td>
                                </tr>
                                <tr >
                                  <td style="background-color:#58ACFA; font-weight: bold;">Total décroché</td>
                                  <td><?php echo $report['TOTAL_DECROCHE']; ?></td>
                                  <td><?php echo $cumultotal['TOTAL_DECROCHE']; ?></td>
                                </tr>
                                <tr >
                                  <td style="background-color:#58ACFA; font-weight: bold;">Refus argumentés</td>
                                  <td><?php echo $report['REFUS_ARGUMENTE']; ?></td>
                                  <td><?php echo $cumultotal['REFUS_ARGUMENTE']; ?></td>
                                </tr>
                                <tr >
                                  <td style="background-color:#58ACFA; font-weight: bold;">Ventes</td>
                                  <td><?php echo $report['OK_VENTE']; ?></td>
                                  <td><?php echo $cumultotal['OK_VENTE']; ?></td>
                                </tr> 
                        </table>
                    </section>  
                  </div>
                  

                    <div class="row">
                    <div class="col-md-3">
                        <a href="<?php echo site_url('excel/newfile/?d=').$date;?>"><button class="btn btn-success btn-md btn-fill">Generer le fichier excel sur serveur</button></a>
                    </div>
                    <div class="col-md-3">
                        <a id="dlink"  style="display:none;"></a>
                        <input type="button" class="btn btn-success btn-md btn-fill" onclick="tableToExcel('report', 'Reporting ZEOP', 'pj.xls')" value="EXPORTER">
                    </div> 
                    <div class="col-md-3">
                        <!--<a href="<?php echo site_url('reporting/envoimail'); ?>"><button class="btn btn-success btn-md">Envoie par mail</button></a>-->
                    </div>
                    <div class="col-md-3">
                        <!--<a href="<?php echo site_url('excel/captureecran'); ?>"><button class="btn btn-success btn-md">Capture écran</button></a>-->
                    </div>                    
                </div>

            </div><!--container fluide-->
     </div> <!--content-->                

</body>

    <!--   Core JS Files   -->
    <script src="<?php echo js_url('jquery.3.2.1.min'); ?>" type="text/javascript"></script>
    <script src="<?php echo js_url('bootstrap.min'); ?>" type="text/javascript"></script>

</html>             
                                    
                                    
                                     