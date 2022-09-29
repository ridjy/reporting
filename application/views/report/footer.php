   <tr class="primary">
                                        <td style="background-color:#006699; color:#FFFFFF">TOTAL</td>
                                        <td><?php echo $TOTAL_GLOBAL ?></td>
                                        <td>0</td>
                                        <td><?php echo $DATE_ENGAGEMENT ?></td>
                                        <td> - </td>
                                        <td><?php echo $DEJA_ENGAGE ?></td>
                                        <td> - </td>
                                        <td><?php echo $REFUS_ARGUMENTE ?></td>
                                        <td> - </td>
                                        <td><?php echo $REFUS_INTRO ?></td>
                                        <td> - </td>
                                        <td><?php echo $REPONDEUR ?></td>
                                        <td> - </td>
                                        <td><?php echo $RELANCE ?></td>
                                        <td> - </td>
                                        <td><?php echo $OK_VENTE ?></td>
                                        <td> - </td>
                                        <td> - </td>
                                        <td> - </td>
                                      </tr>
                                    
                                    <tr>
                                        <?php for($i=0;$i<18;$i++) echo '<td></td>'; ?>
                                    <tr>
                                    <tr>
                                        <?php for($i=0;$i<18;$i++) echo '<td></td>'; ?>
                                    <tr>  
                                    <tr>
                                        <td colspan="6"><p>Date d'engagement : La personne est engage pour plus de 4 mois encore.</p></td>
                                        <?php for($i=0;$i<12;$i++) echo '<td></td>'; ?>
                                    <tr>
                                    <tr>
                                        <td colspan="6"><p><?php //echo $idZeopEntrant.' et '.$idZeopSortant; ?></p></td>
                                        <?php for($i=0;$i<12;$i++) echo '<td></td>'; ?>
                                    <tr>  
                                    <tr>
                                        <td colspan="6"><p>Deja engage : Client Zeop</p></td>
                                        <?php for($i=0;$i<12;$i++) echo '<td></td>'; ?>
                                    <tr> 
                                    <tr>
                                        <td colspan="6"><p><?php //echo $idZeopEntrant.' et '.$idZeopSortant; ?></p></td>
                                        <?php for($i=0;$i<12;$i++) echo '<td></td>'; ?>
                                    <tr>
                                    <tr>
                                        <td colspan="6"><p>Refus Argu : Refus apres avoir presente l'offre</p></td>
                                        <?php for($i=0;$i<12;$i++) echo '<td></td>'; ?>
                                    <tr> 
                                    <tr>
                                        <td colspan="6"><p><?php //echo $idZeopEntrant.' et '.$idZeopSortant; ?></p></td>
                                        <?php for($i=0;$i<12;$i++) echo '<td></td>'; ?>
                                    <tr>
                                    <tr>
                                        <td colspan="6"><p>Refus intro : Le prospect raccroche pendant l'introduction
</p></td>
                                        <?php for($i=0;$i<12;$i++) echo '<td></td>'; ?>
                                    <tr>    
                                    
                            </table>
                         </section>
                </div><!--row-->

                <div class="row">
                    <div class="col-md-4">
                    </div>
                    <div class="col-md-4">
                        <a id="dlink"  style="display:none;"></a>
                        <input type="button" class="btn btn-success btn-md btn-fill" onclick="tableToExcel('report', 'Reporting ZEOP', 'pj.xls')" value="EXPORTER">
                    </div> 
                    <div class="col-md-4">
                        <a href="<?php echo site_url('excel/file?d=2017-12-07'); ?>"><button class="btn btn-success btn-md">Envoie par mail</button></a>
                    </div> 
                                         
                </div>

            </div><!--container fluide-->
     </div> <!--content-->                

</body>

    <!--   Core JS Files   -->
    <script src="<?php echo js_url('jquery.3.2.1.min'); ?>" type="text/javascript"></script>
    <script src="<?php echo js_url('bootstrap.min'); ?>" type="text/javascript"></script>

</html>