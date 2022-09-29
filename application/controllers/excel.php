<?php
class Excel extends CI_Controller
{	
public function __construct()
{
	// Obligatoire
	parent::__construct();
	// Maintenant, ce code sera exécuté chaque fois que ce contrôleur sera appelé.
	$this->load->helper('url');
	$this->load->database();
	$this->load->model('report_model');
	//define('FPDF_FONTPATH',$this->config->item('fonts_path'));
	$this->load->library('PHPExcel');
	$data = array();
	$agent= array();
}

public function pourcentage($total,$valeur)
	{
		$retour=($valeur*100)/$total;
		return round($retour,2).'%';
	}

public function file_ok_location()
{
	$data['titre']='OK Location';
	$data['date']=$_GET['d'];//à revoir
	//$exec_query=$this->report_model->get_ok_location($_GET['d']);

	error_reporting(E_ALL);
	ini_set('display_errors', TRUE); 
	ini_set('display_startup_errors', TRUE);
	
	if (PHP_SAPI == 'cli')
		die('Utiliser un navigateur svp');

	/** Include PHPExcel */
	//require_once dirname(__FILE__) . '/../Classes/PHPExcel.php';


	// Create new PHPExcel object
	$objPHPExcel = new PHPExcel();

	// Set document properties
	$objPHPExcel->getProperties()->setCreator("Informatique OCC BPO")
								 ->setLastModifiedBy("Informatique BPO")
								 ->setTitle("Justel OK location")
								 ->setSubject("Office 2007 XLSX")
								 ->setDescription("Extraction du ".$data['date'])
								 ->setKeywords("office 2007 openxml php")
								 ->setCategory("Result file");

	$style1 = array(
	  'font'  => array(
	  	'bold'  => true,
	  	'name'  => 'Calibri',
	  	'size'  => 11,
	  	'color' => array('rgb' => 'FFFFFFFF'),
	  ),
	  'fill' => array(
	  	'type' => PHPExcel_Style_Fill::FILL_SOLID, 
	  	'color' => array('rgb' => '58ACFA')
	  ),
	  'alignment' => array( 
	  	'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
	  	'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
	  	'wrap' => true // retour à la ligne automatique
	  )
	);

	$stylecolor=array(
	 	'font'  => array(
		  	'bold'  => true,
		  	'name'  => 'Calibri',
		  	'size'  => 11,
		  	'color' => array('rgb' => 'FFFFFFFF'),
	  ),
		'fill' => 
         array( 'type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => 
             array('rgb' => 'F7FE2E') 
         ));


	$sheet=$objPHPExcel->getActiveSheet();

	$sheet->getRowDimension('1')->setRowHeight(35);                

	$objPHPExcel->setActiveSheetIndex(0)
	            ->setCellValue('A1', 'DATE APPEL')
	            ->setCellValue('B1', 'HEURE APPEL')
	            ->setCellValue('C1', 'NOM AGENT')
	            ->setCellValue('D1', 'CIV')
	            ->setCellValue('E1', 'NOM')
	            ->setCellValue('F1', 'PRENOM')
	            ->setCellValue('G1', 'ADRESSE')
	            ->setCellValue('H1', 'CP')
	            ->setCellValue('I1', 'VILLE')
	            ->setCellValue('J1', 'TEL')
	            ->setCellValue('K1', 'MOBILE')
	            ->setCellValue('L1', 'MAIL')
	            ->setCellValue('M1', 'Label')
	            ->setCellValue('N1', 'projet')
	            ->setCellValue('O1', 'bien')
	            ->setCellValue('P1', 'superficie')
	            ->setCellValue('Q1', 'adresse_bien')
	            ->setCellValue('R1', 'adresse_bien_bis')
	            ->setCellValue('S1', 'agence_recontré')
	            ->setCellValue('T1', 'JJ_rdv')
	            ->setCellValue('U1', 'MM_rdv')
	            ->setCellValue('V1', 'AA_rdv')
	            ->setCellValue('W1', 'proprietaire_location')
	            ->setCellValue('X1', 'HH_rdv')
	            ->setCellValue('Y1', 'MN_rdv')
	            ->setCellValue('Z1', 'JJ_rappel')
	            ->setCellValue('AA1', 'MM_rappel')
	            ->setCellValue('AB1', 'AA_rappel')
	            ->setCellValue('AC1', 'HH_rappel')
	            ->setCellValue('AD1', 'MN_rappel')
	            ->setCellValue('AE1', 'commentaire');

	/*style des cellules*/
	$sheet->getStyle('A1:AE1')->applyFromArray($style1);

	$exec_query=$this->report_model->get_ok_location();

	$i=2;
	while ($data = mysql_fetch_array($exec_query))
	{
		# code...
		$objPHPExcel->setActiveSheetIndex(0)
		            ->setCellValue('A'.$i, $data['DateAppel'])
		            ->setCellValue('B'.$i, $data['HeureAppel'])
		            ->setCellValue('C'.$i, $data['NomAgent'])
		            ->setCellValue('D'.$i, $data['Civ'])
		            ->setCellValue('E'.$i, $data['NOM'])
		            ->setCellValue('F'.$i, $data['PRENOM'])
		            ->setCellValue('G'.$i, $data['ADRESSE'])
		            ->setCellValue('H'.$i, $data['CODE_POSTAL'])
		            ->setCellValue('I'.$i, $data['VILLE'])
		            ->setCellValue('J'.$i, $data['TELEPHONE'])
		            ->setCellValue('K'.$i, $data['MOBILE'])
		            ->setCellValue('L'.$i, $data['MAIL'])
		            ->setCellValue('M'.$i, $data['Label'])
		            ->setCellValue('N'.$i, $data['projet'])
		            ->setCellValue('O'.$i, $data['bien'])
		            ->setCellValue('P'.$i, $data['superficie'])
		            ->setCellValue('Q'.$i, $data['adresse_bien'])
		            ->setCellValue('R'.$i, $data['adresse_bien_bis'])
		            ->setCellValue('S'.$i, $data['agence_recontrez'])
		            ->setCellValue('T'.$i, $data['JJ_rdv'])
		            ->setCellValue('U'.$i, $data['MM_rdv'])
		            ->setCellValue('V'.$i, $data['AA_rdv'])
		            ->setCellValue('W'.$i, $data['proprietaire_location'])
		            ->setCellValue('X'.$i, $data['HH_rdv'])
		            ->setCellValue('Y'.$i, $data['MN_rdv'])
		            ->setCellValue('Z'.$i, $data['JJ_rappel'])
		            ->setCellValue('AA'.$i, $data['MM_rappel'])
		            ->setCellValue('AB'.$i, $data['AA_rappel'])
		            ->setCellValue('AC'.$i, $data['HH_rappel'])
		            ->setCellValue('AD'.$i, $data['MN_rappel'])
		            ->setCellValue('AE'.$i, $data['commentaire'])		            ;
		$i++;            			
	}	

	$sheet->getColumnDimension('A')->setWidth(17);
	$sheet->getColumnDimension('B')->setWidth(26);
	$sheet->getColumnDimension('C')->setWidth(10);
	$sheet->getColumnDimension('D')->setWidth(8);
	$sheet->getColumnDimension('E')->setWidth(25);
	$sheet->getColumnDimension('F')->setWidth(28);
	$sheet->getColumnDimension('G')->setWidth(43);
	$sheet->getColumnDimension('H')->setWidth(12);
	$sheet->getColumnDimension('I')->setWidth(18);
	$sheet->getColumnDimension('J')->setWidth(15);
	$sheet->getColumnDimension('K')->setWidth(18);
	$sheet->getColumnDimension('L')->setWidth(17);
	$sheet->getColumnDimension('M')->setWidth(18);
	$sheet->getColumnDimension('N')->setWidth(25);
	$sheet->getColumnDimension('O')->setWidth(23);
	$sheet->getColumnDimension('P')->setWidth(31);
	$sheet->getColumnDimension('Q')->setWidth(50);
	$sheet->getColumnDimension('R')->setWidth(10);
	$sheet->getColumnDimension('S')->setWidth(10);
	$sheet->getColumnDimension('T')->setWidth(10);
	$sheet->getColumnDimension('U')->setWidth(10);
	$sheet->getColumnDimension('V')->setWidth(10);
	$sheet->getColumnDimension('W')->setWidth(10);
	$sheet->getColumnDimension('X')->setWidth(10);
	$sheet->getColumnDimension('Y')->setWidth(10);
	$sheet->getColumnDimension('Z')->setWidth(10);
	$sheet->getColumnDimension('AA')->setWidth(10);
	$sheet->getColumnDimension('AB')->setWidth(10);
	$sheet->getColumnDimension('AC')->setWidth(10);
	$sheet->getColumnDimension('AD')->setWidth(10);
	$sheet->getColumnDimension('AE')->setWidth(10);
		
		// Rename worksheet
	$objPHPExcel->getActiveSheet()->setTitle("OK location Justel");

	/*$sheet->getStyle('C3:C6')->applyFromArray(array('alignment' => array(
    'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
       )
    ));*/

	// Set active sheet index to the first sheet, so Excel opens this as the first sheet
	$objPHPExcel->setActiveSheetIndex(0);


	// Redirect output to a client’s web browser (Excel2007)
	header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	header('Content-Disposition: attachment;filename="Justel_Ok_location_'.$daterdv.'.xls"');
	header('Cache-Control: max-age=0');
	// If you're serving to IE 9, then the following may be needed
	header('Cache-Control: max-age=1');

	// If you're serving to IE over SSL, then the following may be needed
	header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
	header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
	header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
	header ('Pragma: public'); // HTTP/1.0

	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	$objWriter->save('pj/Justel_Ok_location_'.$daterdv.'.xls');

	//$objWriter->save('php://output');

	//$this->testmail($base,$daterdv);
	exit;
}

public function file_ok_vente()
{
	$data['titre']='OK Vente';
	$data['date']=$_GET['d'];//à revoir
	//$exec_query=$this->report_model->get_ok_location($_GET['d']);

	error_reporting(E_ALL);
	ini_set('display_errors', TRUE); 
	ini_set('display_startup_errors', TRUE);
	
	if (PHP_SAPI == 'cli')
		die('Utiliser un navigateur svp');

	/** Include PHPExcel */
	//require_once dirname(__FILE__) . '/../Classes/PHPExcel.php';


	// Create new PHPExcel object
	$objPHPExcel = new PHPExcel();

	// Set document properties
	$objPHPExcel->getProperties()->setCreator("Informatique OCC BPO")
								 ->setLastModifiedBy("Informatique BPO")
								 ->setTitle("Justel OK vente")
								 ->setSubject("Office 2007 XLSX")
								 ->setDescription("Extraction du ".$data['date'])
								 ->setKeywords("office 2007 openxml php")
								 ->setCategory("Result file");

	$style1 = array(
	  'font'  => array(
	  	'bold'  => true,
	  	'name'  => 'Calibri',
	  	'size'  => 11,
	  	'color' => array('rgb' => 'FFFFFFFF'),
	  ),
	  'fill' => array(
	  	'type' => PHPExcel_Style_Fill::FILL_SOLID, 
	  	'color' => array('rgb' => '58ACFA')
	  ),
	  'alignment' => array( 
	  	'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
	  	'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
	  	'wrap' => true // retour à la ligne automatique
	  )
	);

	$stylecolor=array(
	 	'font'  => array(
		  	'bold'  => true,
		  	'name'  => 'Calibri',
		  	'size'  => 11,
		  	'color' => array('rgb' => 'FFFFFFFF'),
	  ),
		'fill' => 
         array( 'type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => 
             array('rgb' => 'F7FE2E') 
         ));


	$sheet=$objPHPExcel->getActiveSheet();

	$sheet->getRowDimension('1')->setRowHeight(35);                

	$objPHPExcel->setActiveSheetIndex(0)
	            ->setCellValue('A1', 'DATE APPEL')
	            ->setCellValue('B1', 'HEURE APPEL')
	            ->setCellValue('C1', 'NOM AGENT')
	            ->setCellValue('D1', 'CIV')
	            ->setCellValue('E1', 'NOM')
	            ->setCellValue('F1', 'PRENOM')
	            ->setCellValue('G1', 'ADRESSE')
	            ->setCellValue('H1', 'CP')
	            ->setCellValue('I1', 'VILLE')
	            ->setCellValue('J1', 'TEL')
	            ->setCellValue('K1', 'MOBILE')
	            ->setCellValue('L1', 'MAIL')
	            ->setCellValue('M1', 'Label')
	            ->setCellValue('N1', 'projet')
	            ->setCellValue('O1', 'bien')
	            ->setCellValue('P1', 'superficie')
	            ->setCellValue('Q1', 'adresse_bien')
	            ->setCellValue('R1', 'adresse_bien_bis')
	            ->setCellValue('S1', 'agence_recontré')
	            ->setCellValue('T1', 'JJ_rdv')
	            ->setCellValue('U1', 'MM_rdv')
	            ->setCellValue('V1', 'AA_rdv')
	            ->setCellValue('W1', 'proprietaire_location')
	            ->setCellValue('X1', 'HH_rdv')
	            ->setCellValue('Y1', 'MN_rdv')
	            ->setCellValue('Z1', 'JJ_rappel')
	            ->setCellValue('AA1', 'MM_rappel')
	            ->setCellValue('AB1', 'AA_rappel')
	            ->setCellValue('AC1', 'HH_rappel')
	            ->setCellValue('AD1', 'MN_rappel')
	            ->setCellValue('AE1', 'commentaire');

	/*style des cellules*/
	$sheet->getStyle('A1:AE1')->applyFromArray($style1);

	$exec_query=$this->report_model->get_ok_vente();

	$i=2;
	while ($data = mysql_fetch_array($exec_query))
	{
		# code...
		$objPHPExcel->setActiveSheetIndex(0)
		            ->setCellValue('A'.$i, $data['DateAppel'])
		            ->setCellValue('B'.$i, $data['HeureAppel'])
		            ->setCellValue('C'.$i, $data['NomAgent'])
		            ->setCellValue('D'.$i, $data['Civ'])
		            ->setCellValue('E'.$i, $data['NOM'])
		            ->setCellValue('F'.$i, $data['PRENOM'])
		            ->setCellValue('G'.$i, $data['ADRESSE'])
		            ->setCellValue('H'.$i, $data['CODE_POSTAL'])
		            ->setCellValue('I'.$i, $data['VILLE'])
		            ->setCellValue('J'.$i, $data['TELEPHONE'])
		            ->setCellValue('K'.$i, $data['MOBILE'])
		            ->setCellValue('L'.$i, $data['MAIL'])
		            ->setCellValue('M'.$i, $data['Label'])
		            ->setCellValue('N'.$i, $data['projet'])
		            ->setCellValue('O'.$i, $data['bien'])
		            ->setCellValue('P'.$i, $data['superficie'])
		            ->setCellValue('Q'.$i, $data['adresse_bien'])
		            ->setCellValue('R'.$i, $data['adresse_bien_bis'])
		            ->setCellValue('S'.$i, $data['agence_recontrez'])
		            ->setCellValue('T'.$i, $data['JJ_rdv'])
		            ->setCellValue('U'.$i, $data['MM_rdv'])
		            ->setCellValue('V'.$i, $data['AA_rdv'])
		            ->setCellValue('W'.$i, $data['proprietaire_location'])
		            ->setCellValue('X'.$i, $data['HH_rdv'])
		            ->setCellValue('Y'.$i, $data['MN_rdv'])
		            ->setCellValue('Z'.$i, $data['JJ_rappel'])
		            ->setCellValue('AA'.$i, $data['MM_rappel'])
		            ->setCellValue('AB'.$i, $data['AA_rappel'])
		            ->setCellValue('AC'.$i, $data['HH_rappel'])
		            ->setCellValue('AD'.$i, $data['MN_rappel'])
		            ->setCellValue('AE'.$i, $data['commentaire'])		            ;
		$i++;            			
	}	

	$sheet->getColumnDimension('A')->setWidth(17);
	$sheet->getColumnDimension('B')->setWidth(26);
	$sheet->getColumnDimension('C')->setWidth(10);
	$sheet->getColumnDimension('D')->setWidth(8);
	$sheet->getColumnDimension('E')->setWidth(25);
	$sheet->getColumnDimension('F')->setWidth(28);
	$sheet->getColumnDimension('G')->setWidth(43);
	$sheet->getColumnDimension('H')->setWidth(12);
	$sheet->getColumnDimension('I')->setWidth(18);
	$sheet->getColumnDimension('J')->setWidth(15);
	$sheet->getColumnDimension('K')->setWidth(18);
	$sheet->getColumnDimension('L')->setWidth(17);
	$sheet->getColumnDimension('M')->setWidth(18);
	$sheet->getColumnDimension('N')->setWidth(25);
	$sheet->getColumnDimension('O')->setWidth(23);
	$sheet->getColumnDimension('P')->setWidth(31);
	$sheet->getColumnDimension('Q')->setWidth(50);
	$sheet->getColumnDimension('R')->setWidth(10);
	$sheet->getColumnDimension('S')->setWidth(10);
	$sheet->getColumnDimension('T')->setWidth(10);
	$sheet->getColumnDimension('U')->setWidth(10);
	$sheet->getColumnDimension('V')->setWidth(10);
	$sheet->getColumnDimension('W')->setWidth(10);
	$sheet->getColumnDimension('X')->setWidth(10);
	$sheet->getColumnDimension('Y')->setWidth(10);
	$sheet->getColumnDimension('Z')->setWidth(10);
	$sheet->getColumnDimension('AA')->setWidth(10);
	$sheet->getColumnDimension('AB')->setWidth(10);
	$sheet->getColumnDimension('AC')->setWidth(10);
	$sheet->getColumnDimension('AD')->setWidth(10);
	$sheet->getColumnDimension('AE')->setWidth(10);
		
		// Rename worksheet
	$objPHPExcel->getActiveSheet()->setTitle("OK vente Justel");

	/*$sheet->getStyle('C3:C6')->applyFromArray(array('alignment' => array(
    'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
       )
    ));*/

	// Set active sheet index to the first sheet, so Excel opens this as the first sheet
	$objPHPExcel->setActiveSheetIndex(0);


	// Redirect output to a client’s web browser (Excel2007)
	header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	header('Content-Disposition: attachment;filename="Justel_Ok_vente_'.$daterdv.'.xls"');
	header('Cache-Control: max-age=0');
	// If you're serving to IE 9, then the following may be needed
	header('Cache-Control: max-age=1');

	// If you're serving to IE over SSL, then the following may be needed
	header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
	header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
	header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
	header ('Pragma: public'); // HTTP/1.0

	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	$objWriter->save('pj/Justel_Ok_vente_'.$daterdv.'.xls');

	//$objWriter->save('php://output');

	//$this->testmail($base,$daterdv);
	exit;
}


public function testmail()
{
		/*$file_name = "pj.xls";
		$path = $_SERVER['DOCUMENT_ROOT']."ReportZeop/pj/";*/
		
		$mail = 'r.rakotomalala@bpooceanindien.com'; // Déclaration de l'adresse de destination.
		//$copie = 'lalaina@bpooceanindien.com, informatique@bpooceanindien.com';
		$copie = '';
		$copie_cachee = '';
		if (!preg_match("#^[a-z0-9._-]+@(hotmail|live|msn).[a-z]{2,4}$#", $mail)) // On filtre les serveurs qui présentent des bogues.
		{
		    $passage_ligne = "\r\n";
		}
		else
		{
		    $passage_ligne = "\n";
		}
		//=====Déclaration des messages au format texte et au format HTML.
		$message_txt = "Bonjour, Ci-joint le reporting de ce jour ".date('d-m-Y H:i');
		$message_html = "<html><head></head><body><b>Bonjour</b>, Ci-joint le reporting de ce jour <i>".date('d-m-Y H:i')."</i>.</body></html>";
		//==========
		  
		//=====Lecture et mise en forme de la pièce jointe.
		$fichier   = fopen("pj/ReportZeop.xlsx", "r");
		$attachement = fread($fichier, filesize("pj/ReportZeop.xlsx"));
		$attachement = chunk_split(base64_encode($attachement));
		fclose($fichier);
		//==========
		  
		//=====Création de la boundary.
		$boundary = "-----=".md5(rand());
		$boundary_alt = "-----=".md5(rand());
		//==========
		  
		//=====Définition du sujet.
		$sujet = "Reporting ZEOP ";
		//=========
		  
		//=====Création du header de l'e-mail.
		$header = "From: \"Informatique OCC\"<informatiqueocc@mail.com>".$passage_ligne;
		$header.= "Reply-to: \"Informatique OCC\" <informatique@bpooceanindien.com>".$passage_ligne;
		$header.= 'Cc: '.$copie.$passage_ligne; // Copie Cc
		$header.= 'Bcc: '.$copie_cachee.$passage_ligne; // Copie cachée Bcc 
		$header.= "MIME-Version: 1.0".$passage_ligne;
		$header.= "Content-Type: multipart/mixed;".$passage_ligne." boundary=\"$boundary\"".$passage_ligne;
		//==========
		  
		//=====Création du message.
		$message = $passage_ligne."--".$boundary.$passage_ligne;
		$message.= "Content-Type: multipart/alternative;".$passage_ligne." boundary=\"$boundary_alt\"".$passage_ligne;
		$message.= $passage_ligne."--".$boundary_alt.$passage_ligne;
		//=====Ajout du message au format texte.
		$message.= "Content-Type: text/plain; charset=\"ISO-8859-1\"".$passage_ligne;
		$message.= "Content-Transfer-Encoding: 8bit".$passage_ligne;
		$message.= $passage_ligne.$message_txt.$passage_ligne;
		//==========
		  
		$message.= $passage_ligne."--".$boundary_alt.$passage_ligne;
		  
		//=====Ajout du message au format HTML.
		$message.= "Content-Type: text/html; charset=\"ISO-8859-1\"".$passage_ligne;
		$message.= "Content-Transfer-Encoding: 8bit".$passage_ligne;
		$message.= $passage_ligne.$message_html.$passage_ligne;
		//==========
		  
		//=====On ferme la boundary alternative.
		$message.= $passage_ligne."--".$boundary_alt."--".$passage_ligne;
		//==========
		  
		$message.= $passage_ligne."--".$boundary.$passage_ligne;
		  
		//=====Ajout de la pièce jointe.
		$message.= "Content-Type: application/vnd.ms-excel; name=\"ReportZeop.xls\"".$passage_ligne;
		$message.= "Content-Transfer-Encoding: base64".$passage_ligne;
		$message.= "Content-Disposition: attachment; filename=\"ReportZeop.xls\"".$passage_ligne;
		$message.= $passage_ligne.$attachement.$passage_ligne.$passage_ligne;
		$message.= $passage_ligne."--".$boundary."--".$passage_ligne;
		//==========
		//=====Envoi de l'e-mail.
		if (mail($mail,$sujet,$message,$header)) // Envoi du message
		{
		    echo 'Votre message a bien été envoyé ';
		}
		else // Non envoyé
		{
		    echo "Votre message n'a pas pu être envoyé";
		}				  
		//==========	
}

public function captureecran()
{
	$im = imagegrabscreen();
	imagepng($im, "myscreenshot.png");
	imagedestroy($im);
}

}
