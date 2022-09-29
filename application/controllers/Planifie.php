<?php
class Planifie extends CI_Controller
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

//fonctions utiles
	private function get_year_du_mois_prec($m)
	{
		//$m=date('n');
		//$mois est le num de mois
		$y = date('Y');
		if ($m==1) {
			return $y-1;
		} else { 
			return $y;
		}
	}

	private function get_mois_prec($m)
	{
		if ($m==1) {
			return 12;
		} else {
			return $m-1;
		}
	}

public function newfile()
{
	$data['titre']='Reporting Zeop';
	$data['date']=date('Y-m-d');

	error_reporting(E_ALL);
	ini_set('display_errors', TRUE); 
	ini_set('display_startup_errors', TRUE);

	/*if (PHP_SAPI == 'cli')
		die('Utiliser un navigateur svp');*/

	/** Include PHPExcel */
	//require_once dirname(__FILE__) . '/../Classes/PHPExcel.php';


	// Create new PHPExcel object
	$objPHPExcel = new PHPExcel();

	// Set document properties
	$objPHPExcel->getProperties()->setCreator("Informatique BPO")
								 ->setLastModifiedBy("Informatique")
								 ->setTitle("Reporting Zeop")
								 ->setSubject("Office 2007 XLSX")
								 ->setDescription("Reporting hebdo Zeop")
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

	$sheet->mergeCells('C1:D1');
	
	//$sheet->getColumnDimension('A2')->setWidth(50);
	//$sheet->getColumnDimension('A1')->setHeight(150);

	$affichedate=date('d-m-Y',strtotime($data['date']));

	$objPHPExcel->setActiveSheetIndex(0)
	            ->setCellValue('A1', $affichedate);

	$objDrawing = new PHPExcel_Worksheet_Drawing();
	$objDrawing->setName('logo OceanCall');
	$objDrawing->setDescription('logo OceanCall');
	$objDrawing->setPath('assets/images/Oceancall Group.png');
	$objDrawing->setHeight(50);
	//$objDrawing->setWidth(50);
	$objDrawing->setCoordinates('B1');
	$objDrawing->setOffsetX(-10);
	$objDrawing->setWorksheet($sheet); 

	$objDrawing = new PHPExcel_Worksheet_Drawing();
	$objDrawing->setName('logo OceanCall');
	$objDrawing->setDescription('logo OceanCall');
	$objDrawing->setPath('assets/images/ZEOP-LOGO.jpg');
	$objDrawing->setHeight(65);
	$objDrawing->setWidth(175);
	$objDrawing->setCoordinates('C1');
	$objDrawing->setOffsetX(-10);
	$objDrawing->setWorksheet($sheet);


	$sheet->getRowDimension('1')->setRowHeight(85);                

	$objPHPExcel->setActiveSheetIndex(0)
	            ->setCellValue('A2', '')
	            ->setCellValue('B2', '')
	            ->setCellValue('C2', 'JOUR')
	            ->setCellValue('D2', 'CUMMUL DEBUT DU MOIS');

	/*style des cellules*/
	$sheet->getStyle('C2')->applyFromArray($stylecolor);
	$sheet->getStyle('D2')->applyFromArray($stylecolor);
	$sheet->getStyle('A3:A6')->applyFromArray($style1);
	$sheet->getStyle('B3:B6')->applyFromArray($style1);

	date_default_timezone_set('Africa/Addis_Ababa');

		//integrer le choix de base de données
		/*on détecte les dates d'archivages par rapport à la date du jour*/
			/*si nous n'avons pas encore passé le 1er lundi or qu'on est dans le mois encours*/
			$mois_du_debut=date('n',strtotime($data['date']));
			$year_du_debut=date('Y',strtotime($data['date']));
			$mois_de_fin=date('n',strtotime($data['date']));
			$year_de_fin=date('Y',strtotime($data['date']));
			$mois_en_cours=date('n');
			$year_en_cours=date('Y'); 
			$mois_prec_debut=$this->get_mois_prec($mois_du_debut);
			$year_mois_prec_debut=$this->get_year_du_mois_prec($mois_du_debut);
			$mois_prec=$this->get_mois_prec($mois_en_cours);
			$year_mois_prec=$this->get_year_du_mois_prec($mois_en_cours);
			
			$time = mktime(0, 0, 0, $mois_du_debut, 1, $year_du_debut);//1er du mois du début
			$time_mois_prec= mktime(0, 0, 0, $mois_prec_debut, 1, $year_mois_prec);//1er du mois prec du début
			$str_firstmonday = strtotime('monday', $time);//moisdudebut
			$premier_mois_prec= strtotime('monday', $time_mois_prec);//1er lundi du mois prec

			$timestamp=strtotime($data['date']);
			$premier_mois_encours=mktime(0, 0, 0, date('n'), 1, date('Y'));
			$str_firstmonday_mois=strtotime('monday', $premier_mois_encours);
			if(time()<$str_firstmonday_mois)
			{
				$archivage1=strtotime('monday', $premier_mois_prec);//archivage du mois dernier
				$mois2=$this->get_mois_prec($mois_prec);
				$year2=$this->get_year_du_mois_prec($mois_prec);		
				$premier_mois_prec=mktime(0, 0, 0, $mois2, 1, $year2); 
				$archivage2=strtotime('monday', $premier_mois_prec);
			}
			else
			{
				$archivage1=strtotime('monday', $premier_mois_encours);//archivage le plus récent
				$premier_mois_prec=mktime(0, 0, 0, $mois_prec, 1, $year_mois_prec); 
				$archivage2=strtotime('monday', $premier_mois_prec);//archivage du mis dernier	
			}
			/**/
		//

		//debutannee today
		$debutmois = mktime(0, 0, 0, date('n',$timestamp), 1, date('Y',$timestamp));
		$debutmois = date('Y-m-d',$debutmois);
		$premierlundi=date('Y-m-d',$str_firstmonday_mois);

		//timestamp est la date choisie du get
		if($timestamp>$archivage1)
		{
			$base='statistic';
		}else if ($timestamp>$archivage2)
		{
			$base='statistichisto';
		}else
		{
			$base='statictic2017';
		}

		//echo $debutannee;

		

	$debutmois = mktime(0, 0, 0, date('m'), 1, date('Y'));
	$data['report']=$this->report_model->reportzeopsortant($data['date'],$base);
	$data['cumul']=$this->report_model->cummulanneezeopsortant($premierlundi,$data['date'],'statistic');
	$data['cumul2']=$this->report_model->cummulanneezeopsortant($debutmois,$premierlundi,'statistichisto');

		$data['cumultotal']['TOTAL_EMIS']=$data['cumul']['TOTAL_EMIS']+$data['cumul2']['TOTAL_EMIS'];
		$data['cumultotal']['TOTAL_DECROCHE']=$data['cumul']['TOTAL_DECROCHE']+$data['cumul2']['TOTAL_DECROCHE'];
		$data['cumultotal']['REFUS_ARGUMENTE']=$data['cumul']['REFUS_ARGUMENTE']+$data['cumul2']['REFUS_ARGUMENTE'];
		$data['cumultotal']['OK_VENTE']=$data['cumul']['OK_VENTE']+$data['cumul2']['OK_VENTE'];

	$sheet->mergeCells('A3:A6');
	$objPHPExcel->setActiveSheetIndex(0)
			            ->setCellValue('A3', 'MADA');
			
	$objPHPExcel->setActiveSheetIndex(0)
	            ->setCellValue('B3', 'Total Appel émis')
	            ->setCellValue('B4', 'Total décroché')
	            ->setCellValue('B5', 'Refus argumentés')
	            ->setCellValue('B6', 'Ventes');

	/*$sheet->getStyle('A'.$i)->applyFromArray($stylecolor);            
		
			            $i++;*/

	$objPHPExcel->setActiveSheetIndex(0)
		            ->setCellValue('C3', $data['report']['TOTAL_EMIS'])
		            ->setCellValue('C4', $data['report']['TOTAL_DECROCHE'])
		            ->setCellValue('C5', $data['report']['REFUS_ARGUMENTE'])
		            ->setCellValue('C6', $data['report']['OK_VENTE']);

    $objPHPExcel->setActiveSheetIndex(0)
    ->setCellValue('D3', $data['cumultotal']['TOTAL_EMIS'])
    ->setCellValue('D4', $data['cumultotal']['TOTAL_DECROCHE'])
    ->setCellValue('D5', $data['cumultotal']['REFUS_ARGUMENTE'])
    ->setCellValue('D6', $data['cumultotal']['OK_VENTE']);

	$sheet->getColumnDimension('A')->setWidth(30);
	$sheet->getColumnDimension('B')->setWidth(30);
	$sheet->getColumnDimension('C')->setWidth(18);
	$sheet->getColumnDimension('D')->setWidth(30);
		
		// Rename worksheet
	$objPHPExcel->getActiveSheet()->setTitle('Reporting');

	$sheet->getStyle('C3:C6')->applyFromArray(array('alignment' => array(
    'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
       )
    ));  
	$sheet->getStyle('D3:D6')->applyFromArray(array('alignment' => array(
    'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
       )
    ));


	// Set active sheet index to the first sheet, so Excel opens this as the first sheet
	$objPHPExcel->setActiveSheetIndex(0);


	// Redirect output to a client’s web browser (Excel2007)
	/*header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	header('Content-Disposition: attachment;filename="ReportZeop.xlsx"');
	header('Cache-Control: max-age=0');
	// If you're serving to IE 9, then the following may be needed
	header('Cache-Control: max-age=1');

	// If you're serving to IE over SSL, then the following may be needed
	header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
	header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
	header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
	header ('Pragma: public'); // HTTP/1.0*/

	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	$objWriter->save('pj/'.date('d-m-Y').'.xlsx');


	//$this->testmail();
	$mail = 'r.rakotomalala@bpooceanindien.com'; // Déclaration de l'adresse de destination.
	//$copie = '';
	$copie = 'j.deau@oceancallcentre.com, cedric@oceancallcentre.com, n.randrianarijaona@bpooceanindien.com, raphael@oceancallcentre.com, t.baho@oceancallcentre.com, p.caroux@oceancallcentre.com, informatique@bpooceanindien.com';
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
	$message_txt = "Bonjour, Ci-joint le reporting de ce jour pour Madagascar ".date('d-m-Y H:i');
	$message_html = "<html><head></head><body><b>Bonjour</b>, Ci-joint le reporting de ce jour pour Madagascar <i>".date('d-m-Y H:i')."</i>.</body></html>";
	//==========
	  
	//=====Lecture et mise en forme de la pièce jointe.
	$fichier   = fopen('pj/'.date('d-m-Y').'.xlsx', "r");
	$attachement = fread($fichier, filesize('pj/'.date('d-m-Y').'.xlsx'));
	$attachement = chunk_split(base64_encode($attachement));
	fclose($fichier);
	//==========
	  
	//=====Création de la boundary.
	$boundary = "-----=".md5(rand());
	$boundary_alt = "-----=".md5(rand());
	//==========
	  
	//=====Définition du sujet.
	$sujet = "Reporting ZEOP ".date('d-m-Y H:i');
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
	$message.= "Content-Type: application/vnd.ms-excel; name=\"".date('d-m-Y').".xlsx\"".$passage_ligne;
	$message.= "Content-Transfer-Encoding: base64".$passage_ligne;
	$message.= "Content-Disposition: attachment; filename=\"".date('d-m-Y').".xlsx\"".$passage_ligne;
	$message.= $passage_ligne.$attachement.$passage_ligne.$passage_ligne;
	$message.= $passage_ligne."--".$boundary."--".$passage_ligne;
	//==========
	//=====Envoi de l'e-mail.
	mail($mail,$sujet,$message,$header); // Envoi du message


	exit;
}

public function testmail()
{
		/*$file_name = "pj.xls";
		$path = $_SERVER['DOCUMENT_ROOT']."ReportZeop/pj/";*/
		
		$mail = 'r.rakotomalala@bpooceanindien.com'; // Déclaration de l'adresse de destination.
		//$copie = 'lalaina@bpooceanindien.com, informatique@bpooceanindien.com';
		$copie = 'lalaina@bpooceanindien.com';
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

}