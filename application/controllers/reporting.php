<?php
class Reporting extends CI_Controller
{
	public function __construct()
	{
		// Obligatoire
		parent::__construct();
		// Maintenant, ce code sera exécuté chaque fois que ce contrôleur sera appelé.
		$this->load->helper('url');
		$this->load->library('session');
		$this->load->database();
		$this->load->model('report_model');
		$this->load->library('form_validation');
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

	public function pourcentage($total,$valeur)
	{
		$retour=($valeur*100)/$total;
		return round($retour,2).'%';
	}

	public function index()
	{
		$data['titre']='Accueil';
		$this->load->view("report/index",$data);	
	}

	public function report()
	{
		$data['titre']='Reporting Zeop';
		$data['date']=$_GET['d'];
		$this->load->view("report/reporting",$data);
		$exec_query=$this->report_model->get_stat_par_agent($_GET['d']);
		$data['TOTAL_GLOBAL']=0;
		$data['DATE_ENGAGEMENT']=0;
		$data['DEJA_ENGAGE']=0;
		$data['REFUS_ARGUMENTE']=0;
		$data['REFUS_INTRO']=0;
		$data['REPONDEUR']=0;
		$data['RELANCE']=0;
		$data['OK_VENTE']=0;
		while($result = mysql_fetch_array($exec_query))
		{
			//$result['sortant']=$this->report_model->get_total_sortant('27/11/2017',$result['agent']);
			$data['TOTAL_GLOBAL']+=$result['TOTAL_GLOBAL'];
			$data['DATE_ENGAGEMENT']+=$result['DATE_ENGAGEMENT'];
			$data['DEJA_ENGAGE']+=$result['DEJA_ENGAGE'];
			$data['REFUS_ARGUMENTE']+=$result['REF'];
			$data['REFUS_INTRO']+=$result['REFUS_INTRO'];
			$data['REPONDEUR']+=$result['REPONDEUR'];
			$data['RELANCE']+=$result['RELANCE'];
			$data['OK_VENTE']+=$result['OK_VENTE'];

			/***pourcentage*******/
			$result['pcent_date']=$this->pourcentage($result['TOTAL_GLOBAL'],$result['DATE_ENGAGEMENT']);
			$result['pcent_engage']=$this->pourcentage($result['TOTAL_GLOBAL'],$result['DEJA_ENGAGE']);
			$result['pcent_arg']=$this->pourcentage($result['TOTAL_GLOBAL'],$result['REF']);
			$result['pcent_intro']=$this->pourcentage($result['TOTAL_GLOBAL'],$result['REFUS_INTRO']);
			$result['pcent_rep']=$this->pourcentage($result['TOTAL_GLOBAL'],$result['REPONDEUR']);
			$result['pcent_rel']=$this->pourcentage($result['TOTAL_GLOBAL'],$result['RELANCE']);
			$result['pcent_ok']=$this->pourcentage($result['TOTAL_GLOBAL'],$result['OK_VENTE']);
			
			/*****fin pourcentage*****/
			$this->load->view('report/contenu_stat',$result);
			//$compteur++;
		}

		$this->load->view("report/footer",$data);
		//var_dump($agent);	
	}

	public function testmail()
	{
		/*$file_name = "pj.xls";
		$path = $_SERVER['DOCUMENT_ROOT']."ReportZeop/pj/";*/
		
		$mail = 'r.rakotomalala@bpooceanindien.com'; // Déclaration de l'adresse de destination.
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
		$message_html = "<html><head></head><body><b>Bonjour</b>, Ci-joint le reporting de ce jour <i>date('d-m-Y H:i')</i>.</body></html>";
		//==========
		  
		//=====Lecture et mise en forme de la pièce jointe.
		$fichier   = fopen("pj.xls", "r");
		$attachement = fread($fichier, filesize("pj.xls"));
		$attachement = chunk_split(base64_encode($attachement));
		fclose($fichier);
		//==========
		  
		//=====Création de la boundary.
		$boundary = "-----=".md5(rand());
		$boundary_alt = "-----=".md5(rand());
		//==========
		  
		//=====Définition du sujet.
		$sujet = "Hey mon ami !";
		//=========
		  
		//=====Création du header de l'e-mail.
		$header = "From: \"Informatique OCC\"<informatiqueocc@mail.fr>".$passage_ligne;
		$header.= "Reply-to: \"Informatique BPO\" <informatique@bpooceanindien.com>".$passage_ligne;
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
		$message.= "Content-Type: application/vnd.ms-excel; name=\"pj.xls\"".$passage_ligne;
		$message.= "Content-Transfer-Encoding: base64".$passage_ligne;
		$message.= "Content-Disposition: attachment; filename=\"pj.xls\"".$passage_ligne;
		$message.= $passage_ligne.$attachement.$passage_ligne.$passage_ligne;
		$message.= $passage_ligne."--".$boundary."--".$passage_ligne;
		//==========
		//=====Envoi de l'e-mail.
		mail($mail,$sujet,$message,$header);
		  
		//==========
	
		
	}

	public function newreport()
	{
		$data['titre']='Reporting Zeop';
		$data['date']=$_GET['d'];
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

			$timestamp=strtotime($_GET['d']);
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

		//cumul pour avant archivage

		$data['report']=$this->report_model->reportzeopsortant($_GET['d'],$base);
		$data['cumul']=$this->report_model->cummulanneezeopsortant($premierlundi,$_GET['d'],'statistic');
		$data['cumul2']=$this->report_model->cummulanneezeopsortant($debutmois,$premierlundi,'statistichisto');

		$data['cumultotal']['TOTAL_EMIS']=$data['cumul']['TOTAL_EMIS']+$data['cumul2']['TOTAL_EMIS'];
		$data['cumultotal']['TOTAL_DECROCHE']=$data['cumul']['TOTAL_DECROCHE']+$data['cumul2']['TOTAL_DECROCHE'];
		$data['cumultotal']['REFUS_ARGUMENTE']=$data['cumul']['REFUS_ARGUMENTE']+$data['cumul2']['REFUS_ARGUMENTE'];
		$data['cumultotal']['OK_VENTE']=$data['cumul']['OK_VENTE']+$data['cumul2']['OK_VENTE'];

		//print_r($data);

		$this->load->view("report/newreport",$data);
	}
	
}