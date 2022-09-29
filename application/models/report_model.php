<?php if ( ! defined('BASEPATH')) exit('No direct script access
allowed');

class Report_model extends CI_Model
{
	protected $table = 'campaign';
	public function __construct()
	{
		// Obligatoire
		parent::__construct();
		//$this->load->database('base1',TRUE);
		//$DB2=$this->load->database('default',TRUE);
		$host = "10.0.3.5";
		$user = "root";
		$pass = "";
		$db = "ireflet";
		$con = mysql_connect($host,$user,$pass);
		$select_bd = mysql_select_db($db);
	}
	
	public function get_campagneID($c)
	{
		$select_id_campagne = 'SELECT c.Id AS "id" FROM ireflet.campaign AS C WHERE C.`Name`="'.$c.'"';
		$id_campagne = mysql_query($select_id_campagne);
		$idcampagne = mysql_fetch_array($id_campagne);
		return $idcampagne['id'];
	}


	public function get_ok_location()
	{
		
		$requete='SELECT clients.DateAppel, clients.HeureAppel, clients.NomAgent, clients.Civ, clients.NOM, clients.PRENOM, clients.ADRESSE, clients.CODE_POSTAL, clients.VILLE, clients.TELEPHONE, clients.MOBILE, clients.MAIL, vqualification.Label, clients.projet, clients.bien, clients.superficie, clients.adresse_bien, clients.adresse_bien_bis, clients.agence_recontrez, clients.JJ_rdv, clients.MM_rdv, clients.AA_rdv, clients.proprietaire_location, clients.HH_rdv, clients.MN_rdv, clients.JJ_rappel, clients.MM_rappel, clients.AA_rappel, clients.HH_rappel, clients.MN_rappel, clients.commentaire
			FROM (clients INNER JOIN sessiondigest ON clients.idSession = sessiondigest.SessionId) INNER JOIN vqualification ON sessiondigest.QualificationId = vqualification.Id
			WHERE (((clients.DateAppel)=Format(Date(),"dd/mm/yyyy")) AND ((vqualification.Label) Like "*ok*") AND ((clients.TRANSACTION_ou_LOCATION)="location"))';
		$exec_requete = mysql_query($requete);
		return $exec_requete;
		//echo $requete;
	}

	public function get_ok_vente()
	{
		$requete='SELECT clients.DateAppel, clients.HeureAppel, clients.NomAgent, clients.Civ, clients.NOM, clients.PRENOM, clients.ADRESSE, clients.CODE_POSTAL, clients.VILLE, clients.TELEPHONE, clients.MOBILE, clients.MAIL, vqualification.Label, clients.projet, clients.bien, clients.superficie, clients.adresse_bien, clients.adresse_bien_bis, clients.agence_recontrez, clients.JJ_rdv, clients.MM_rdv, clients.AA_rdv, clients.proprietaire_location, clients.HH_rdv, clients.MN_rdv, clients.JJ_rappel, clients.MM_rappel, clients.AA_rappel, clients.HH_rappel, clients.MN_rappel, clients.commentaire
			FROM (clients INNER JOIN sessiondigest ON clients.idSession = sessiondigest.SessionId) INNER JOIN vqualification ON sessiondigest.QualificationId = vqualification.Id
			WHERE (((clients.DateAppel)=Format(Date(),"dd/mm/yyyy")) AND ((vqualification.Label) Like "*ok*") AND ((clients.TRANSACTION_ou_LOCATION)="TRANSACTION"))';
		$exec_requete = mysql_query($requete);
		return $exec_requete;
		//echo $requete;
	}

} 