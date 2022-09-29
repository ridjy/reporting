<!doctype html>
<html lang="en">
<head>
	<meta charset="utf-8" />
	<link rel="icon" type="image/png" href="<?php echo img_url('logo.png'); ?>">
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
	<title>Bienvenue <?php echo $this->session->userdata('user_login'); ?></title>
	<meta content='width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0' name='viewport' />
    <meta name="viewport" content="width=device-width" />
    <!-- Bootstrap core CSS     -->
    <link href="<?php echo css_url('bootstrap.min'); ?>" rel="stylesheet" />

    <!--  Light Bootstrap Table core CSS    -->
    <link href="<?php echo css1_url('light-bootstrap-dashboard.css?v=1.4.0'); ?>" rel="stylesheet"/>

    <!--     Fonts and icons     -->
    <link href="http://maxcdn.bootstrapcdn.com/font-awesome/4.2.0/css/font-awesome.min.css" rel="stylesheet">
    <link href="<?php echo css_url('pe-icon-7-stroke'); ?>" rel="stylesheet" />
</head>


<body>
<div class="wrapper">
    <div class="sidebar" data-color="red" data-image="<?php //echo img_url('LOGO2.gif');?>">
    <!--
        Tip 1: you can change the color of the sidebar using: data-color="blue | azure | green | orange | red | purple"
        Tip 2: you can also add an image using data-image tag
    -->
    	<div class="sidebar-wrapper">
            <div class="logo">
                <img src="<?php echo img_url('logo.png');?>" alt="logo" />
                <!--<a href="#" class="simple-text">
                    Allo-Taxi
                </a>-->
            </div>
            <ul class="nav">
                <li class="active">
                    <a href="dashboard.html">
                        <i class="pe-7s-graph"></i>
                        <p>Statistiques</p>
                    </a>
                </li>
                <li >
                    <a href="dashboard.html">
                        <i class="pe-7s-volume1"></i>
                        <p>Enregistrements</p>
                    </a>
                </li>
                <li >
                    <a href="dashboard.html">
                        <i class="pe-7s-id"></i>
                        <p>Utilisateurs</p>
                    </a>
                </li>
            </ul>
    	</div>
    </div>

    <div class="main-panel">
        <nav class="navbar navbar-default navbar-fixed">
            <div class="container-fluid">
                <div class="navbar-header">
                    <button type="button" class="navbar-toggle" data-toggle="collapse">
                        <span class="sr-only">Toggle navigation</span>
                        <span class="icon-bar"></span>
                        <span class="icon-bar"></span>
                        <span class="icon-bar"></span>
                    </button>
                    <a class="navbar-brand" href="#">ROA : Christian RAMAMPY&nbsp;&nbsp;<i class="pe-7s-phone"></i>TELEPHONE : 0170613244&nbsp;&nbsp;<i class="pe-7s-mail"></i>EMAIL : r.christian@bpooceanindien</a>
                </div>
                <div class="collapse navbar-collapse">
                    <ul class="nav navbar-nav navbar-left">
                        <li>
                            <a href="#" class="dropdown-toggle" data-toggle="dropdown">
                               <!-- <i class="fa fa-dashboard"></i> -->
                            </a>
                        </li>
                    </ul>

                    <ul class="nav navbar-nav navbar-right">
                        <li>
                           <a href="">
                               Déconnexion
                            </a>
                        </li>
                    </ul>
                </div>
            </div>
        </nav>


        <div class="content">
            <div class="container-fluid">
                <div class="row">
                        <!--<p>Tout le contenu ici</p>-->
                        

                </div>
            </div>
        </div>


        <footer class="footer">
            <div class="container-fluid">
                <nav class="pull-left">
                    <ul>
                        <li><a href="#">Statistiques</a></li>
                        <li><a href="#">Enregistrements</a></li>
                        <li><a href="#">Utilisateurs</a></li>
                    </ul>
                </nav>
                <p class="copyright pull-right">
                    &copy; <script>document.write(new Date().getFullYear())</script> <a href="#">BPOOI</a>, Tous droits réservés.
                </p>
            </div>
        </footer>

    </div>
</div>


</body>

    <!--   Core JS Files   -->
    <script src="<?php echo js_url('jquery.3.2.1.min'); ?>" type="text/javascript"></script>
	<script src="<?php echo js_url('bootstrap.min'); ?>" type="text/javascript"></script>

	<!--  Charts Plugin -->
	<script src="<?php echo js_url('chartist.min'); ?>"></script>

    <!--  Notifications Plugin    -->
    <script src="<?php echo js_url('bootstrap-notify'); ?>"></script>

    <!-- Light Bootstrap Table Core javascript and methods for Demo purpose -->
	<script src="<?php echo js_url('light-bootstrap-dashboard'); ?>"></script>

	<!-- Light Bootstrap Table DEMO methods, don't include it in your project! -->
	<script src="<?php echo js_url('demo'); ?>"></script>


</html>
