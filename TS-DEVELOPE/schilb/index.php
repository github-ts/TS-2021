<!-- SCHILB.2020.V1.5.3-TS_INT -->
<!DOCTYPE html>
<html>

<head>
<meta content="de" http-equiv="Content-Language">
<meta content="text/html; charset=utf-8" http-equiv="Content-Type">
<link rel="icon" type="image/png" href="./.favicon.png">
<link rel="icon" type="image/x-icon" href="./.favicon.ico">
<!-- <link rel="shortcut icon" href="http://x.bplaced.com/x/favicon.ico" type="image/x-icon"> -->
<meta name="author" content="Thomas Schilb">
<meta name="publisher" content="Thomas Schilb">
<meta name="copyright" content="© 2020 SCHILB.">
<meta content="schilb, private, intern" name="keywords">
<meta content="private homepage" name="description">
<meta name="robots" content="all, index, follow"> 
<meta name="googlebot" content="all, index, follow"> 
<meta name="pagerank" content="10"> 
<meta name="msnbot" content="all,index,follow"> 
<meta name="revisit" content="2 Days"> 
<meta name="revisit-after" content="2 Days"> 
<meta name="alexa" content="100">
<meta content="de" http-equiv="Content-Language">
<base target="_blank">
<link href="https://fonts.googleapis.com/css?family=Inter:300" rel="stylesheet">
<head>
<style type="text/css">
font {font-family: 'Inter', sans-serif;}
a {
	color: #959595;
	text-decoration: none;
}
a:visited {
	color: #959595;
}
a:active {
	color: #959595;
}
a:hover {
	color: #000000;
}
.border-1px {
	border-width: 1px;
}
.text-align-right {
	text-align: right;
}
hr.style-one {
    border: 0;
    height: 2px;
    background: #333;
    background-image: linear-gradient(to right, #ccc, #000, #ccc);
}
input textarea {
color: #000000;
background-color: #FFFFFF;
}
.myButton {
	box-shadow:inset 0px 1px 0px 0px #ffffff;
	background:linear-gradient(to bottom, #ffffff 5%, #f6f6f6 100%);
	background-color:#ffffff;
	border-radius:6px;
	border:1px solid #dcdcdc;
	display:inline-block;
	cursor:pointer;
	color:#666666;
	#font-family:Arial;
	font-size:15px;
	#font-weight:bold;
	padding:6px 24px;
	text-decoration:none;
	text-shadow:0px 1px 0px #ffffff;
}
.myButton:hover {
	background:linear-gradient(to bottom, #f6f6f6 5%, #ffffff 100%);
	background-color:#f6f6f6;
}
.myButton:active {
	position:relative;
	top:1px;
}
.line {
	color: #000000;
	height: 2px;
	outline-color:black;
}
</style>
<title>Familie | Schilb</title>
<base target="_self">
</head>

<body style="color: #909090">
<font>
<table align="center" cellpadding="5" cellspacing="0" style="width: 720px">
	<tr>
		<td valign="top" style="width: 391px">
		<table cellpadding="0" cellspacing="0" style="width: 100%">
			<tr>
				<td style="width: 177px">
				<font><a href="http://schilb.info" target="_self"><img align="top" alt="LOGO.PNG" height="128" longdesc="LOGO.PNG" src=".img/logo.png" width="134" class="border-1px"></a></font></td>
				<td>
				<font>
				<a href="?link_id=1" target="_self"><strong>|&nbsp;&nbsp; </strong>Über Uns</a><br>
				<a href="?link_id=2" target="_self"><strong>|&nbsp;&nbsp; </strong>Fotos</a><br>
				<a href="?link_id=3" target="_self"><strong>|&nbsp;&nbsp; </strong>Information</a><br>
				<a href="?link_id=6" target="_self"><strong>|&nbsp;&nbsp; </strong>Gästebuch</a><br>
				<a href="?link_id=4" target="_self"><strong>|&nbsp;&nbsp; </strong>Impressum</a><br>
				<a href="?link_id=5" target="_self"><strong>|&nbsp;&nbsp; </strong>Kontakt</a></font></td>
			</tr>
		</table>
		</td>
		<td valign="top">
		&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
		<hr class="style-one"></td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
		<table cellpadding="10" cellspacing="0" style="width: 100%">
			<tr>
				<td style="height: 39px">
						<?php
# get
$to = $_GET['link_id'];
# include given
if ($to == "1")       			{include("./.content/uberuns.html");}
elseif ($to == "2")  			{include("./.content/fotos.html");}
elseif ($to == "3")  			{include("./.content/information.html");}
elseif ($to == "4")  			{include("./.content/impressum.html");}
elseif ($to == "5")  			{include("./.content/kontakt.html");}
elseif ($to == "0")  			{include("./.content/willkommen.html");}
elseif ($to == "6")  			{include("gb.html");}
elseif ($to == "7")			{include("add.php");}
# secure
	else {
include(".content/willkommen.html");
#else {echo '<meta target="_self" http-equiv="refresh" content="0; URL=/?to=login">';}
}
?>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td colspan="2" valign="top">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2" valign="top"><hr class="style-one"></td>
	</tr>
	<tr>
		<td valign="top" style="width: 391px">&nbsp; ©&nbsp;2021 Schilb. Alle Rechte Vorbehalten. </td>
		<td valign="top" class="text-align-right">&nbsp;<?php include("c.php"); ?>&nbsp;&nbsp; </td>
	</tr>
</table></font>
