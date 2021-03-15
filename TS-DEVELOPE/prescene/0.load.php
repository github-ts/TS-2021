<!DOCTYPE html>
<html>
<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type">
<meta content="en-us" http-equiv="Content-Language">
<meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
<link rel="icon" type="image/png" href="./.favicon.png">
<link rel="icon" type="image/x-icon" href="./.favicon.ico">
<meta name="author" content="prescene">
<meta name="publisher" content="prescene">
<meta name="copyright" content="© PRESCENE">
<meta name="description" content="ANONYM!ZE - BBS - FTP - ONION - P2P - USENET - WEB - XDCC">
<meta name="keywords" content="pre, prescene, share, sites, software, linklist, disclaimer, legal-disclosure, imprint, nocopy!, bbs, ftp, emule, client, mods, sites, serverlist, kad, torrent, torrents, magnet, dl, search, tracker, usenet, binaries, news, provider, dl, search, web, 0day-releases, zer0day, pre, search, dl, xdcc, bots, channels, network, packets, dl, search, onion, darknet, deepweb">
<meta name="robots" content="all, index, follow"> 
<meta name="googlebot" content="all, index, follow">
<meta name="pagerank" content="10"> 
<meta name="msnbot" content="all,index,follow"> 
<meta name="revisit" content="2 Days"> 
<meta name="revisit-after" content="2 Days"> 
<meta name="alexa" content="100">
<base target="_blank">
<style type="text/css">
@import url('https://fonts.googleapis.com/css2?family=Share+Tech+Mono&display=swap');
font {font-family: 'Share Tech Mono', monospace;font-size: 27px;}
a {
	color: #808080;
	text-decoration: none;
}
a:visited {
	color: #808080;
}
a:active {
	color: #808080;
}
a:hover {
	color: #FFFFFF;
}
.ps1-style1 {
	font-family: "Share Tech Mono", Monospace;
}
.header-title {
	font-family: "Share Tech Mono", Monospace;
	font-size: 128px;
	text-align: center;
}
.ps1-style3 {
	background-color: #333333;
}
.ps1-style4 {
	color: #00FF00;
}
.ps1-style5 {
	color: #FFFFFF;
}
.ps1-style6 {
	font-family: "Share Tech Mono", Monospace;
	text-align: center;
	font-size: 27px;
	background-color: #494949;
}

.ps1-style7 {
	font-family: "Share Tech Mono", Monospace;
	background-color: #555555;
}
.ps1-style8 {
	font-size: 27px;
}
.ps1-style9 {
	text-align: center;
}
.ps1-style10 {
	text-align: left;
	color: #808080;
	font-size: '27px';
}
.ps1-style11 {
	text-align: right;
	color: #808080;
}
.ps1-style12 {
	border-width: 0px;
}
.ps1-style13 {
	text-align: center;
	font-family: "Share Tech Mono";
}
.ps1-style14 {
	border-width: 1px;
	font-size: 27px;
}
.ps1-style15 {
	border-width: 1px;
}
.ps1-style16 {
	text-align: center;
	font-family: "Share Tech Mono";
	font-size: 18px;
}
.ps2-style1 {
	font-size: 27px;
	color: #FFFFFF;
}
.ps2-style2 {
	color: #808080;
}
.ps2-style3 {
	font-family: "Share Tech Mono", Monospace;
	background-color: #494949;
}
.ps2-style4 {
	text-align: left;
	color: #808080;
}
.ps3-style1 {
	text-align: right;
	color: #555555;
	font-size: 27px;
}
.ps3-style2 {
	text-align: center;
	font-family: "Share Tech Mono";
	font-size: 18px;
	color: #555555;
}
.ps4-style1 {
	text-align: center;
	color: #808080;
}
.no-style {
	color: #555555;
	font-size: 27px;
	align: left;
}
</style>
<meta content="pre, scene, re-scene, links" name="description" />
</head>

<body style="color: #808080; margin: 0; background-color: #333333">
<font>
<table cellpadding="15" cellspacing="0" class="ps1-style3" style="width: 100%; height: 100%;">
	<tr>
		<td class="ps3-style1">PRESCENE.2021.V0.1REV5-TS</td>
	</tr>
	<tr>
		<td><br class="ps1-style1" />
		<table align="center" cellpadding="0" cellspacing="0" style="width: 720px">
			<tr>
				<td style="width: 128px">
				<a href="http://prescene.us.to/" target="_self" title="Link Us!">
				<img alt="" class="ps1-style1" src="./.img/medicine.png" /></a></td>
				<td class="header-title"><span class="ps1-style4">PRE</span><span class="ps1-style5">SCENE</span></td>
			</tr>
		</table>
		&nbsp;</td>
	</tr>
	<tr>
		<td class="no-style">No. <?php
/* Connection Data */
$db_location = "127.0.0.1";$db_username = "root";$db_password = "";$db_name = "ts";
/* Open Connection */
$con = mysqli_connect($db_location, $db_username, $db_password) or die ('OFFLINE'); 
mysqli_select_db($con, $db_name) or die ('DATABASE DOES NOT EXISTS'); 
/* Get Remote IP */
$ip = $_SERVER['REMOTE_ADDR'];
/* Query DB */
$test = mysqli_query($con, "SELECT * FROM `prescene` WHERE ip='$ip'"); // test, if ip is already given
if(mysqli_num_rows($test) != 1) {mysqli_query($con, "INSERT INTO `prescene` SET ip='$ip'");} // else, make db entry
$count1 = mysqli_query($con, "SELECT cid FROM `prescene`"); // select, all ids in db
$count = mysqli_num_rows($count1); // count, all ips in db
/* Counter Output */
$alltheyears = $count; // add counter state, exists before truncate the db
echo $alltheyears;
/* Close Connection */
mysqli_close($con);
?></font>
		</td>
	</tr>
	<tr>
		<td class="ps1-style6" style="height: 50px">
		<p><a href="?to=index:section:anonym!ze" target="_self">ANONYM!ZE</a><span class="ps1-style5"> | 
		</span><a href="?to=index:section:bbs" target="_self">BBS</a><span class="ps1-style5"> |
		</span><a href="?to=index:section:ftp" target="_self">FTP</a><span class="ps1-style5"> | 
		</span><a href="?to=index:section:p2p" target="_self">P2P</a><span class="ps1-style5"> |
		</span><a href="?to=index:section:onion" target="_self">ONION</a><span class="ps1-style5"> | 
		</span><a href="?to=index:section:usenet" target="_self">USENET</a><span class="ps1-style5"> | 
		</span><a href="?to=index:section:web" target="_self">WEB</a><span class="ps1-style5"> |
		</span><a href="?to=index:section:xdcc" target="_self">XDCC</a><span class="ps1-style5"> 
		| </span><a href="?to=index:section:addsite" target="_self">+</a></p>
		</td>
	</tr>
	<tr><td class="ps1-style1" style="height: 56px">
<!-- CONTENT LOADER -->
<?php
# get
$to = $_GET['to'];
# welcome
if ($to == "welcome")          				{include("./.content/section/ps-welcome.html");}
# anonym!ze
elseif ($to == "index:section:anonym!ze")  	{include("./.content/section/ps-anonym!ze.html");}
# sections
elseif ($to == "index:section:bbs")  		{include("./.content/section/ps-bbs.html");}
elseif ($to == "index:section:ftp")  		{include("./.content/section/ps-ftp.html");}
elseif ($to == "index:section:addsite")  		{include("./.content/section/ps-addsite.html");}
elseif ($to == "index:section:onion")    	{include("./.content/section/ps-onion.html");}
elseif ($to == "index:section:p2p")    		{include("./.content/section/ps-p2p.html");}
elseif ($to == "index:section:usenet")   	{include("./.content/section/ps-usenet.html");}
elseif ($to == "index:section:web")  		{include("./.content/section/ps-web.html");}
elseif ($to == "index:section:xdcc")   		{include("./.content/section/ps-xdcc.html");}
# site:legal
elseif ($to == "legal:terms")   	{include("./.content/legal/ps-terms.html");}
elseif ($to == "legal:privacy")   		{include("./.content/legal/ps-privacy.html");}
# simple:secure
elseif ($to == "")                  				{include("./.content/section/ps-welcome.html");}
else {include("./.content/section/ps-welcome.html");}
?>
<!-- CONTENT LOADER -->
	<?php
				# simple api for urls
				if (isset($_GET['add'])) {
				# test for valid url
				$url = $_GET['add'];
				if (filter_var($url, FILTER_VALIDATE_URL)) {
				# add valid url to dbfile
					echo '<center><font size="18px" color="#00ff00">+ '.$url.'<br><br>';
					$line = "$url \r\n";
					file_put_contents("./.db/db.txt", $line, FILE_APPEND);
					echo 'Thanks! We have successfully saved your URL. xD</font></center><br>';
				}
				# url not valid
				else {echo '<center><font size="18px" color="red">- '.$url.' ?<br><br>Sorry! No valid URL could be found.<br><br>';
				echo 'Use it like this:<br>http://prescene.us.to/?add=<strong>https://yoursite.com</strong><br></font></center>';}
				}
				# url is not given, do nothing
				else {echo '';}
				?>
		</td>
	</tr>
	<tr><td class="ps1-style1" style="height: 56px">
		&nbsp;</td>
	</tr>
	<tr>
		<td class="ps2-style3">
		
		
		
		
		
		
		
		
		<table align="left" cellpadding="10" cellspacing="0" style="width: 555px">
			<tr class="ps1-style4">
				<td class="ps1-style8" colspan="2" valign="top">&nbsp;</td>
				<td class="ps1-style8" valign="top">&nbsp;</td>
			</tr>
			<tr class="ps1-style4">
				<td class="ps2-style1" valign="top" style="height: 51px">Index</td>
				<td class="ps2-style1" valign="top" style="height: 51px"></td>
				<td class="ps2-style1" valign="top" style="height: 51px">Legal</td>
			</tr>
			<tr>
				<td valign="top" style="width: 180px"><span class="ps1-style8">
				<a href="?to=index:section:anonym!ze" target="_self">Anonym!ze</a></span><br class="ps1-style8" />
				<span class="ps1-style8"><a href="?to=index:section:bbs" target="_self">BBS</a></span><br class="ps1-style8" />
				<span class="ps1-style8"><a href="?to=index:section:ftp" target="_self">FTP</a></span><br class="ps1-style8" />
				<span class="ps1-style8"><a href="?to=index:section:onion" target="_self">Onion</a></span><br class="ps1-style8" />
				<span class="ps1-style8"><a href="?to=index:section:p2p">P2P</a></span></td>
				<td class="ps1-style8" valign="top" style="width: 180px">
				<a href="?to=index:section:usenet" target="_self">Usenet</a><br />
				<a href="?to=index:section:web" target="_self">Web</a><br />
				<a href="?to=index:section:xdcc" target="_self">XDCC</a></td>
				<td class="ps1-style8" valign="top" style="width: 180px">
				<a href="?to=legal:terms" target="_self">Terms and Conditions</a><br><br />
				<a href="?to=legal:privacy" target="_self">Privacy Policy</a></td>
			</tr>
			<tr class="ps1-style4">
				<td class="ps1-style8" colspan="2" valign="top">&nbsp;</td>
				<td class="ps1-style8" valign="top">&nbsp;</td>
			</tr>
			</font>
			</table>
		</td>
	</tr>
</table>

<p class="ps4-style1">&copy; 2021 PRESCENE. ALL RIGHTS RESERVED.&nbsp;</p>
<p class="ps1-style16">
				<a href="http://prescene.us.to/" target="_blank" title="PS | PRESCENE">
				<img alt="ps-banner" height="31" longdesc="prescene-banner-small" src="./.img/prescene-banner.png" width="88" class="ps1-style12" /></a>&nbsp;<br />
				<span class="ps2-style2">88x31px</span></p>
<p class="ps3-style2">&nbsp;</p>
<p class="ps3-style2">			
</p>
</body>

</html>
