<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta content="de" http-equiv="Content-Language" />
<link rel="icon" type="image/png" href="./.favicon.png" />
<link rel="icon" type="image/x-icon" href="./.favicon.ico" />
<meta content="ts, thomasschilb, thomas, schilb, web, radio, audio, stream, audiostream, webradio, dsr, demoscene, radio" name="keywords" />
<meta content="Radio" name="description" />
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>TS | Radio</title>
<link href="https://fonts.googleapis.com/css2?family=Roboto+Mono:wght@100&display=swap" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Roboto+Mono+Thin:@100&display=swap" rel="stylesheet">
<style type="text/css">

/* HYPERLINK */

a {
	color: #56D4FF;
	text-decoration: none;
}
a:visited {
	color: #56D4FF;
}
a:active {
	color: #56D4FF;
}
a:hover {
	color: #FFFFFF;
	text-decoration: underline;
}

/* FONT */

font {
	font-family: 'Roboto Mono', monospace;
	font-size: 23pt;
}

/* LINE */

hr { 
    position: relative; 
    top: 0px; 
    border: 0px; 
    height: 5px; 
    background: #56D4FF; 
    margin-bottom: 0px; 
}

/* TITLE */

.title {
	font-size: 72pt;
	text-align: center;
}

/* BUTTON */

#play-pause-button{
  font-size: 50px;
  cursor: pointer;
}

/* REST */

.text-center {
	text-align: center;
}
.ts-style3 {
	font-family: "Roboto Mono", Monospace;
	font-size: 23pt;
}
.black {
	color: #808080;
}
.font-family {
	font-family: "Roboto Mono", Monospace;
}
.ts-style6 {
	text-align: left;
	font-size: 23pt;
	background-color: #222222;
	color: #808080;
	font-family: 'Roboto Mono', monospace;
}
.ts-style9 {
	background-color: #232323;
}
.ts-style10 {
	font-family: "Roboto Mono", Monospace;
	font-size: 23pt;
	text-align: right;
}
.ts-style11 {
	background-color: #222222;
}
.ts-style13 {
	font-family: "Roboto Mono", Monospace;
	font-size: 23pt;
	text-align: center;
}
.ts-style14 {
	background-color: #222222;
	text-align: center;
	font-family: "Roboto Mono", monospace;
	font-weight: 100;
	font-size: 88pt;
	color: #FFFFFF;
}
.ts-style15 {
	background-color: #222222;
	text-align: right;
	font-family: "Roboto Mono", Monospace;
	font-size: 23pt;
	color: #FFFFFF;
}
.ts-style16 {
	color: #56D4FF;
}
.ts-style17 {
	text-align: left;
}
.ts-style21 {
	font-family: "Roboto Mono", Monospace;
	font-size: 23pt;
	color: #808080;
}
.ts-style22 {
	text-decoration: underline;
}
.ts-style23 {
	color: #FFFFFF;
}
.ts-style24 {
	color: #FF0000;
}
.ts-style25 {
	color: #00FF00;
}
.ts-style26 {
	font-family: "Roboto Mono", Monospace;
	font-size: 23pt;
	text-align: center;
	color: #FFFFFF;
}
.ts-style27 {
	color: #FFFFFF;
	background-color: #333333;
}
</style>
<base target="_self" />
</head>



<script type="text/javascript">
var audio = new Audio("http://ts-radio.com:8064");

$('#play-pause-button').on("click",function(){
  if($(this).hasClass('fa-play'))
   {
     $(this).removeClass('fa-play');
     $(this).addClass('fa-pause');
     audio.play();
   }
  else
   {
     $(this).removeClass('fa-pause');
     $(this).addClass('fa-play');
     audio.pause();
   }
});

audio.onended = function() {
     $("#play-pause-button").removeClass('fa-pause');
     $("#play-pause-button").addClass('fa-play');
};
</script>

<body style="color: #56D4FF; margin: 0; background-color: #232323">

<table cellpadding="23" cellspacing="0" style="width: 100%">
	<tr>
		<td class="ts-style10">
		<div class="ts-style17">
			<strong>STATION</strong><br></div>
		<hr></td>
	</tr>
	<tr>
		<td class="ts-style11">
		<table align="center" cellpadding="23" cellspacing="0" style="width: 100%" class="ts-style9">
			<tr>
				<td class="ts-style15">
				&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td class="ts-style14" valign="top">
				<span class="ts-style16">TS</span>.RADIO</td>
			</tr>
			<tr>
				<td class="ts-style6">
				<br />
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td class="ts-style3"><strong>COVER</strong><hr></td>
	</tr>
	<tr>
		<td class="ts-style13">&nbsp;<br />
		<a href="radio.png" target="_blank">
		<img alt="tsr-logoV3" height="256" longdesc="tsr-logoV3" src="radio.png" width="256" /></a></td>
	</tr>
	<tr>
		<td class="ts-style3"><strong><br />
		TITLE</strong><hr></td>
	</tr>
	<tr>
		<td class="ts-style13"><?php include("get-title.php");?>&nbsp;
		<a id="play-pause-button" class="fa fa-play"></a>
		<br /></td>
	</tr>

	<tr>
		<td class="ts-style3"><strong><br />
		PLAY</strong><hr style="left: -3px; top: 0px"></td>
	</tr>
<tr><td><center><font>
&nbsp;<span class="ts-style23"><br />
		<audio controls style="width: 266px; height: 80px"><source src="http://radio.thomasschilb.online:8032" id="radio" autoplay></audio>
		<br />
		</span><span class="ts-style25">[ONLINE]</span><span class="ts-style23"><br />
		</span><span class="black"><br />
	32K / AAC+ / MONO</span><br /></font></center>
		</td>
	</tr>
	<tr>
		<td class="ts-style3"><strong>INFO</strong><hr style="left: -3px; top: 0px"></td>
	</tr>
	<tr>
		<td class="ts-style3"><font size="5" class="font-family">
		<span class="ts-style23">Stream-URL<br />
		<br />
		</span>
		<table align="left" style="width: 100%">
			<tr>
				<td class="ts-style27">
				WEB</td>
			</tr>
			<tr>
				<td>
				<a href="http://radio.thomasschilb.online:8032">http://radio.thomasschilb.online:8032</a></td>
			</tr>
			<tr>
				<td class="black">
				<em>32K, AAC+, MONO, PORT: 8032</em></td>
			</tr>
			<tr>
				<td class="black">
				&nbsp;</td>
			</tr>
			<tr>
				<td class="ts-style27">STANDARD</td>
			</tr>
			<tr>
				<td><font size="5" class="font-family">
				<a href="http://radio.thomasschilb.online:8001">
				http://radio.thomasschilb.online:8001</a></font></td>
			</tr>
			<tr>
				<td class="black"><em>192K, MP3, STEREO, PORT: 8001</em></td>
			</tr>
			</table>
		<span class="ts-style22">
		<span class="black"><br />
		</span></span><br />
		<br />
<br />
		<br />
		<br class="black" />
		<span class="black"><br class="ts-style22" />
		<br />
		</span>
		<span class="ts-style23"><br />
		<br />
		Genre</span><span class="ts-style22"><span class="black"><br />
		</span></span>
		<span class="black"><br />
		Techno, Hardtechno &amp; Schranz<br />
		<em><br />
		</em></span>
		<em>
		<br class="ts-style24" />
		<span class="black">
		<span class="ts-style24">We are playing the latest pre-releases &amp; rare!<br />
		</span></span></em><br />
		<br />
		</em>
		<span class="ts-style23">
		Need a Player?</span><span class="black">
		Try 
		</span>
		<a href="https://qmmp.ylsoftware.com" target="_blank">QMMP</a><span class="black">, 
		<a href="winamp58_3660_beta_full_en-us.exe">Winamp</a> 
		or <a href="https://www.videolan.org/vlc/index.html" target="_blank">VLC</a></span></font></td>
	</tr>
		<tr>
		<td class="ts-style3"><strong><br />
		SUPPORT</strong><hr></td>
	</tr>
	<tr>
		<td class="ts-style21">
		<span class="ts-style23">Groups</span><span class="black"><br />
		<br />
		1KiNG,
		AGITB, AOV, BABAS, BF, JUSTiFY, KNOWN, NDE, PTC, TALiON, TraX, wAx, x0x0s,<br />
		and the others...<br />
		<br />
		</span><span class="ts-style23">Web</span><span class="black"><br />
		<br />
		RARBG, Tixati</span></td>
	</tr>
	<tr>
		<td class="ts-style21"><span class="ts-style16"><strong><br />
		COPYRIGHT</strong></span><hr></td>
	</tr>
	<tr>
		<td class="ts-style26">&copy; 2021 
		TS-RADIO. All Rights 
		Reserved.<br />
		<br />
		<br />
</td>
	</tr>
</table>

</body>

</html>
