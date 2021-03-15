<?php

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
?>
