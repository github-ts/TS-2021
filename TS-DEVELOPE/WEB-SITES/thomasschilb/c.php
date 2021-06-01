<!DOCTYPE html>
<html>
<head>
<style type="text/css">
.counter {
	color: #FFFFFF;
	font-size: 12pt;
}
</style>
</head>
<?php

$db_location = "127.0.0.1";$db_username = "root";$db_password = "";$db_name = "ts";

/* Counter Start */

$con = mysqli_connect($db_location, $db_username, $db_password) or die ('OFFLINE'); 
mysqli_select_db($con, $db_name) or die ('DATABASE DOES NOT EXISTS'); 
$ip = $_SERVER['REMOTE_ADDR']; // get ip
$test = mysqli_query($con, "SELECT * FROM `thomasschilb` WHERE ip='$ip'"); // test if ip is given
if(mysqli_num_rows($test) != 1) {
mysqli_query($con, "INSERT INTO `thomasschilb` SET ip='$ip'"); // else make db entry
}
$query = mysqli_query($con, "SELECT cid FROM `thomasschilb`");
$num_rows = mysqli_num_rows($query); // get entry rows to number
$overall_counter_state = $num_rows + 2200;


/* Output */

echo $overall_counter_state;
mysqli_close($con);
 
/* Counter End */

?>
