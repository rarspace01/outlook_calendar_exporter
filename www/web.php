<?php
	
	$remoteFile = $_POST['calendarFile'];
	
	$bin = base64_decode($remoteFile);
	
	file_put_contents("cal.ics",$bin);
	
?>
