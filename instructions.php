<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">


<?php 

////////////////////////////
/// PHP Inventory System ///
////////////////////////////

include 'settings.php';

?>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>PHP Inventory System</title>
<meta name="keywords" content="" />
<meta name="description" content="" />
<link href="css/style.css" rel="stylesheet" type="text/css" media="screen" />
</head>
<body>
<div id="logo">
	<h1><a href="#">PHP Inventory System</a></h1>
	<h2></h2>
</div>
<div id="menu">
	<ul>
		<li class="first"><a href="index.php" accesskey="1" title="">Home</a></li>
		<li><a href="instructions.php" accesskey="2" title="">Instructions</a></li>
		<li><a href="about.php" accesskey="4" title="">About</a></li>
	</ul>
</div>
<hr />
<!-- start page -->
<div id="page">
	<!-- start content -->
	<div id="content">
		<div class="post">

			<div class="entry">
				

<b>Instructions</b><br><br>

1) Copy the 'Inventory.vbs' from the scripts folder to your network scripts folder<br><br>

2) Change settings in 'Inventory.vbs' to your required file paths<br><br>

3) Set permissions on 'inventorylog.csv' in your desired location so that domain users can write to it<br><br>

4) Test the Inventory.vbs by double clicking on it manually. You should see the Inventorylog.csv file populate with some data from your machine<br><br>

5) Set Inventory.vbs as a computer login script using Group Policy (Computer Configuration >> Policies >> Windows Settings >> Scripts >> Logon)<br><br>

6) Setup a web server with PHP installed. Easy way to do it on windows here http://www.microsoft.com/web/platform/phponwindows.aspx<br><br>

7) Copy the Inventorylog.csv file into the /csv folder of your web server. Usually something like c:\inetpub\wwwroot\inventory\csv\<br><br>

8) Change any settings in the 'settings.php' file such as Brands and CPU search strings. You can also change the locations of the csv file if you wish.<br><br>

9) Load up a web browser and navigate to http://localhost/inventory<br><br>
				
				
			</div>
		
		</div>
		
	</div>
	<!-- end content -->
	<!-- start sidebar -->
	<div id="sidebar">
		<ul>
			
			<li>
				<h2>About</h2>
				<p>A simple system to show the assets in your organization</p>
			</li>
			<li>
				<h2>Links</h2>
				<ul>
					<li><a href="#">Running VBS files on Login</a></li>
					<li><a href="#">All MAC Addresses Mod</a></li>
				</ul>
			</li>
		</ul>
	</div>
	<!-- end sidebar -->
</div>
<!-- end page -->
<div id="footer">
	<p class="legal">OpenSourced 2013</p>
	<p class="credit">Sourcode <a href="http://www.github.com">GitHub</a></p>
</div>
</body>
</html>
