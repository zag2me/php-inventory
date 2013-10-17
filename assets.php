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
		<li><a href="network.php" accesskey="2" title="">Network</a></li>
		<li><a href="assets.php" accesskey="2" title="">Assets</a></li>
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
			<h2 class="title"><a href="#">Device Inventory</a></h2>
			<div class="entry">
				<h3>List of all devices found on network</h3>
				
				<table><tr>
				<td width='150'><b>Name</b></td>
				<td width='60'><b>Brand</b></td>
				<td width='320'><b>Model</b></td>
				<td width='200'><b>Serial</b></td>
				
				<?php
				
				// Define Variables
				$counter = 0;
							
				// Check CSV exists and open
				if (($handle = fopen($csvlocation, "r")) !== FALSE) 
				{
				
				// Loop Through CSV and put into an array
				while (($data = fgetcsv($handle, 50000, ",")) !== FALSE) 
				{
					// Reset some checking variables on each device
					$checkbrand = 0;
					$checkcpu = 0;
					
					// Dump Variable (uncomment below if you want to see the entire array)
					if ($debug == 'yes'){
						echo "<pre>".print_r($data)." <br /></pre>";
					}
							
					// Keep a counter of devices
					$counter ++;
		
					// Skip the first item
					if ($counter != 1)
					{
					
					// Define easy name variables
					$name = $data[3];
					$serial = $data[14];
					$model = $data[13];
					$brand = $data[12];
					
					// Print the device table
					echo "<tr>"; 
					
					// Display the name and set Anchor
					echo "<td>$name <a name='$name'></td>";
					
					// Loop through the Brand list and display an icon
					foreach($brandconfig as $brandname=>$brandpic)
					{
						if (strpos($brand,$brandname) !== false) {
							echo "<td><img src='images/brands/$brandpic'></td>";
						$checkbrand = 1;
						}
					}
					if ($checkbrand != 1) {echo "<td></td>";}
					
					// Display the Model Name
					echo "<td>".$model."</td>";			
				
					// Display the Serial
					echo "<td>".$serial."</td>";	
					
				
					//End device item row
					echo "</tr>";
					}
				}
				
				}
				?>
				
				</table>
				
				
			</div>
			<p class="meta">Total Found &nbsp;|&nbsp; <?php $counter = $counter - 1 ; echo $counter;?> Devices</p>
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
