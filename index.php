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
			<h2 class="title"><a href="#">Device Inventory</a></h2>
			<div class="entry">
				<h3>List of all devices found on network</h3>
				
				<table><tr>
				<td width='180'><b>Name</b></td>
				<td width='80'><b>Brand</b></td>
				<td width='80'><b>RAM</b></td>
				<td width='80'><b>CPU</b></td>
				<td width='80'><b>HDD</b></td>
				<td width='80'><b>Size</b></td>
				<td width='80'><b>Free</b></td></tr>
				
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
					$hddtotal = $data[21];
					$hddfree = $data[22];
					$hddfreeint = strlen($hddfree);
					$hddtype = $data[25];
					$ramtotal = $data[6];
					$cpu = $data[34];
					$ramtotal = round(number_format((substr($ramtotal, 0, 2)/10),1));
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
					
					// Display the RAM total
					echo "<td>".$ramtotal." GB</td>";
						
					// Loop through the CPU list and display an icon
					foreach($cpuconfig as $cpuname=>$cpupic)
					{
						if (strpos($cpu,$cpuname) !== false) {
							echo "<td><img src='images/cpu/$cpupic'></td>";
						$checkcpu = 1;
						}
					}
					if ($checkcpu != 1) {echo "<td></td>";}
					
					// Display the Hard Disk Type
					if (strpos($hddtype,'SSD') !== false) {echo "<td><img src='images/disk/ssd.png'></td>";}
					else {echo "<td><img src='images/disk/harddisk.png'></td>";}
					
					// Display the Hard Disk Total
					echo "<td>".$str = str_replace("1:", "", $hddtotal)."</td>";
					
					// Display the Hard Disk Free. Warn in red if under 10GB left
					if ($hddfreeint < 8) {echo "<td><font color='red'>".$str = str_replace("1:", "", $hddfree)."</font></td>";}
					else {echo "<td>".$str = str_replace("1:", "", $hddfree)."</td>";}
					
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
