PHP Computer Inventory System

Install

1) Copy the 'Inventory.vbs' from the scripts folder to your network scripts folder

2) Change settings in 'Inventory.vbs' to your required file paths

3) Set permissions on 'inventorylog.csv' in your desired location so that domain users can write to it

4) Test the Inventory.vbs by double clicking on it manually. You should see the Inventorylog.csv file populate with some data from your machine

5) Set Inventory.vbs as a computer login script using Group Policy (Computer Configuration >> Policies >> Windows Settings >> Scripts >> Logon)

6) Setup a web server with PHP installed. Easy way to do it on windows here http://www.microsoft.com/web/platform/phponwindows.aspx

7) Copy the Inventorylog.csv file into the /csv folder of your web server. Usually something like c:\inetpub\wwwroot\inventory\csv\

8) Change any settings in the 'settings.php' file such as Brands and CPU search strings. You can also change the locations of the csv file if you wish.

9) Load up a web browser and navigate to http://localhost/inventory