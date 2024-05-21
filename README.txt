<#==============================================================================
         File Name : IP-Phone-Inventory.ps1
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr.com)
                   : 
       Description : Uses the built-in web server on Cisco VoIP phones to pull down and record
                   : information across multiple network ranges, then stores it all in an Excel
                   : spreadsheet.
                   : 
             Notes : Normal operation is with no command line options. Edit the list of ranges below
                   : to include all vlans you wish to scan.  The node ranges are set for .15 thru
                   : .254.  The script checkes each IP in the range.  The ranges are assumed to only 
                   : contain Cisco VoIP phones.  Other devices can choke then script. If a ping response 
                   : is seen, it then attempts to find if TCP port 80 is open.  If found it attempts to 
                   : load three HTML pages from the device.  It then parses the output to collect the data.
                   : NOTE: Write-host statements are at each step of the parsing step and commented out 
                   : for trouble shooting purposes.
                   :
      Requirements : Requires the PSParseHTML module which will be loaded automatically.
                   : PowerShell v5 is required.  MS Excel is reqired.
                   : A flat text file is required in the script folder and named "ranges.txt"
                   : This file should contain the first 3 octets of a subnet to scan, a comma, and 
                   : a location name.  Example: 192.168.10,IDF1.  Add as many lines as required.
                   : Any line prefixed with a "#" will be ignored. 
                   : 
   Option Switches : See descriptions above.
                   :
          Warnings : Nothing in this script is destructive.
                   :   
             Legal : Public Domain. Modify and redistribute freely. No rights reserved.
                   : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF 
                   : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED.
                   : That being said, feel free to ask if you have questions...
                   :
           Credits : Code snippets and/or ideas came from many sources...
                   : 
    Last Update by : Kenneth C. Mazie                                           
   Version History : v1.00 - 05-21-24 - Original release
    Change History : v1.10 - 00-00-00 - 
                   : 
                   :                  
==============================================================================#>
