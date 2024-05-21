
Param(
    [switch]$Console = $false,         #--[ Set to true to enable local console result display. Defaults to false ]--
    [switch]$Debug = $False            #--[ Generates extra console output for debugging.  Defaults to false ]--
 
    )

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
Clear-Host
#Requires -version 5

#$Console = $true   #--[ uncomment for trouble shooting ]--

#--[ Install the PSParseHTML module on demand ]--
If (-not (Get-Module -ErrorAction Ignore -ListAvailable PSParseHTML)){
    If ($Console){Write-Host "Installing PSParseHTML module for the current user..."}
    Install-Module -Scope CurrentUser PSParseHTML -ErrorAction Stop
  }

#--[ Injest target file ]--
$Ranges = @()
if (Test-Path -Path "$PSScriptRoot\Ranges.txt" -PathType Leaf){  
    ForEach ($Line in (Get-Content "$PSScriptRoot\Ranges.txt")){
        If (!($Line -like '#*')){
            $Address = $Line.Split(",")[0]
            $Location = $Line.Split(",")[1]
            $Ranges += ,@($Address,$Location)
        }
    }Else{
        Write-Host "MISSING IP RANGE FILE.  File is required.  Script aborted..." -ForegroundColor " Red"
        break;break;break
    }
}

#--[ Create new Excel COM object ]--
$Excel = New-Object -ComObject Excel.Application -ErrorAction Stop
#StatusMsg "Creating new Spreadsheet..." "green"  
$Workbook = $Excel.Workbooks.Add()
$Worksheet = $Workbook.Sheets.Item(1)
$Worksheet.Activate()
$WorkSheet.Name = "VoIP Phones"
[int]$Col = 1
[Int]$Row = 1
$WorkSheet.cells.Item($Row,$Col++) = "IP Address"          # A
$WorkSheet.cells.Item($Row,$Col++) = "Location"            # B
$WorkSheet.cells.Item($Row,$Col++) = "Model #"             # C
$WorkSheet.cells.Item($Row,$Col++) = "Serial #"            # D
$WorkSheet.cells.Item($Row,$Col++) = "Name"                # E
$WorkSheet.cells.Item($Row,$Col++) = "MAC Address"         # F
$WorkSheet.cells.Item($Row,$Col++) = "DN"                  # G
$WorkSheet.cells.Item($Row,$Col++) = "Version"             # H
$WorkSheet.cells.Item($Row,$Col++) = "Hardware Rev"        # I  
$WorkSheet.cells.Item($Row,$Col++) = "Port Speed"          # J
$WorkSheet.cells.Item($Row,$Col++) = "TFTP Server"         # K
$WorkSheet.cells.Item($Row,$Col++) = "DHCP Server"         # L
$WorkSheet.cells.Item($Row,$Col++) = "Unified CM 1"        # M
$WorkSheet.cells.Item($Row,$Col++) = "Unified CM 2"        # N   
$Range = $WorkSheet.Range(("A1"),("N1")) 
$Range.font.bold = $True
$Range.HorizontalAlignment = -4108  # Alignment Middle
$Range.Font.ColorIndex = 44
$Range.Font.Size = 12
$Range.Interior.ColorIndex = 56
$Range.font.bold = $True
1..4 | ForEach-Object {
    $Range.Borders.Item($_).LineStyle = 1
    $Range.Borders.Item($_).Weight = 4
}
$Resize = $WorkSheet.UsedRange
[Void]$Resize.EntireColumn.AutoFit()
$Excel.Visible = $True
$WorkSheet = $Workbook.WorkSheets.Item("VoIP Phones")
$WorkSheet.activate()

If ($Console){Write-Host "--[ Begin ]-------------------------------------------"}
$Row = 2

ForEach ($Range in $Ranges){
    $Node = 15
    While ($Node -le 254){
        $Col = 1
        $Target = $Range[0]+"."+$Node
        $Site = $Range[1]

        If ($Console){Write-Host "`nCurrent Target :"$Target -ForegroundColor Yellow }
       
        If ([System.Net.Sockets.TcpClient]::new().ConnectAsync($Target, 80).Wait(100) ){  #--[ Quick port 80 ping ]--
            $DeviceUri = "http://$Target/CGI/Java/Serviceability?adapter=device.statistics.device"
            $NetworkUri = "http://$Target/CGI/Java/Serviceability?adapter=device.statistics.configuration"
            $PortUri = "http://$Target/CGI/Java/Serviceability?adapter=device.statistics.port.network"
            $Device = ConvertFrom-Html -Engine AngleSharp -Url $DeviceUri 
            $DeviceDetail = ($Device.textcontent).Split([Environment]::NewLine, [StringSplitOptions]::RemoveEmptyEntries)
            $Network = ConvertFrom-Html -Engine AngleSharp -Url $NetworkUri 
            $NetworkDetail = ($Network.textcontent).Split([Environment]::NewLine, [StringSplitOptions]::RemoveEmptyEntries)
            $Port = ConvertFrom-Html -Engine AngleSharp -Url $PortUri 
            $PortDetail = ($Port.textcontent).Split([Environment]::NewLine, [StringSplitOptions]::RemoveEmptyEntries)
    
            $Model =""
            $Serial =""
            $Name =""
            $MAC =""
            $DN =""
            $Line = ""
            $Ver = ""
            $HardwareRev = ""
            $IP = ""
            $Speed = ""
            $TFTP = ""
            $DHCP = ""
            $UCM = ""
            $Len = ""
            $UCM1 = ""
            $UCM2 = ""

            #--[ Known Model Numbers ]------------------------
              #   CP-7962G
              #   CP-7942G
              #   CP-8851
              #   CP-8841
              #   CP-8831   
                    
            #==[ Device ]=================================================================    
            $Counter = 0
            Foreach ($line in $DeviceDetail){
            # write-host $Counter  $line -ForegroundColor $Counter
                If ($line.split(" ")[4] -Like "*CP*"){
                    $Model = $line.split(" ")[4]
                    #write-host $Model -ForegroundColor $Counter
                }
                If ($line.split(" ")[5] -Like "*CP*"){
                    $Model = $line.split(" ")[5]
                    #write-host $Model -ForegroundColor $Counter
                }    
                If ($Counter -eq 12){
                    #write-host $line -ForegroundColor $Counter
                    foreach ($item in $line){
                        If ($Model -like "*CP-7942*"){
                            #  $line.split(" ")[15]
                            #$Model = $x.split(" ")[83]                    #--[ Alternate ]--
                            #$Name = $item.split(" ")[4].split("e")[1]     #--[ Alternate ]--
                            #$Name = $x.split(" ")[51].split("e")[2]       #--[ Alternate ]--
                            #$MAC = (($item.split(" ")[2]).split("s")[2]) -replace '..(?!$)', '$&:'    #--[ Alternate ]--
                            $DN =  ($item.split(" ")[6]).Split("N")[1]
                            $Serial = ($item.split(" ")[17]).split("r")[1]     
                            #$Serial = ($item.split(" ")[31])              #--[ Alternate ]--
                            $Ver = $item.split(" ")[30]
                            $HardwareRev = ($item.split(" ")[15]).split("n")[1]     
                        }
                        If ($Model -like "*CP-7962*"){
                            #  $line.split(" ")[15]
                            #$Model = $x.split(" ")[83]                    #--[ Alternate ]--
                            #$Name = $item.split(" ")[4].split("e")[1]     #--[ Alternate ]--
                            #$Name = $x.split(" ")[51].split("e")[2]       #--[ Alternate ]--
                            #$MAC = (($item.split(" ")[2]).split("s")[2]) -replace '..(?!$)', '$&:'    #--[ Alternate ]--
                            $DN =  ($item.split(" ")[6]).Split("N")[1]
                            $Serial = ($item.split(" ")[23]).split("r")[1]     
                            #$Serial = ($item.split(" ")[31])              #--[ Alternate ]--
                            $Ver = $item.split(" ")[36]
                            $HardwareRev = ($item.split(" ")[21]).split("n")[1]     
                        }
                        If ($Model -like "*CP-8831*"){     
                            #$item.split(" ")[45]
                            #$Model = $x.split(" ")[83]                    #--[ Alternate ]--
                            #$Name = $item.split(" ")[4].split("e")[1]     #--[ Alternate ]--
                            #$Name = $item.split(" ")[51].split("e")[2]    #--[ Alternate ]--
                            #$MAC = (($item.split(" ")[2]).split("s")[2]) -replace '..(?!$)', '$&:'   #--[ Alternate ]--
                            $DN =  ($item.split(" ")[6]).Split("N")[1]
                            $Serial = ($item.split(" ")[11]).split("r")[1]     
                            #$Serial = ($item.split(" ")[25])              #--[ Alternate ]--
                            $Ver = $item.split(" ")[24]
                            $HardwareRev = ($item.split(" ")[9]).split("n")[1]   
                        }
                        If ($Model -like "*CP-8841*"){     
                            #$item.split(" ")[45]
                            #$Model = $x.split(" ")[83]                    #--[ Alternate ]--
                            #$Name = $item.split(" ")[10].split("e")[1]    #--[ Alternate ]--
                            #$Name = $x.split(" ")[51].split("e")[2]       #--[ Alternate ]--
                            #$MAC = (($item.split(" ")[8]).split("s")[2]) -replace '..(?!$)', '$&:'   #--[ Alternate ]--
                            $DN =  ($item.split(" ")[12]).Split("N")[1]
                            $Serial = ($item.split(" ")[23]).split("r")[1]     
                            #$Serial = ($item.split(" ")[44])              #--[ Alternate ]--
                            $Ver = $item.split(" ")[35]
                            $HardwareRev = ($item.split(" ")[21]).split("n")[1]   
                        }
                        If ($Model -like "*CP-8851*"){     
                            #$item.split(" ")[45]
                            #$Model = $x.split(" ")[83]                    #--[ Alternate ]--
                            #$Name = $item.split(" ")[10].split("e")[1]    #--[ Alternate ]--
                            #$Name = $x.split(" ")[51].split("e")[2]       #--[ Alternate ]--
                            #$MAC = (($item.split(" ")[8]).split("s")[2]) -replace '..(?!$)', '$&:'
                            $DN =  ($item.split(" ")[12]).Split("N")[1]
                            $Serial = ($item.split(" ")[31]).split("r")[1]     
                            #$Serial = ($item.split(" ")[44])              #--[ Alternate ]--
                            $Ver = $item.split(" ")[43]
                            $HardwareRev = ($item.split(" ")[29]).split("n")[1]   
                        }
                    }
                }
                $Counter++
            }

            #==[ Network ]=================================================================
            $Counter = 0
            Foreach ($line in $NetworkDetail){
                # write-host $Counter  $line -ForegroundColor $Counter
                If ($Counter -eq 12){
                    # write-host $line -ForegroundColor $Counter
                    foreach ($item in $line){
                        $ErrorActionPreference = "silentlycontinue"
                        If ($Model -like "*CP-7942*"){
                            #  $line.split(" ")[62]
                            $IP = $item.split(" ")[12].split("s")[2] 
                            $Name = $item.split(" ")[8].substring(4) 
                            $MAC = (($item.split(" ")[6]).split("s")[2]) -replace '..(?!$)', '$&:'
                            $TFTP = ($item.split(" ")[17]).substring(1) 
                            $DHCP = ($item.split(" ")[2]).Split("r")[2]
                            $UCM = ($item.split(" ")[56])
                            $Len = $UCM.Length
                            $UCM1 = $UCM.Substring($Len -18)
                            $UCM2 = ($item.split(" ")[62]).Substring($Len -18) 
                        }
                        If ($Model -like "*CP-7962*"){
                            # $line.split(" ")[62]
                            $IP = $item.split(" ")[12].split("s")[2] 
                            $Name = $item.split(" ")[8].substring(4) 
                            $MAC = (($item.split(" ")[6]).split("s")[2]) -replace '..(?!$)', '$&:'
                            $TFTP = ($item.split(" ")[17]).substring(1)
                            $DHCP = ($item.split(" ")[2]).Split("r")[2]
                            $UCM = ($item.split(" ")[56])
                            $Len = $UCM.Length
                            $UCM1 = $UCM.Substring($Len -18)
                            $UCM2 = ($item.split(" ")[62]).Substring($Len -18) 
                        }
                        If ($Model -like "*CP-8831"){
                            # $line.split(" ")[6]
                            $IP = $item.split(" ")[10].split("s")[2] 
                            $Name = $item.split(" ")[6].substring(4) 
                            $MAC = (($item.split(" ")[4]).split("s")[2]) -replace '..(?!$)', '$&:'
                            $TFTP = ($item.split(" ")[15]).substring(1)  
                            $DHCP = ($item.split(" ")[2]).Split("r")[2]
                            $UCM = ($item.split(" ")[57])
                            $Len = $UCM.Length
                            $UCM1 = $UCM.Substring($Len -18)
                            $UCM2 = ($item.split(" ")[63]).Substring($Len -18)
                            }
                        If ($Model -like "*CP-8841"){
                            # $line.split(" ")[62]
                            $IP = $item.split(" ")[13].split("s")[2] 
                            $Name = $item.split(" ")[4].substring(4) 
                            $MAC = (($item.split(" ")[2]).split("s")[2]) -replace '..(?!$)', '$&:'
                            $TFTP = ($item.split(" ")[31]).substring(1)
                            $DHCP = ($item.split(" ")[8]).Split("r")[2]
                            $UCM = ($item.split(" ")[45])
                            $Len = $UCM.Length
                            $UCM1 = $UCM.Substring($Len -18)
                            $UCM2 = ($item.split(" ")[49]).Substring($Len -18)         
                        }
                        If ($Model -like "*CP-8851"){
                            # $line.split(" ")[62]
                            $IP = $item.split(" ")[13].split("s")[2] 
                            $Name = $item.split(" ")[4].substring(4) 
                            $MAC = (($item.split(" ")[2]).split("s")[2]) -replace '..(?!$)', '$&:'
                            $TFTP = ($item.split(" ")[31]).substring(1)
                            $DHCP = ($item.split(" ")[8]).Split("r")[2]
                            $UCM = ($item.split(" ")[45])
                            $Len = $UCM.Length
                            $UCM1 = $UCM.Substring($Len -18)
                            $UCM2 = ($item.split(" ")[49]).Substring($Len -18) 
                        }
                    }
                }
                $Counter++
            }
            #==[ Port Speed ]=================================================================
            $Counter = 0
            Foreach ($line in $PortDetail){
                #write-host $Counter  $line -ForegroundColor $Counter
                If ($Counter -eq 12){
                    $Speed = $line.split(" ")[-1]
                }
                $Counter++
            }

            #==[ Results ]====================================================================
            If ($Console){
                Write-host "  Row          = "$Row
                Write-host "  Site         = "$Site
                Write-host "  Model        = "$Model
                Write-host "  Serial       = "$Serial
                Write-host "  Name         = "$Name
                Write-host "  MAC Address  = "$MAC
                Write-host "  DN           = "$DN
                Write-host "  Version      = "$Ver
                Write-host "  Hardware Rev = "$HardwareRev
                Write-host "  IP Address   = "$IP
                Write-host "  Port Speed   = "$Speed
                Write-host "  TFTP Server  = "$TFTP
                Write-host "  DHCP Server  = "$DHCP
                Write-host "  Unified CM 1 = "$UCM1
                Write-host "  Unified CM 2 = "$UCM2
            }

            $WorkSheet.cells.Item($Row, $Col++) = $IP
            $WorkSheet.cells.Item($Row, $Col++) = $Site
            $WorkSheet.cells.Item($Row, $Col++) = $Model
            $WorkSheet.cells.Item($Row, $Col++) = $Serial
            $WorkSheet.cells.Item($Row, $Col++) = $Name
            $WorkSheet.cells.Item($Row, $Col++) = $MAC
            $WorkSheet.cells.Item($Row, $Col++) = $DN
            $WorkSheet.cells.Item($Row, $Col++) = $Ver
            $WorkSheet.cells.Item($Row, $Col++) = $HardwareRev
            $WorkSheet.cells.Item($Row, $Col++) = $Speed
            $WorkSheet.cells.Item($Row, $Col++) = $TFTP
            $WorkSheet.cells.Item($Row, $Col++) = $DHCP
            $WorkSheet.cells.Item($Row, $Col++) = $UCM1
            $WorkSheet.cells.Item($Row, $Col++) = $UCM2
            $Resize = $WorkSheet.UsedRange
            [Void]$Resize.EntireColumn.AutoFit()
            $Node++
            $Row++
        }Else{
            #$WorkSheet.cells.Item($Row,1) = $Target
            #$WorkSheet.cells.Item($Row,4) = "No Response"
            If ($Console){Write-host "  No Response..." -ForegroundColor Red}
            $Node++
        }
        $WorkSheet.cells.Style.HorizontalAlignment = -4131 # -4108 # -4152
    }
}

$FileName = "$PSScriptRoot\IP-Phone-Inventory_{0:MM-dd-yy_HHmmss}.xlsx" -f (Get-Date)
$Workbook.SaveAs($FileName)
If ($Console){Write-Host `n"Saving spreadsheet... " -ForegroundColor Green}
$WorkBook.Close($true)
$Excel.quit() 
[Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) #--[ Release the COM object ]--
If ($Console){Write-Host `n"--- COMPLETED ---" -ForegroundColor Red}
#>

