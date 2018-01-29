$Version = "1.0"

Set-Location C:\

# Step 1                                                              #
# Checks to see if scrip has already run once and skips if it has not #
$ChkFile = "C:\PS_Update" 
$FileExists = Test-Path $ChkFile 

# Step 2                             #
# If script has run once, runs 3 - 7 #
If ($FileExists -eq $True) {

pause

# Step 3                                            #
# Sets USB Drive with files as veriable "$USBDrive" #
$USBDrive = Get-WMIObject Win32_Volume | Where-Object { $_.Label -eq 'Storage' } | Select-Object -expand driveletter

# Step 4                       #
# Removed junk Windows 10 apps #
$apps = @(
    # default Windows 10 apps
    "*surfacehub*"
    "*MSPaint*"
    "*Microsoft3DViewer*"
    "*3DBuilder*"
    "*Appconnector*"
    "*BingFinance*"
    "*BingNews*"
    "*BingSports*"
    "*BingWeather*"
    "*FreshPaint*"
    "*Getstarted*"
    "*MicrosoftOfficeHub*"
    "*MicrosoftSolitaireCollection*"
    "*Office.OneNote*"
    "*People*"
    "*SkypeApp*"
    "*WindowsAlarms*"
    "*WindowsCamera*"
    "*WindowsMaps*"
    "*WindowsPhone*"
    "*WindowsSoundRecorder*"
    "*XboxApp*"
    "*ZuneMusic*"
    "*ZuneVideo*"
    "*windowscommunicationsapps*"
    "*MinecraftUWP*"
    "*MicrosoftPowerBIForWindows*"
    "*NetworkSpeedTest*"
    
    # Threshold 2 apps
    "*CommsPhone*"
    "*ConnectivityStore*"
    "*Messaging*"
    "*Office.Sway*"
    "*OneConnect*"
    "*WindowsFeedbackHub*"


    #Redstone apps
    "*BingFoodAndDrink*"
    "*BingTravel*"
    "*BingHealthAndFitness*"
    "*WindowsReadingList*"

    # non-Microsoft
    "*Twitter*"
    "*PandoraMediaInc*"
    "*Flipboard*"
    "*Shazam*"
    "*CandyCrushSaga*"
    "*CandyCrushSodaSaga*"
    "*king.com.*"
    "*iHeartRadio*"
    "*Netflix*"
    "*Wunderlist*"
    "*DrawboardPDF*"
    "*PhotoStudio*"
    "*FarmVille*"
    "*TuneInRadio*"
    "*Asphalt8Airborne*"
    "*NYTCrossword*"
    "*CyberLinkMediaSuiteEssentials*"
    "*Facebook*"
    "*RoyalRevolt*"
    "*CaesarsSlotsFreeCasino*"
    "*MarchofEmpires*"
    "*Keeper*"
    "*PhototasticCollage*"
    "*XING*"
    "*AutodeskSketchBook*"
    "*Duolingo*"
    "*Eclipse*"
    "*ActiproSoftwareLLC*" # next one is for the Code Writer from Actipro Software LLC
    "*DolbyAccess*"
    "*SpotifyMusic*"
    "*DisneyMagicKingdoms*"
    "*WinZipUniversal*"
    )

    foreach ($app in $apps) {

    Get-AppxPackage -Name $app -AllUsers | Remove-AppxPackage -AllUsers

    Get-AppxProvisionedPackage -online |
        Where-Object DisplayName -like $app |
        Remove-AppxProvisionedPackage -Online
    }

$a = new-object -ComObject wscript.shell

$a.Popup("Move PC in LabTech to Correct Location", 0, "Everything_Script_Mk_II", 0)

$a = new-object -ComObject wscript.shell

# Checks is updates installed #
$intanswer = $a.popup("Did Updates Install?", 0, "Everything_Script_Mk_II", 4)
If ($intAnswer -eq 7) {
Get-ChildItem C:\PS_Update\*.psm1 | Unblock-File
Get-ChildItem C:\PS_Update\*ps1xml | Unblock-File
set-executionpolicy unrestricted -Force
Import-Module "C:\PS_Update\PSWindowsUpdate"
Get-WUInstall -acceptall
#Set-ExecutionPolicy remoteSigned
}

$intanswer = $a.popup("Did Ninite Install?", 0, "Everything_Script_Mk_II", 4)
If ($intAnswer -eq 7) {
& "$USBDrive\Deployment Toolkit\Ninite 7Zip Air Chrome Flash Java NET 462 Reader Installer.exe" /silent
}

$intanswer = $a.popup("Did LabTech Install?", 0, "Everything_Script_Mk_II", 4)
If ($intAnswer -eq 7) {
& "$USBDrive\Deployment Toolkit\Agent_Install.exe"
}

# Step 5
# Removes folders created in steps 12 and 14 #
Remove-Item -Path "C:\PS_Update" -Recurse
Remove-Item -Path "C:\PS_RunOnce" -Recurse

# Step 6                                                         #
# Asks if you would like to install different versions of Office #
# Versions Offered as of 11/09/17 -                              #
#                   Office 2016 Pro Plus                         #
#					Office 2016 Business                         #
#					Office 2013 Standard VLSC                    #
#					Office 2013 ProPlus  VLSC                    #
#					Office 2016 Standard VLSC                    #
$a = new-object -ComObject wscript.shell

# OFFICE 2016 PROPLUS #

$intanswer = $a.popup("Install Office 2016 ProPlus?", 0, "Everything_Script_Mk_II", 4)
If ($intAnswer -eq 6) {
& "$USBDrive\Deployment Toolkit\O365ProPlus\setup.exe" /configure "$USBDrive\Deployment Toolkit\O365ProPlus\O365ProPlus.xml"
} else {

# OFFICE 2016 BUSINESS #

$intanswer = $a.popup("Install Office 2016 Business?", 0, "Everything_Script_Mk_II", 4)
If ($intAnswer -eq 6) {
& "$USBDrive\Deployment Toolkit\O365Business\setup.exe" /configure "$USBDrive\Deployment Toolkit\O365Business\O365Business.xml"
} else { 

# OFFICE 2013 STANDARD #

$intanswer = $a.popup("Install Office 2013 Standard?", 0, "Everything_Script_Mk_II", 4)
If ($intAnswer -eq 6) {
& "$USBDrive\Deployment Toolkit\Office\office2013Standard\setup.exe"
} else { 

# OFFICE 2013 PRO PLUS #

$intanswer = $a.popup("Install Office 2013 ProPlus?", 0, "Everything_Script_Mk_II", 4)
If ($intAnswer -eq 6) {
& "$USBDrive\Deployment Toolkit\Office\office2013proplus\setup.exe"
} else { 

# OFFICE 2016 STANDARD #

$intanswer = $a.popup("Install Office 2016 Standard?", 0, "Everything_Script_Mk_II", 4)
If ($intAnswer -eq 6) {
& "$USBDrive\Deployment Toolkit\Office\office2016standard\setup.exe"
} else { 

# NO OFFICE INSTALLED #
$a.Popup("No Office Installed")
}}}}}

$a = new-object -ComObject wscript.shell
$intanswer = $a.popup("Install DisplayLink?", 0, "Everything_Script_Mk_II", 4)
if ($intAnswer -eq 6) {
& "$USBDrive\Deployment Toolkit\DisplayLink_Software.exe" -stageDrivers -silent
}

$Domain = Read-Host -Prompt 'Input the Domain name'
$User = Read-Host -Prompt 'Input the user name'
$PCName = Read-Host -Prompt 'Input the Computer Name (Check Naming Scheme)'

$JoinedName = Get-WmiObject Win32_ComputerSystem | Select-Object -expand "Domain"

# Join a Domain

Add-Computer -DomainName "$Domain" -PassThru -Verbose

# Grant local admin

Add-LocalGroupMember -Group "Administrators" -Member "$User" -Confirm

#Rename the PC

Rename-Computer -NewName "$PCName" -PassThru -DomainCredential "$JoinedName\mstech"

# Step 7        #
# Closes script #
Set-ExecutionPolicy Default -Force
Exit
}

$a = new-object -ComObject wscript.shell

$a.Popup("Make Sure Internet Is Connected", 0, "Everything_Script_Mk_II", 0)

# TEST INTERNET CONNECTION #
$internet = Test-Connection google.com -Quiet
if ($internet -eq "True" ) {
    $a = new-object -ComObject wscript.shell
    $a.Popup("Internet Is Connected", 0, "Everything_Script_Mk_II", 0)
} else {
    do {
        $a = new-object -ComObject wscript.shell
        $a.Popup("Internet is not connected. Check connection and hit OK", 0, "Everything_Script_Mk_II", 0)
        $internet = Test-Connection google.com -Quiet
        } while ($internet -ne "True" )
}

# Step 8                         #
# Sets Windows 10 power settings #

# Enable Remote Desktop

(Get-WmiObject Win32_TerminalServiceSetting -Namespace root\cimv2\TerminalServices).SetAllowTsConnections(1,1) | Out-Null
(Get-WmiObject -Class "Win32_TSGeneralSetting" -Namespace root\cimv2\TerminalServices -Filter "TerminalName='RDP-tcp'").SetUserAuthenticationRequired(0) | Out-Null
 Get-NetFirewallRule -DisplayName "Remote Desktop*" | Set-NetFirewallRule -enabled true

# Set time zone to Eastern Standard Time #
Set-TimeZone "Eastern Standard Time"
Start-Service w32time
w32tm /resync
 
# Turn off NIC option "Allow this device to turn off to Save Power" 

$file = "C:\NICpowerChange.log"
"Searching for Dell / Intel" | Add-Content -Path $file

#find relevant network adapters
$nics = Get-WmiObject Win32_NetworkAdapter | Where-Object {$_.Name.Contains('Dell') -or $_.Name.Contains('Intel')}

$nicsFound = $nics.Count
"number of network adapters found: ", $nicsFound | Add-Content -Path $file

foreach ($nic in $nics)
{
   $powerMgmt = Get-WmiObject MSPower_DeviceEnable -Namespace root\wmi | Where-Object {$_.InstanceName -match [regex]::Escape($nic.PNPDeviceID)}
 
   # check to see if power management can be turned off
   if(Get-Member -inputobject $powerMgmt -name "Enable" -Membertype Properties){

     # check if powermanagement is enabled
     if ($powerMgmt.Enable -eq $True){
       $nic.Name, "----- Enabled method exists. PowerSaving is set to enabled, disabling..." | Add-Content -Path $file
       $powerMgmt.Enable = $False
       $powerMgmt.psbase.Put()
     }
     else
     {
       $nic.Name, "----- Enabled method exists. PowerSaving is already set to disabled. skipping..." | Add-Content -Path $file
     }
   }
   else
   {
     $nic.Name, "----- Enabled method doesnt exist, so PowerSaving cannot be set" | Add-Content -Path $file 
   }
}

# Windows 10 Power Configuration.

# Step 9                                            #
# Sets USB Drive with files as veriable "$USBDrive" #
$USBDrive = Get-WMIObject Win32_Volume | Where-Object { $_.Label -eq 'Storage' } | Select-Object -expand driveletter

Write-Output Now configuring your Power Plan.
$GUIDNew = 'b23303fa-44e5-48f8-a2cf-358c11d6d4f1'
powercfg /import "$USBDrive\Deployment Toolkit\PowerPlan\Mainstay.pow" $GUIDNew
powercfg -setactive $GUIDNew
Write-Output Your Power Plan has been configured.

# Step 10                                                                                       #
# Installs Ninite Programs (7Zip, Air, Chrome, Flash, Java, NET 4.62, Adobe Reader DC) silently #

$a = new-object -ComObject wscript.shell
$intanswer = $a.popup("Install Ninite?", 0, "Everything_Script_Mk_II", 4)
if ($intAnswer -eq 6) {
& "$USBDrive\Deployment Toolkit\Ninite 7Zip Air Chrome Flash Java NET 462 Reader Installer.exe" /silent
}

# Step 11                                 #
# Removes preinstalled versions of Office #

$a = new-object -ComObject wscript.shell
$intanswer = $a.popup("Remove all versions of Office?", 0, "Everything_Script_Mk_II", 4)
if ($intAnswer -eq 6) {
& "$USBDrive\Deployment Toolkit\Junk Removal\OffScrub_O15msi.vbs"
& "$USBDrive\Deployment Toolkit\Junk Removal\OffScrub_O16msi.vbs"
& "$USBDrive\Deployment Toolkit\Junk Removal\OffScrubc2r.vbs"
}

# Step 12                                                            #
# Moves PowerShell module to install Windows Updates to C:\PS_Update #
$a = new-object -ComObject wscript.shell
$intanswer = $a.popup("Install Updates?", 0, "Everything_Script_Mk_II", 4)
if ($intAnswer -eq 6) {
Copy-Item -path "$USBDrive\Deployment Toolkit\PSWindowsUpdate" -Destination "C:\PS_Update" -Recurse


# Step 13                                                                     #
# Unblocks Update module, adds module to powershell, and runs Windows Updates #
Get-ChildItem C:\PS_Update\*.psm1 | Unblock-File
Get-ChildItem C:\PS_Update\*ps1xml | Unblock-File
set-executionpolicy unrestricted -Force
Import-Module "C:\PS_Update\PSWindowsUpdate"
Get-WUInstall -acceptall -IgnoreReboot
}

Do {
$return = Get-WUInstallerStatus -silent
Get-WUInstallerStatus
Start-Sleep -Seconds 60
} while ($return -eq "False")

# Wait for Ninite to finish #
Wait-Process Ninite

# Step 14                            #
# Moves prep script to C:\PS_RunOnce #
new-Item c:\PS_RunOnce -ItemType Directory
Copy-Item -Path "$USBDrive\Everything_Script_Mk_II.ps1" -Destination "C:\PS_RunOnce"
Copy-Item -Path "$USBDrive\RunOnce.ps1" -Destination "C:\PS_RunOnce"
Get-ChildItem C:\PS_RunOnce\*.ps1 | Unblock-File

# Step 15                                                       #
# Creates RunOnce registry key to run script after rebooting PC #
$regCheck = Test-Path HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce
If ($regCheck -eq $True) {
Set-Location -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce'
Set-ItemProperty -Path . -Name removePrograms -Value 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -File C:\PS_RunOnce\RunOnce.ps1'
} else {
Set-Location -Path HKCU:\Software\Microsoft\Windows\CurrentVersion
New-Item -Path . -Name RunOnce
Set-ItemProperty -Path .\RunOnce -Name removePrograms -Value 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -File C:\PS_RunOnce\RunOnce.ps1'
}

# Step 16                       #
# Install default LabTech agent #
& "$USBDrive\Deployment Toolkit\Agent_Install.exe"

Pause

# Step 17
# Reboots PC #
Restart-Computer