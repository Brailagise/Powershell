#Requires -RunAsAdministrator
Add-Type -AssemblyName PresentationCore, PresentationFramework

#region Functions (LONG)
Function DebloatBlacklist {
    [CmdletBinding()]

    Param ()

    $Bloatware = @(

        #Unnecessary Windows 10 AppX Apps
        "7ee7776c.LinkedInforWindows"
        "DellInc.DellCommandUpdate"
        "DellInc.DellDigitalDelivery"
        "DellInc.DellSupportAssistantforPCs"
        "Microsoft.BingNews"
        "Microsoft.DesktopAppInstaller"
        "Microsoft.GetHelp"
        "Microsoft.Getstarted"
        "Microsoft.Messaging"
        "Microsoft.Microsoft3DViewer"
        "Microsoft.MicrosoftOfficeHub"
        "Microsoft.MicrosoftSolitaireCollection"
        "Microsoft.Wallet"
        "Microsoft.NetworkSpeedTest"
        "Microsoft.Office.OneNote"
        "Microsoft.Office.Sway"
        "Microsoft.OneConnect"
        "Microsoft.People"
        "Microsoft.Print3D"
        "Microsoft.SkypeApp"
        "Microsoft.StorePurchaseApp"
        "Microsoft.WindowsAlarms"
        "Microsoft.WindowsCamera"
        "microsoft.windowscommunicationsapps"
        "Microsoft.WindowsFeedbackHub"
        "Microsoft.WindowsMaps"
        "Microsoft.WindowsSoundRecorder"
        "Microsoft.XboxApp"
        "Microsoft.XboxGameOverlay"
        "Microsoft.XboxIdentityProvider"
        "Microsoft.XboxSpeechToTextOverlay"
        "Microsoft.ZuneMusic"
        "Microsoft.ZuneVideo"
             
        #Sponsored Windows 10 AppX Apps
        #Add sponsored/featured apps to remove in the "*AppName*" format
        "*EclipseManager*"
        "*ActiproSoftwareLLC*"
        "*AdobeSystemsIncorporated.AdobePhotoshopExpress*"
        "*Duolingo-LearnLanguagesforFree*"
        "*PandoraMediaInc*"
        "*CandyCrush*"
        "*Wunderlist*"
        "*Flipboard*"
        "*Twitter*"
        "*Facebook*"
        "*Spotify*"
        "*Minecraft*"
        "*Royal Revolt*"
             
        #Optional: Typically not removed but you can if you need to for some reason
        #"*Microsoft.Advertising.Xaml_10.1712.5.0_x64__8wekyb3d8bbwe*"
        #"*Microsoft.Advertising.Xaml_10.1712.5.0_x86__8wekyb3d8bbwe*"
        "*Microsoft.BingWeather*"
        "*Microsoft.MSPaint*"
        #"*Microsoft.MicrosoftStickyNotes*"
        #"*Microsoft.Windows.Photos*"
        #"*Microsoft.WindowsCalculator*"
        #"*Microsoft.WindowsStore*"
    )
    foreach ($Bloat in $Bloatware) {
        Get-AppxPackage -Name $Bloat| Remove-AppxPackage -ErrorAction SilentlyContinue
        Get-AppxProvisionedPackage -Online | Where-Object DisplayName -like $Debloat | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
        Write-Output "Trying to remove $Bloat."
    }
}

Function Remove-Keys {
        
    [CmdletBinding()]
            
    Param()
        
    #These are the registry keys that it will delete.
            
    $Keys = @(
            
        #Remove Background Tasks
        "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\46928bounde.EclipseManager_2.2.4.51_neutral__a5h4egax66k6y"
        "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\ActiproSoftwareLLC.562882FEEB491_2.6.18.18_neutral__24pqs290vpjk0"
        "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\Microsoft.MicrosoftOfficeHub_17.7909.7600.0_x64__8wekyb3d8bbwe"
        "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\Microsoft.PPIProjection_10.0.15063.0_neutral_neutral_cw5n1h2txyewy"
        "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\Microsoft.XboxGameCallableUI_1000.15063.0.0_neutral_neutral_cw5n1h2txyewy"
        "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\Microsoft.XboxGameCallableUI_1000.16299.15.0_neutral_neutral_cw5n1h2txyewy"
            
        #Windows File
        "HKCR:\Extensions\ContractId\Windows.File\PackageId\ActiproSoftwareLLC.562882FEEB491_2.6.18.18_neutral__24pqs290vpjk0"
            
        #Registry keys to delete if they aren't uninstalled by RemoveAppXPackage/RemoveAppXProvisionedPackage
        "HKCR:\Extensions\ContractId\Windows.Launch\PackageId\46928bounde.EclipseManager_2.2.4.51_neutral__a5h4egax66k6y"
        "HKCR:\Extensions\ContractId\Windows.Launch\PackageId\ActiproSoftwareLLC.562882FEEB491_2.6.18.18_neutral__24pqs290vpjk0"
        "HKCR:\Extensions\ContractId\Windows.Launch\PackageId\Microsoft.PPIProjection_10.0.15063.0_neutral_neutral_cw5n1h2txyewy"
        "HKCR:\Extensions\ContractId\Windows.Launch\PackageId\Microsoft.XboxGameCallableUI_1000.15063.0.0_neutral_neutral_cw5n1h2txyewy"
        "HKCR:\Extensions\ContractId\Windows.Launch\PackageId\Microsoft.XboxGameCallableUI_1000.16299.15.0_neutral_neutral_cw5n1h2txyewy"
            
        #Scheduled Tasks to delete
        "HKCR:\Extensions\ContractId\Windows.PreInstalledConfigTask\PackageId\Microsoft.MicrosoftOfficeHub_17.7909.7600.0_x64__8wekyb3d8bbwe"
            
        #Windows Protocol Keys
        "HKCR:\Extensions\ContractId\Windows.Protocol\PackageId\ActiproSoftwareLLC.562882FEEB491_2.6.18.18_neutral__24pqs290vpjk0"
        "HKCR:\Extensions\ContractId\Windows.Protocol\PackageId\Microsoft.PPIProjection_10.0.15063.0_neutral_neutral_cw5n1h2txyewy"
        "HKCR:\Extensions\ContractId\Windows.Protocol\PackageId\Microsoft.XboxGameCallableUI_1000.15063.0.0_neutral_neutral_cw5n1h2txyewy"
        "HKCR:\Extensions\ContractId\Windows.Protocol\PackageId\Microsoft.XboxGameCallableUI_1000.16299.15.0_neutral_neutral_cw5n1h2txyewy"
               
        #Windows Share Target
        "HKCR:\Extensions\ContractId\Windows.ShareTarget\PackageId\ActiproSoftwareLLC.562882FEEB491_2.6.18.18_neutral__24pqs290vpjk0"
    )
        
    #This writes the output of each key it is removing and also removes the keys listed above.
    ForEach ($Key in $Keys) {
        Write-Output "Removing $Key from registry"
        Remove-Item $Key -Recurse
    }
}
Function Protect-Privacy {
        
    [CmdletBinding()]
        
    Param()
            
    #Disables Windows Feedback Experience
    Write-Output "Disabling Windows Feedback Experience program"
    $Advertising = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AdvertisingInfo"
    If (Test-Path $Advertising) {
        Set-ItemProperty $Advertising Enabled -Value 0 
    }
            
    #Stops Cortana from being used as part of your Windows Search Function
    Write-Output "Stopping Cortana from being used as part of your Windows Search Function"
    $Search = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
    If (Test-Path $Search) {
        Set-ItemProperty $Search AllowCortana -Value 0 
    }

    #Disables Web Search in Start Menu
    Write-Output "Disabling Bing Search in Start Menu"
    $WebSearch = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
    Set-ItemProperty "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" BingSearchEnabled -Value 0 
    If (!(Test-Path $WebSearch)) {
        New-Item $WebSearch
    }
    Set-ItemProperty $WebSearch DisableWebSearch -Value 1 
            
    #Stops the Windows Feedback Experience from sending anonymous data
    Write-Output "Stopping the Windows Feedback Experience program"
    $Period = "HKCU:\Software\Microsoft\Siuf\Rules"
    If (!(Test-Path $Period)) { 
        New-Item $Period
    }
    Set-ItemProperty $Period PeriodInNanoSeconds -Value 0 

    #Prevents bloatware applications from returning and removes Start Menu suggestions               
    Write-Output "Adding Registry key to prevent bloatware apps from returning"
    $registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
    $registryOEM = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
    If (!(Test-Path $registryPath)) { 
        New-Item $registryPath
    }
    Set-ItemProperty $registryPath DisableWindowsConsumerFeatures -Value 1 

    If (!(Test-Path $registryOEM)) {
        New-Item $registryOEM
    }
    Set-ItemProperty $registryOEM  ContentDeliveryAllowed -Value 0 
    Set-ItemProperty $registryOEM  OemPreInstalledAppsEnabled -Value 0 
    Set-ItemProperty $registryOEM  PreInstalledAppsEnabled -Value 0 
    Set-ItemProperty $registryOEM  PreInstalledAppsEverEnabled -Value 0 
    Set-ItemProperty $registryOEM  SilentInstalledAppsEnabled -Value 0 
    Set-ItemProperty $registryOEM  SystemPaneSuggestionsEnabled -Value 0          
    
    #Preping mixed Reality Portal for removal    
    Write-Output "Setting Mixed Reality Portal value to 0 so that you can uninstall it in Settings"
    $Holo = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Holographic"    
    If (Test-Path $Holo) {
        Set-ItemProperty $Holo  FirstRunSucceeded -Value 0 
    }

    #Disables Wi-fi Sense
    Write-Output "Disabling Wi-Fi Sense"
    $WifiSense1 = "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\WiFi\AllowWiFiHotSpotReporting"
    $WifiSense2 = "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\WiFi\AllowAutoConnectToWiFiSenseHotspots"
    $WifiSense3 = "HKLM:\SOFTWARE\Microsoft\WcmSvc\wifinetworkmanager\config"
    If (!(Test-Path $WifiSense1)) {
        New-Item $WifiSense1
    }
    Set-ItemProperty $WifiSense1  Value -Value 0 
    If (!(Test-Path $WifiSense2)) {
        New-Item $WifiSense2
    }
    Set-ItemProperty $WifiSense2  Value -Value 0 
    Set-ItemProperty $WifiSense3  AutoConnectAllowedOEM -Value 0 
        
    #Disables live tiles
    Write-Output "Disabling live tiles"
    $Live = "HKCU:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\PushNotifications"    
    If (!(Test-Path $Live)) {      
        New-Item $Live
    }
    Set-ItemProperty $Live  NoTileApplicationNotification -Value 1 
        
    #Turns off Data Collection via the AllowTelemtry key by changing it to 0
    Write-Output "Turning off Data Collection"
    $DataCollection1 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection"
    $DataCollection2 = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"
    $DataCollection3 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Policies\DataCollection"    
    If (Test-Path $DataCollection1) {
        Set-ItemProperty $DataCollection1  AllowTelemetry -Value 0 
    }
    If (Test-Path $DataCollection2) {
        Set-ItemProperty $DataCollection2  AllowTelemetry -Value 0 
    }
    If (Test-Path $DataCollection3) {
        Set-ItemProperty $DataCollection3  AllowTelemetry -Value 0 
    }
    
    #Disabling Location Tracking
    Write-Output "Disabling Location Tracking"
    $SensorState = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Sensor\Overrides\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}"
    $LocationConfig = "HKLM:\SYSTEM\CurrentControlSet\Services\lfsvc\Service\Configuration"
    If (!(Test-Path $SensorState)) {
        New-Item $SensorState
    }
    Set-ItemProperty $SensorState SensorPermissionState -Value 0 
    If (!(Test-Path $LocationConfig)) {
        New-Item $LocationConfig
    }
    Set-ItemProperty $LocationConfig Status -Value 0 
        
    #Disables People icon on Taskbar
    Write-Output "Disabling People icon on Taskbar"
    $People = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\People"    
    If (!(Test-Path $People)) {
        New-Item $People
    }
    Set-ItemProperty $People  PeopleBand -Value 0 
        
    #Disables scheduled tasks that are considered unnecessary 
    Write-Output "Disabling scheduled tasks"
    Get-ScheduledTask  XblGameSaveTaskLogon | Disable-ScheduledTask
    Get-ScheduledTask  XblGameSaveTask | Disable-ScheduledTask
    Get-ScheduledTask  Consolidator | Disable-ScheduledTask
    Get-ScheduledTask  UsbCeip | Disable-ScheduledTask
    Get-ScheduledTask  DmClient | Disable-ScheduledTask
    Get-ScheduledTask  DmClientOnScenarioDownload | Disable-ScheduledTask

    Write-Output "Stopping and disabling Diagnostics Tracking Service"
    #Disabling the Diagnostics Tracking Service
    Stop-Service "DiagTrack"
    Set-Service "DiagTrack" -StartupType Disabled

    
     Write-Output "Removing CloudStore from registry if it exists"
     $CloudStore = 'HKCUSoftware\Microsoft\Windows\CurrentVersion\CloudStore'
     If (Test-Path $CloudStore) {
     Stop-Process Explorer.exe -Force
     Remove-Item $CloudStore
     Start-Process Explorer.exe -Wait
   }
}

Function DisableCortana {
    Write-Host "Disabling Cortana"
    $Cortana1 = "HKCU:\SOFTWARE\Microsoft\Personalization\Settings"
    $Cortana2 = "HKCU:\SOFTWARE\Microsoft\InputPersonalization"
    $Cortana3 = "HKCU:\SOFTWARE\Microsoft\InputPersonalization\TrainedDataStore"
    If (!(Test-Path $Cortana1)) {
        New-Item $Cortana1
    }
    Set-ItemProperty $Cortana1 AcceptedPrivacyPolicy -Value 0 
    If (!(Test-Path $Cortana2)) {
        New-Item $Cortana2
    }
    Set-ItemProperty $Cortana2 RestrictImplicitTextCollection -Value 1 
    Set-ItemProperty $Cortana2 RestrictImplicitInkCollection -Value 1 
    If (!(Test-Path $Cortana3)) {
        New-Item $Cortana3
    }
    Set-ItemProperty $Cortana3 HarvestContacts -Value 0
    
}
Function UninstallOneDrive {

    Write-Output "Checking for pre-existing files and folders located in the OneDrive folders..."
    Start-Sleep 1
    If (Get-Item -Path "$env:USERPROFILE\OneDrive\*") {
        Write-Output "Files found within the OneDrive folder! Checking to see if a folder named OneDriveBackupFiles exists."
        Start-Sleep 1
              
        If (Get-Item "$env:USERPROFILE\Desktop\OneDriveBackupFiles" -ErrorAction SilentlyContinue) {
            Write-Output "A folder named OneDriveBackupFiles already exists on your desktop. All files from your OneDrive location will be moved to that folder." 
        }
        else {
            If (!(Get-Item "$env:USERPROFILE\Desktop\OneDriveBackupFiles" -ErrorAction SilentlyContinue)) {
                Write-Output "A folder named OneDriveBackupFiles will be created and will be located on your desktop. All files from your OneDrive location will be located in that folder."
                New-item -Path "$env:USERPROFILE\Desktop" -Name "OneDriveBackupFiles"-ItemType Directory -Force
                Write-Output "Successfully created the folder 'OneDriveBackupFiles' on your desktop."
            }
        }
        Start-Sleep 1
        Move-Item -Path "$env:USERPROFILE\OneDrive\*" -Destination "$env:USERPROFILE\Desktop\OneDriveBackupFiles" -Force
        Write-Output "Successfully moved all files/folders from your OneDrive folder to the folder 'OneDriveBackupFiles' on your desktop."
        Start-Sleep 1
        Write-Output "Proceeding with the removal of OneDrive."
        Start-Sleep 1
    }
    Else {
        If (!(Get-Item -Path "$env:USERPROFILE\OneDrive\*")) {
            Write-Output "Either the OneDrive folder does not exist or there are no files to be found in the folder. Proceeding with removal of OneDrive."
            Start-Sleep 1
        }
    }

    Write-Output "Uninstalling OneDrive"
    
    New-PSDrive  HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT
    $onedrive = "$env:SYSTEMROOT\SysWOW64\OneDriveSetup.exe"
    $ExplorerReg1 = "HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
    $ExplorerReg2 = "HKCR:\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
    Stop-Process -Name "OneDrive*"
    Start-Sleep 2
    If (!(Test-Path $onedrive)) {
        $onedrive = "$env:SYSTEMROOT\System32\OneDriveSetup.exe"
    }
    Start-Process $onedrive "/uninstall" -NoNewWindow -Wait
    Start-Sleep 2
    Write-Output "Stopping explorer"
    Start-Sleep 1
    .\taskkill.exe /F /IM explorer.exe
    Start-Sleep 3
    Write-Output "Removing leftover files"
    Remove-Item "$env:USERPROFILE\OneDrive" -Force -Recurse
    Remove-Item "$env:LOCALAPPDATA\Microsoft\OneDrive" -Force -Recurse
    Remove-Item "$env:PROGRAMDATA\Microsoft OneDrive" -Force -Recurse
    If (Test-Path "$env:SYSTEMDRIVE\OneDriveTemp") {
        Remove-Item "$env:SYSTEMDRIVE\OneDriveTemp" -Force -Recurse
    }
    Write-Output "Removing OneDrive from windows explorer"
    If (!(Test-Path $ExplorerReg1)) {
        New-Item $ExplorerReg1
    }
    Set-ItemProperty $ExplorerReg1 System.IsPinnedToNameSpaceTree -Value 0 
    If (!(Test-Path $ExplorerReg2)) {
        New-Item $ExplorerReg2
    }
    Set-ItemProperty $ExplorerReg2 System.IsPinnedToNameSpaceTree -Value 0
    Write-Output "Restarting Explorer that was shut down before."
    Start-Process explorer.exe -NoNewWindow
    
    Write-Host "Enabling the Group Policy 'Prevent the usage of OneDrive for File Storage'."
        $OneDriveKey = 'HKLM:Software\Policies\Microsoft\Windows\OneDrive'
        If (!(Test-Path $OneDriveKey)) {
            Mkdir $OneDriveKey 
        }

        $DisableAllOneDrive = 'HKLM:Software\Policies\Microsoft\Windows\OneDrive'
        If (Test-Path $DisableAllOneDrive) {
            New-ItemProperty $DisableAllOneDrive -Name OneDrive -Value DisableFileSyncNGSC -Verbose 
        }
}
#endregion


$TempFolder = "C:\Temp\"
If (Test-Path $TempFolder) {
    Write-Output "$TempFolder exists. Skipping."
}
Else {
    Write-Output "The folder "$TempFolder" doesn't exist. This folder will be used for storing logs created after the script runs. Creating now."
    Start-Sleep 1
    New-Item -Path "$TempFolder" -ItemType Directory -Force
    Write-Output "The folder $TempFolder was successfully created."
}

Start-Transcript -OutputDirectory "$TempFolder"

Set-TimeZone -id "Eastern Standard Time"

Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser
$AdminPassword = ConvertTo-SecureString -AsPlainText "#Snipped#" -Force
Enable-LocalUser -Name Administrator
Set-LocalUser -Name Administrator -Password $AdminPassword
$NewComputerName = Read-Host -Prompt "Please enter new computer name"
Rename-Computer -NewName $NewComputerName -Confirm:$false -PassThru



#Choices edited#
Do {$NetworkChoice = Read-Host -Prompt "Please choose a network
1. (Office)
2. (Production / No Internet / Debloats ALL non-essential apps)"} while (1,2 -notcontains $NetworkChoice)

if ($NetworkChoice -eq 1) {

Copy-Item -Path "D:\Wi-Fi 2-Qcells.xml" -Destination "$TempFolder\Wi-Fi 2-Qcells.xml"
netsh wlan add profile filename="$TempFolder\Wi-Fi 2-Qcells.xml"
Sleep -Seconds 10
$WirelessAdapter = Get-NetAdapter -Name *Wi-Fi*
$WirelessConfig = Get-NetIPConfiguration -InterfaceIndex $WirelessAdapter.ifIndex
$WirelessDNSConfig = $WirelessAdapter | Get-DnsClientServerAddress -AddressFamily IPv4
Set-NetConnectionProfile -InterfaceIndex $WirelessAdapter.ifIndex -NetworkCategory Private
Remove-Item -Path "$TempFolder\Wi-Fi 2-Qcells.xml" -Force
$ComputerMac = $WirelessAdapter.MacAddress
$Time = Get-Date
$FileTime = Get-Date -format MMdd
$FilePathLog = "D:\$FileTime Log.txt"
$UserLogText = Write-Output "
Time: $Time
Computer Name: $NewComputerName
IPV4 Address: $WirelessConfig.IPV4Address
MAC Address: $ComputerMac
"
$UserLogText | Out-File -Append -FilePath $FilePathLog

}
elseif ($NetworkChoice -eq 2) {

Write-Host "Starting APP debloat"
DebloatBlacklist
Remove-Keys
Protect-Privacy
DisableCortana
UninstallOneDrive

Write-Host "User Configuration"
$NewUserPrompt = Read-Host -Prompt "Create local user with admin? admin rights can be removed at end
1. Yes
2. No"
if ($NewUserPrompt -eq 1) 
{
$NewUserName = Read-Host -Prompt "Enter Account Name"
$NewUserPassword = Read-Host -Prompt "Enter password"
if ($NewUserPassword -eq "")
{
New-LocalUser $NewUserName -NoPassword -FullName $NewUserName -Description "Description of this account."
Add-LocalGroupMember -Group "Administrators" -Member "$NewUserName"
$AutoLogonChoice = Read-Host -Prompt "Do you want to set this account to auto-logon or have admin??
1. Yes
2. No"
if ($AutoLogonChoice -eq 1)
    {
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
Set-ItemProperty $RegPath "AutoAdminLogon" -Value "1" -type String 
Set-ItemProperty $RegPath "DefaultUsername" -Value "$NewUserName" -type String 
Set-ItemProperty $RegPath "DefaultPassword" -Value "" -type String
    }
}
else {
$NewUserPasswordSecure = ConvertTo-SecureString $NewUserPassword -AsPlainText -Force
New-LocalUser $NewUserName -Password $NewUserPasswordSecure -FullName $NewUserName -Description "Description of this account."
Add-LocalGroupMember -Group "Administrators" -Member "$NewUserName"
$AutoLogonChoice = Read-Host -Prompt "Do you want to set this account to auto-logon?
1. Yes
2. No"
if ($AutoLogonChoice -eq 1)
    {
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
Set-ItemProperty $RegPath "AutoAdminLogon" -Value "1" -type String 
Set-ItemProperty $RegPath "DefaultUsername" -Value "$NewUserName" -type String 
Set-ItemProperty $RegPath "DefaultPassword" -Value "$NewUserPassword" -type String
    }
}
}

Write-Host "Starting Network Setup"
$IPV4Set = Read-Host -Prompt "Please enter static IPV4 assignment"
Copy-Item -Path "D:\Wi-Fi-Qcells_OIPC.xml" -Destination "$TempFolder\Wi-Fi-Qcells_OIPC.xml"
netsh wlan delete profile name="#Snipped#"
netsh wlan add profile filename="C:\Temp\Wi-Fi-Qcells_OIPC.xml"
Sleep -Seconds 60
$WirelessAdapter = Get-NetAdapter -Name *Wi-Fi*
$WirelessConfig = Get-NetIPConfiguration -InterfaceIndex $WirelessAdapter.ifIndex
New-NetIPAddress -InterfaceIndex $WirelessAdapter.ifIndex -AddressFamily IPv4 -IPAddress $IPV4Set -PrefixLength "24" -DefaultGateway #Snipped#
Set-NetIPAddress -InterfaceIndex $WirelessAdapter.ifIndex -AddressFamily IPv4 -IPAddress $IPV4Set -PrefixLength "24"
Set-DnsClientServerAddress -InterfaceIndex $WirelessAdapter.ifIndex -ServerAddresses #Snipped#
Set-NetConnectionProfile -InterfaceIndex $WirelessAdapter.ifIndex -NetworkCategory Private
Remove-Item -Path "$TempFolder\#Snipped#.xml" -Force
$ComputerMac = $WirelessAdapter.MacAddress
$Time = Get-Date
$FileTime = Get-Date -format MMdd
$FilePathLog = "D:\$FileTime Log.txt"
$UserLogText = Write-Output "
Time: $Time
Computer Name: $NewComputerName
IPV4 Address: $IPV4Set
MAC Address: $ComputerMac
"
$UserLogText | Out-File -Append -FilePath $FilePathLog

}

#Office Check#
$InstalledApps = Get-AppxProvisionedPackage -Online | Select-Object DisplayName | Select-String -Pattern "Microsoft.Office.Desktop" | ft
if ($InstalledApps -eq "") {
Write-Host "Installation of Office not Found"
Write-Host "Starting Office Install"
$OfficeRunning = Start-Process -FilePath "D:\Setup.exe" -ArgumentList "/configure OfficeConfiguration.xml" -PassThru
}
else
{
Write-Host "Installation of office found! customization starting."
$OfficeRunning = Start-Process -FilePath "D:\Setup.exe" -ArgumentList "/customize OfficeConfiguration.xml" -PassThru
}
$OfficeRunning | Wait-Process
Do{
$SoftwareInstall = Read-Host -Prompt "Select special software installation
1. MESOI
2. SAP
3. Done"
Switch ($SoftwareInstall) {

1 {
   Write-Host "Starting MESOI installation"
   $MESOIInstall = Start-Process -FilePath "D:\MESOIClient\setup.exe" -PassThru
   $MESOIInstall | Wait-Process
   Write-Host "Installation complete, don't forget to configure.
Server Address: #Snipped#
Site: #Snipped#
Station Mode: #Snipped#
Factory: #Snipped#
Theme: Dark
Press any key to exit..."
    $InstallationBreak = Read-Host -prompt "More special software to install?
    1. Yes
    2. No"
    if ($InstallationBreak -eq "2") {
    $InstallationBreak = ""
    $SoftwareInstall = 3}
} #close choice 1 MESOI Install

2{

   Write-Host "Starting first part of SAP Install"
   $SAPInstall = Start-Process -FilePath "D:\SAPGUI\01.SAPGUI740\PRES1\GUI\WINDOWS\WIN32\SetupAll.exe" -PassThru
   $SAPInstall | Wait-Process
   Write-Host "Install Complete, starting patch"
   $SAPInstall2 = Start-Process -FilePath "D:\SAPGUI\02.gui740_9-10013011.exe" -PassThru
   $SAPInstall2 | Wait-Process
   Write-Host "Attempting to input registration file"
   Sleep -Seconds 2
   Copy-Item -Path "D:\SAPGUI\saplogon.ini" -Destination "$env:APPDATA\SAP\Common\saplogon.ini" -Force
   $INIConfirm = Test-Path -Path "$env:APPDATA\SAP\Common\saplogon.ini"
   if ($INIConfirm -like "false") {Write-Host "Failed to apply saplogon.ini, please try again"}
   $InstallationBreak = Read-Host -prompt "Do you want to install any additional special software?
1. Yes
2. No"
    if ($InstallationBreak -eq "2") {
    $InstallationBreak = ""
    $SoftwareInstall = 3}
 
} #close choice 2 SAP Install

} #close switch

} until ($SoftwareInstall -eq 3) #close repeat

Write-Host "Please change any required rights"
$PasswordChangeWindow = Start-Process netplwiz
$PasswordChangeWindow | Wait-Process

$RebootAsk = Read-Host -Prompt "A reboot is recommended, proceed?
1. Yes
2. No"
if ($RebootAsk -eq "1") {
Write-Host "Restarting Computer, Complete"
Restart-Computer}
Write-Host "Ending task"


