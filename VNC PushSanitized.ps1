$UserCredential = Get-Credential
$ServerList = @(
"*snipped*"
)
foreach ($srv in $ServerList) 
{ $srv;Invoke-Command -ComputerName $srv -Credential $UserCredential -ScriptBlock {

if ( ! (Test-Path "C:\Staging") ) {
    New-Item -Path "C:\Staging" -ItemType Directory -Force
    }
New-SmbMapping -LocalPath 'J:' -RemotePath '\\*snipped*' -UserName "*snipped*" -Password "*snipped*"
Copy-Item -LiteralPath "J:\*snipped*\UltraVNC_1_2_24_X64_Setup.exe" -Destination "C:\Staging"
Copy-Item -LiteralPath "J:\*snipped*\UltraVNCSetup.inf" -Destination "C:\Staging"
Copy-Item -LiteralPath "J:\*snipped*\ultravnc.ini" -Destination "C:\Staging"
if ( (Test-Path -Path "C:\*snipped*\UltraVNC_1_2_24_X64_Setup.exe") -and (Test-Path -Path "C:\*snipped*\UltraVNC_1_2_24_X64_Setup.exe") -and (Test-Path -Path "C:\*snipped*\UltraVNC_1_2_24_X64_Setup.exe") {
    Write-Host "UltraVNC Setup not found, please rectify and try again"
    Exit
    }
C:\*snipped*\UltraVNC_1_2_24_X64_Setup.exe /verysilent /norestart /loadinf='C:\*snipped*\UltraVNCSetup.inf'
Sleep -seconds 15
Stop-Service uvnc_service
$VNCDIRTEST = Test-Path -Path "C:\Program Files\uvnc bvba\UltraVNC\"
   if ($VNCDIRTEST -eq "False"){
   New-Item -Path "C:\Program Files\uvnc bvba\UltraVNC\" -ItemType Directory -Force
   }
   Copy-Item -Path "C:\Staging\ultravnc.ini" -Destination "C:\Program Files\uvnc bvba\UltraVNC\ultravnc.ini" -Force -Verbose
   $VNCConfirm = Test-Path -Path "C:\Program Files\uvnc bvba\UltraVNC\ultravnc.ini"
   if ($VNCConfirm -like "false") {Write-Host "Failed to apply ultravnc.ini, please try again" }
Start-Service uvnc_service
Remove-Item -LiteralPath "C:\*snipped*\UltraVNC_1_2_24_X64_Setup.exe"
Remove-Item -LiteralPath "C:\*snipped*\ultravnc.ini"
Remove-Item -LiteralPath "C:\*snipped*\UltraVNCSetup.inf"
Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 0
Enable-NetFirewallRule -DisplayGroup "Remote Desktop"
Enable-NetFirewallRule -Name "WINRM-HTTP-In-TCP-NoScope"
Set-NetFirewallRule -Name "WINRM-HTTP-In-TCP-NoScope" -RemoteAddress "*snipped*", "*snipped*"
New-NetFirewallRule -DisplayName "Allow inbound ICMPv4" -Direction Inbound -Protocol ICMPv4 -IcmpType 8 -RemoteAddress *snipped* -Action Allow
New-NetFirewallRule -DisplayName "Allow inbound ICMPv6" -Direction Inbound -Protocol ICMPv6 -IcmpType 8 -RemoteAddress *snipped* -Action Allow
$VNCFirewallRules = Get-NetFirewallRule | Where-Object -Property DisplayName -CMatch "vnc"
foreach ($Rule in $VNCFirewallRules.DisplayName) 
    { Set-NetFirewallRule -DisplayName $Rule -Profile Any -RemoteAddress "*snipped*", "*snipped*" , "*snipped*"}

$env:COMPUTERNAME
} }
