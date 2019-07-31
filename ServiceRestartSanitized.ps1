if ( ! (Test-Path "C:\Logging") ) {

    New-Item "C:\Logging" -TypeName Folder

    }

Start-Transcript -Path ( "C:\Logging\" + (Get-Date -Format "MM.dd.yyyy.HH.mm") + "CPUAlertLog.txt" ) -IncludeInvocationHeader -Verbose

$time = Try{

$request = $null
$ReqTime = Measure-Command { $request = Invoke-WebRequest -Uri http://127.0.0.1/*snipped* }
$ReqTime.TotalMilliseconds  
  
}Catch

{

    $request = $_.Exception
    $time = -1

}

if ($request.StatusCode -eq 200) {

    Write-Host "Heartbeat worked! Exiting automatic restart sequence"
    $request | Out-Host
    Write-Host "Time for request"
    $time | Out-Host
    Stop-Transcript
    Exit

}

Write-Host "Heartbeat did not return a 200! Outputting response and response timing"
$response | Out-Host
if ($time -eq -1) {
Write-Host "TIMEDOUT!"
}
else {
Write-Host "$time"
}


Write-Host "SERVICE TRANSCRIPTION"
Get-Service | Sort-Object -Property Status -Descending

Write-Host "PROCESS TRANSCRIPTION"
Get-Process | Sort-Object -Property CPU -Descending


Get-Service | Where-Object -Property Name -In -Value ( "*snipped*" , "*snipped*" , "*snipped*" ) | Select * -OutVariable Services
ForEach ($Service in $Services) {
    if ($Service.Status -eq "Stopped") {
        Write-Host "$Service.Name is stopped before the restart! This may indicate a previous crash"
    }
}

Write-Host "STOPPING PROCESSES"
Stop-Service $Services.Name -Verbose

$*snipped* = Get-Service "*snipped*"
$*snipped* = Get-Service "*snipped*"
$*snipped* = Get-Service "*snipped*"

Write-Host "Waiting for all processes to stop"
$*snipped*.WaitForStatus('Stopped')
$*snipped* | Out-Host

$*snipped*.WaitForStatus('Stopped')
$*snipped* | Out-Host

$*snipped*.WaitForStatus('Stopped')
$*snipped* | Out-Host

Write-Host "All processes should be Stopped"
Get-Service | Where-Object -Property Name -In -Value ( "*snipped*" , "*snipped*" , "*snipped*" )

Write-Host "Starting *snipped*"
Start-Service "*snipped*"
$*snipped*.WaitForStatus('Running')
$*snipped* | Out-Host

Write-Host "Sleeping for 10 seconds"
Sleep -Seconds 10

Write-Host "Starting *snipped*"
Start-Service "*snipped*"
$*snipped*.WaitForStatus('Running')
$*snipped* | Out-Host

Write-Host "Sleeping for 10 seconds"
Sleep -Seconds 10

Write-Host "Starting *snipped*"
Start-Service "*snipped*"
$*snipped*.WaitForStatus('Running')
$*snipped* | Out-Host

Stop-Transcript