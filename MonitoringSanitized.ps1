#region Functions

#Intakes all existing variables so that they can be cleared later in the script
$ExistingVariables = Get-Variable | Select-Object -ExpandProperty Name

# Sets up conversion function for use later on in converting sizes
function Convert-Size {
    [cmdletbinding()]
    param(
        [validateset("Bytes", "KB", "MB", "GB", "TB")]
        [string]$From,
        [validateset("Bytes", "KB", "MB", "GB", "TB")]
        [string]$To,
        [Parameter(Mandatory = $true)]
        [double]$Value,
        [int]$Precision = 4
    )
    switch ($From) {
        "Bytes" { $value = $Value }
        "KB" { $value = $Value * 1024 }
        "MB" { $value = $Value * 1024 * 1024 }
        "GB" { $value = $Value * 1024 * 1024 * 1024 }
        "TB" { $value = $Value * 1024 * 1024 * 1024 * 1024 }
    }

    switch ($To) {
        "Bytes" { return $value }
        "KB" { $Value = $Value / 1KB }
        "MB" { $Value = $Value / 1MB }
        "GB" { $Value = $Value / 1GB }
        "TB" { $Value = $Value / 1TB }

    }

    return [Math]::Round($value, $Precision, [MidPointRounding]::AwayFromZero)

}


#endregion

#region Import/Install Packages/Assembly"


if ( ! ( Import-PackageProvider -Name NuGet -Force | Out-Null ) ) { 
    Install-PackageProvider -Name NuGet -Scope CurrentUser -Force
    Write-Host "Installing NuGet"
    }

$NeededModules = @(
    'Unity-Powershell'
    'VMware.PowerCLI'
    'iDRAC4redfish'
    'ImportExcel'
)

foreach ($Module in $NeededModules) {
    if ( ! ( Get-Module -ListAvailable -Name $Module -Force | Out-Null ) ) {
        Write-Host "Installing $Module"
        Install-Module -Name $Module -Scope CurrentUser -Force
    }
    else {
        Import-Module $Module
    }
}

Import-Module "$env:HOMEDRIVE\Program Files\WindowsPowerShell\Scripts\Test-ConnectionAsync.ps1"
Import-Module "$env:HOMEDRIVE\Program Files\WindowsPowerShell\Scripts\Test-IsFileLocked.ps1"


#endregion


#region Variable Setup


$Template = "$env:HOMEDRIVE\test\Work Shift Log_V2.3_TEMPLATE.xlsx"
$AutomationDir = ("$env:UserProfile" + "\Documents\Automation\")
$FileDate = Get-Date -Format "MM-dd-yy tt"
$ExcelFileDate = Get-Date -Format "yyyyMMdd"
$Filepath = "$env:HOMEDRIVE\test\$FileDate MonitorLog.txt"
$excelpath = "$env:HOMEDRIVE\test\Temp.xlsx"
$st = (Get-Date).adddays(-1)
$Shift = Switch ($env:UserProfile) {
    "$env:HOMEDRIVE\Users\*snip*" { "C" }
    "$env:HOMEDRIVE\Users\*snip*" { "B" }
    "$env:HOMEDRIVE\Users\*snip*" { "D" }
    "$env:HOMEDRIVE\Users\*snip*" { "A" }
}
$exceldestination = ("\\*snip*\" + (Get-Date -Format MM) + "." + (Get-Date -Format yyyy))
$excelfiledestination = ($exceldestination + "\Work Shift Log_V2.3_" + $ExcelFileDate + "_" + "$Shift" + ".xlsx")



#endregion


#region Credentialing

# Goes through and tests for the path where credentialing is stored, if the folder does not exists it is created.

if (-not (Test-Path -Path $AutomationDir)) {
    New-Item -Path $AutomationDir -ItemType "directory" -Force
}

# Series of tests checking for existance of .txt containing encrypted form of credentials. If it does not exists creates them by prompting the user for input.
# Then creates a PS credential combining the username and password

if (-not (Test-Path -Path ($AutomationDir + "Vspherecred.txt"))) {
    Read-Host -AsSecureString "Please enter VMWare Vsphere password only!" | ConvertFrom-SecureString | Out-File ($AutomationDir + "Vspherecred.txt")
}
$VMWarePass = Get-Content ($AutomationDir + "Vspherecred.txt") | ConvertTo-SecureString
$VMWareCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "administrator@vsphere.local", $VMWarePass


if (-not (Test-Path -Path ($AutomationDir + "admincred.txt"))) {
    Read-Host -AsSecureString "Please enter domain administrator password only!" | ConvertFrom-SecureString | Out-File ($AutomationDir + "admincred.txt")
}
$AdminPass = Get-Content ($AutomationDir + "admincred.txt") | ConvertTo-SecureString
$UserCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "usga\administrator", $AdminPass


if (-not (Test-Path -Path ($AutomationDir + "Storagecred.txt"))) {
    Read-Host -AsSecureString "Please enter Unity password only!" | ConvertFrom-SecureString | Out-File ($AutomationDir + "Storagecred.txt")
}
$StoragePassword = Get-Content ($AutomationDir + "Storagecred.txt") | ConvertTo-SecureString
$StorageCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "admin", $StoragePassword


if (-not (Test-Path -Path ($AutomationDir + "iDRACcred.txt"))) {
    Read-Host -AsSecureString "Please enter iDRAC password only!" | ConvertFrom-SecureString | Out-File ($AutomationDir + "iDRACcred.txt")
}
$iDracPass = Get-Content ("$AutomationDir" + "iDRACcred.txt") | ConvertTo-SecureString
$iDracCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "root", $iDracPass


if (-not (Test-Path -Path ($AutomationDir + "EmailPassword.txt"))) {
    Read-Host -AsSecureString "Please enter email password only!" | ConvertFrom-SecureString | Out-File ($AutomationDir + "EmailPassword.txt")
}
$EmailPassword = Get-Content ("$AutomationDir" + "EmailPassword.txt") | ConvertTo-SecureString

if (-not (Test-Path -Path ($AutomationDir + "UserEmail.txt"))) {
    Read-Host -AsSecureString "Please enter your email only!" | ConvertFrom-SecureString | Out-File ($AutomationDir + "UserEmail.txt")
}
$UserEmail = Get-Content ($AutomationDir + "UserEmail.txt")

$EmailCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserEmail, $EmailPassword


#endregion


#region XLSX Setup


# Tests to see if the template file that is used is locked. If it does prompt user to make sure it is unlocked until the lock breaks and it is rechecked.
# Copies template file to a temporary alternative as to not destroy the template or require extra processing

$FileInitial = Test-IsFileLocked -Path $Template
if ($FileInitial.IsLocked) {
    Do {
        Read-Host "Template File seems to be locked? Please close file and hit enter to re-check."
        $FileInitial2 = Test-IsFileLocked -Path $Template
        if ($FileInitial2.IsLocked -ne 'True') {
            Set-Variable -Name FileUnlocked -Value 1
        }
    } Until ($FileUnlocked -eq 1)
}
Copy-Item $Template -Destination $excelpath -Force

#endregion


#region Server-List

# List of various servers that is used later on in the script

$ServerList = [Ordered]@{

    "10.*snip*"  = '*snip*'
    "10.*snip*"  = '*snip*'
    "10.*snip*"  = '*snip*'
    "10.*snip*"  = '*snip*'

}

$iDracIPs = @(
    "10.*snip*"
    "10.*snip*"
    "10.*snip*"

)

$DBList = [Ordered]@{
    "10.*snip*" = "*snip*"
    "10.*snip*" = "*snip*"
    "10.*snip*" = "*snip*"
    "10.*snip*" = "*snip*"
    "10.*snip*" = "*snip*"
}

$ADPPing = @{
    "10.*snip*" = "None"
    "10.*snip*" = "None"
    "10.*snip*" = "None"
    "10.*snip*" = "None"
}

$PrinterPing = @{
    "10.*snip*" = "None"
    "10.*snip*"  = "None"
    "10.*snip*" = "None"
    "10.*snip*"  = "None"

}

$WebList = @(
    "https://10.*snip*"
    "https://10.*snip*/ui/"
    "http://10.*snip*"
    "http://10.*snip*:8080"

)


#endregion


#region iDRAC Call / Output

# For each iDRAC IP address connects, retrieves system status, outputs it to a file and adds a counter if something does not = "OK"

foreach ($IDracIPAddress in $iDracIPs) {
    Write-Output "$IDracIPAddress" | Tee-Object $Filepath -Append
    Connect-iDRAC -iDRAC_IP $IDracIPAddress -iDRAC_Port 443 -Credentials $iDracCredential -trustCert -ErrorAction SilentlyContinue
    Write-Output "IDrac System" | Out-File $Filepath -Append
    Get-iDRACSystemElement -OutVariable iDRACStatus
    $iDRACStatus.Status | Out-File $Filepath -Append
    Write-Output "IDrac Chassis" | Out-File $Filepath -Append
    $iDRACStatus = Get-iDRACChassisElement
    $iDRACStatus.Status | Out-File $Filepath -Append
    $iDRACStatus = $iDRACStatus.Status.Health
    foreach ($member in $iDRACStatus) {
        if ($member -notcontains "OK") { $iDRACERRORCOUNTER++ }
    }
}

# If counter exists outputs to log file and changes template file

if ($iDRACERRORCOUNTER) {
    Write-Output "ERRORS FOUND!! $iDRACERRORCOUNTER unhealthy statuses found" | Tee-Object $Filepath -Append
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData 'ABNORMAL' -StartRow 11 -StartColumn 6
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData 'ABNORMAL' -StartRow 12 -StartColumn 6
}
else {
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData 'NORMAL' -StartRow 11 -StartColumn 6
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData 'NORMAL' -StartRow 12 -StartColumn 6
}


#endregion


#region Server Error Log Check / Output

# For each server reach out and get system logfile within the past day filtered by level 2/3 events and output them to log.

$Events = foreach ($srv in $ServerList.Keys)
{ $srv; Get-WinEvent -computername $srv -FilterHashtable @{logname = "system"; level = 2, 3; starttime = $st } -Credential $UserCredential -ErrorAction SilentlyContinue | format-table id, timecreated, message -auto
}

$Events | Tee-Object $Filepath -Append


#endregion


#region SAN Storage Pool Check / Output

# Connects to SAN Storage server and retrieves storage size, converting to TB then outputting to log and excel sheet

Write-Output "SAN Storage POOL" | Tee-Object $Filepath -Append

Connect-Unity -Server *snip* -Credentials $StorageCredential -ErrorAction SilentlyContinue

Get-UnityPool -OutVariable UnityPoolSize
$UnityPoolSizeConverted = [String](Convert-Size -From Bytes -To TB -Value $UnityPoolSize.sizeFree) + " TB"
$UnityPoolSizeConverted | Out-File $Filepath -Append

Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "$UnityPoolSizeConverted" -StartRow 13 -StartColumn 7


#endregion


#region SAN Storage ALERTS Check / Output

#Connects to SAN Storage server and retrieves any alerts. If it exists writes to $UnityAlerts variable for processing and writing to excel sheet and log

Write-Output "SAN Storage ALERTS" | Tee-Object $Filepath -Append

Get-UnityAlert -OutVariable UnityAlerts | Tee-Object $Filepath -Append

if ($UnityAlerts.Id.Count -ge 1) {
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "ABNORMAL" -StartRow 14 -StartColumn 6
    foreach ($Alert in $UnityAlerts) {
        if ($Alert.IsAcknowledged -eq $false) {
            Set-UnityAlert -ID $Alert.Id -isAcknowledged:$true -Confirm:$false
        }
        $UnityAlertMessage += ([STRING]$Alert.Message + " ")
    }
}
else {
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "NORMAL" -StartRow 14 -StartColumn 6
}

if ($UnityAlertMessage) {
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "$UnityAlertMessage" -StartRow 14 -StartColumn 7
}

Disconnect-Unity -Confirm:$False


#endregion


#region SAN Freespace F/E Check / Output

# Connects to SAN Server and retrieves remaining storage outputting to log file and excel sheet

Write-Output "SAN Storage Drive Remaining" | Tee-Object $Filepath -Append

$SanFreeSpace = Get-WmiObject -ComputerName *snip* -Credential $UserCredential -Class Win32_LogicalDIsk | Select-Object -Property DeviceID, FreeSpace | Where-Object { ($_.DeviceID -contains 'E:') -or ($_.DeviceID -contains 'F:') }

$SanFreeE = $SanFreeSpace.FreeSpace | Select-Object -Index 0
"E: Free Space " + ( Convert-Size -From Bytes -To TB -Value $SanFreeE ) + " TB" | Tee-Object $Filepath -Append -OutVariable SanFreeE
$SanFreeE | Tee-Object $Filepath -Append

$SanFreeF = $SanFreeSpace.FreeSpace | Select-Object -Index 1
"F: Free Space " + ( Convert-Size -From Bytes -To TB -Value $SanFreeF ) + " TB" | Tee-Object $Filepath -Append -OutVariable SanFreeF

Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData ("$SanFreeE" + " " + "$SanFreeF") -StartRow 15 -StartColumn 7


#endregion


#region Veritas Storage Pull

# Only way to access Veritas remaining storage with my tools at hand is through a web portal. Due to complexity with attempting to decode HTTP responses I outsourced this to python and selenium framework
# Takes the "return" of python script and interprets it appropriately

Write-Output "Retrieving Veritas Remaining Storage" | Tee-Object $Filepath -Append

$VeritasRemainingStorage = python $env:HOMEDRIVE\test\WebDrivertest.py
if ( (-not $VeritasRemainingStorage) -or ($VeritasRemainingStorage -eq "GB") -or ($VeritasRemainingStorage -eq "")) {

    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "Error Retrieving Storage" -StartRow 16 -StartColumn 7
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "ABNORMAL" -StartRow 16 -StartColumn 6


}
else {

    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "$VeritasRemainingStorage" -StartRow 16 -StartColumn 7

}

$VeritasRemainingStorage | Tee-Object $Filepath -Append


#endregion


#region DB Cluster Status Check

# Connects to DB Servers and ensures all clusters are listed as "Stable". If there is not an expected amount of stables outputs the server with less than desired results to log and excel file

Write-Output "STARTING DB CLUSTER STATUS CHECK" | Tee-Object $Filepath -Append

Write-Output "*snip*" | Tee-Object $Filepath -Append
Invoke-Command *snip* -Credential $UserCredential { crsctl status resource -t } | Tee-Object -Variable DBCheck1 | Tee-Object $Filepath -Append
$DBConfirm1 = $DBCheck1 | Select-String -SimpleMatch Stable

Write-Output "*snip*" | Out-File $Filepath -Append
Invoke-Command *snip* -Credential $UserCredential { crsctl status resource -t } | Tee-Object -Variable DBCheck2 | Tee-Object $Filepath -Append
$DBConfirm2 = $DBCheck2 | Select-String -SimpleMatch Stable

if ($DBConfirm1.Count -and $DBConfirm2.Count -eq 29) {
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "NORMAL" -StartRow 22 -StartColumn 6
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "NORMAL" -StartRow 23 -StartColumn 6
}
elseif ( ($DBConfirm1.Count -ne 29) -and ($DBConfirm2 -ne 29) ) {
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "10.60.110.50 AND 10.60.110.60 Issue" -StartRow 22 -StartColumn 6
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "10.60.110.50 AND 10.60.110.60 Issue" -StartRow 23 -StartColumn 6
}
elseif ($DBConfirm1.Count -ne 29) {
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "10.60.110.50 Issue" -StartRow 22 -StartColumn 6
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "10.60.110.50 Issue" -StartRow 23 -StartColumn 6
}
elseif ($DBCONfirm2.Count -ne 29) {
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "10.60.110.60 Issue" -StartRow 22 -StartColumn 6
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "10.60.110.60 Issue" -StartRow 23 -StartColumn 6
}


#endregion


#region DB Process Check

# Confirms backup processes are running, compares the ones running against expected list and outputs the difference building a string that is outputted to log and excel

Write-Output "STARTING DB SERVER PROCESS CHECK, 6 REQUIRED" | Tee-Object $Filepath -Append

$DBExpectedProcesses = @(
    '*snip*'
    '*snip*'
)

$DBManagerExpectedProcesses = @(
    '*snip*'
    '*snip*'
)

foreach ($DBServ in $DBList.Keys) {
    Write-Output $DBServ | Out-File $Filepath -Append
    $DBResult = Invoke-Command $DBServ -Credential $UserCredential { Get-Process | Where-Object { $_.ProcessName -like "*shp*" } } -Verbose
    $DBResult | Out-File -FilePath $Filepath -Append
    if ($DBResult.count -le 5) {
        if ($DBServ -eq "*snip*") {
            $MissingProcess = $DBManagerExpectedProcesses | Where-Object { $_ -notin $DBResult }
            foreach ($Process in $MissingProcess) {
                $MisProcString += "$Process "
            }
        }
        else {
            $MissingProcess = $DBExpectedProcesses | Where-Object { $_ -notin $DBResult }
            foreach ($Process in $MissingProcess) {
                $MisProcString += "$Process "
            }
        }
        Write-Output "Error with Services, less than 6 running!" | Tee-Object $Filepath -Append
        $ErrorDBResultHost += ( "$DBList[$DBServ] " + "$DBServ " + "MISSING PROCESS: " + $MisProcString )
        ( "$DBList[$DBServ] " + "$DBServ " + "MISSING PROCESS: " + $MisProcString ) | Tee-Object $Filepath -Append
        $DBResultCount++
    }
    $MisProcString = $null
}

if ( -not $DBResultCount) {
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "NORMAL" -StartRow 25 -StartColumn 6
}
else {
    $ErrorDBResultHost += "Error"
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "ABNORMAL" -StartRow 25 -StartColumn 6
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "$ErrorDBResultHost" -StartRow 25 -StartColumn 7
}


#endregion


#region VMWare Information

# Connects to VMWare server and filters based on non-informational alerts within the last 24 hours

Write-Output "VMWARE Information START" | Tee-Object $Filepath -Append

Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
Connect-VIServer -Server *snip* -Credential $VMWareCredential

$VmWareError = Get-VIEvent -Start $st -MaxSamples ([int]::MaxValue) | Where-Object { $_ -is [VMware.Vim.AlarmStatusChangedEvent] } |
Group-Object -Property { $_.Entity.Entity } | ForEach-Object {
    $_.Group | Sort-Object -Property CreatedTime | Select-Object -last 1 |
    Where-Object { "green", "gray" -notcontains $_.To } | Select-Object CreatedTime, @{N = "Entity"; E = { $_.Entity.Name } }, To, @{N = "Alarm"; E = { $_.Alarm.Name } }
} -Verbose

Disconnect-VIServer -Server *snip* -Confirm:$false

if ($VmWareError) {
    $VmWareError | Out-File $Filepath -Append
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "ABNORMAL" -StartRow 19 -StartColumn 6
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "ABNORMAL" -StartRow 20 -StartColumn 6
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "$VmWareError".TrimStart("@" , "{").TrimEnd("}") -StartRow 19 -StartColumn 7
}
else {
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "NORMAL" -StartRow 19 -StartColumn 6
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "NORMAL" -StartRow 20 -StartColumn 6
}

Write-Output "VMWARE Information END" | Tee-Object $Filepath -Append


#endregion


#region Ping Test

# Goes through and pings one set of devices to ensure that they are online. It does this by cloning a hashtable and using that collection with a ping async module
# The values from this are passed back to that hashtable. If the list contains timedout iterates over it again pinging what is not successfully pinged. It does this
# a total of five times waiting five seconds between before eventually settling on timedout and writing out an error

Write-Output "ADP Ping Block Starting" | Tee-Object $Filepath -Append


$ADPPingResult = Test-ConnectionAsync $ADPPing.Keys
$ADPPingResult | ForEach-Object {
    [String]$Compname = $_.Computername
    [String]$Result = $_.Result

    $ADPPing."$Compname" = $Result
}

if ($ADPPing.Values -contains "TimedOut") {
    Do {
        $ADPPingCopy = $ADPPing.Clone()
        $ADPPingCopy.GetEnumerator() | ForEach-Object {

            [String]$CompnameRepeat = $_.Name
            [String]$ResultRepeat = $_.Value

            if ($ResultRepeat -eq "TimedOut") {
                $ADPPingResult = Test-ConnectionAsync $CompnameRepeat -Verbose
                $ADPPingResult | ForEach-Object {
                    [String]$CompnameRepeat = $_.Computername
                    [String]$ResultRepeat = $_.Result

                    $ADPPing."$CompnameRepeat" = $ResultRepeat
                }
            }
            Start-Sleep -Seconds 5
        }
        $Count++
    }Until ( ($Count -ge 5) -or ($ADPPing.Values -notcontains "TimedOut" ) )
}

if ($ADPPing.Values -contains "TimedOut") {

    $ADPPingCopy.GetEnumerator() | ForEach-Object {

        [String]$CompnameRepeat = $_.Name
        [String]$ResultRepeat = $_.Value

        if ($ResultRepeat -eq "TimedOut") {
            $ADPTimedoutString += "$CompnameRepeat ,"
        }
    }

    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "ABNORMAL" -StartRow 29 -StartColumn 6
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData ("$ADPTimedoutString" + " Timed Out") -StartRow 29 -StartColumn 7

}

$ADPPing | Tee-Object $Filepath -Append

Write-Output "Printer Ping Block Starting" | Tee-Object $Filepath -Append


$PrinterPingResult = Test-ConnectionAsync $PrinterPing.Keys

$PrinterPingResult | ForEach-Object {
    [String]$Compname = $_.Computername
    [String]$Result = $_.Result

    $PrinterPing."$Compname" = $Result
}

if ($PrinterPing.Values -contains "TimedOut") {
    Do {
        $PrinterPingCopy = $PrinterPing.Clone()
        $PrinterPingCopy.GetEnumerator() | ForEach-Object {

            [String]$CompnameRepeat = $_.Name
            [String]$ResultRepeat = $_.Value

            if ($ResultRepeat -eq "TimedOut") {
                $PrinterPingResult = Test-ConnectionAsync $CompnameRepeat -Verbose
                $PrinterPingResult | ForEach-Object {
                    [String]$CompnameRepeat = $_.Computername
                    [String]$ResultRepeat = $_.Result

                    $PrinterPing."$CompnameRepeat" = $ResultRepeat
                }
            }
            Start-Sleep -Seconds 5
        }
        $PrinterCount++
    }Until ( ($PrinterCount -ge 5) -or ($PrinterPing.Values -notcontains "TimedOut" ) )
}

if ($PrinterPing.Values -contains "TimedOut") {

    $PrinterPingCopy.GetEnumerator() | ForEach-Object {

        [String]$CompnameRepeat = $_.Name
        [String]$ResultRepeat = $_.Value

        if ($ResultRepeat -eq "TimedOut") {
            $PrinterTimedoutString += "$CompnameRepeat ,"
        }
    }

    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData "ABNORMAL" -StartRow 30 -StartColumn 6
    Export-Excel -NoNumberConversion * -Path $excelpath -WorksheetName '1. Work Shift Log' -TargetData ("$PrinterTimedoutString" + " Timed Out") -StartRow 30 -StartColumn 7

}

$PrinterPing | Tee-Object $Filepath -Append

#endregion


#region Email

# Attaches created log file and emails it to the user who is running the script

Send-MailMessage -SmtpServer smtp.office365.com -Attachments $Filepath -Body "Hello! Please see attached documents" -To $UserEmail -From $UserEmail -Subject ("$FileDate " + "$Shift Shift " + "Server Report") -Credential $EmailCredential -UseSsl


#endregion


#region Excel Output

# Confirms directory exists where excel file should be outputted to and if it does not exist it creates it. It then copies the temporary file with a formatted name
# and deletes the old temporary file

if ( -not ( Test-Path -Path $exceldestination ) ) {
    New-Item -Path $exceldestination -ItemType "directory" -Force
}
Copy-Item $excelpath -Destination $excelfiledestination -Force -Verbose

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $true
$Workbook = $Excel.Workbooks.Open($excelfiledestination)

Remove-Item "$env:HOMEDRIVE\test\Temp.xlsx"

#endregion


#region Browser Open

# Checks if chrome is open, if it is prompts user if they want to 1) close chrome 2) use existing windows 3) continue without opening the files. This is done
# due to --ignore-certificate-errors which ignores SSL certificate errors on many internally hosted sites. Given a large amount of sites this can add significant time

if (-not (Get-Process "chrome") ) {

    foreach ($Site in $WebList) {
        start-process "$env:HOMEDRIVE\Program Files (x86)\Google\Chrome\Application\chrome.exe" "$Site", '--profile-directory="Default" --ignore-certificate-errors'
    }

}
else {
    Do {
        $ChromeSwitch = Read-Host "Chrome is currently running! While Chrome is running opening new windows will not ignore SSL errors. Continue?
        1. Open tabs in existing Chrome (Will have SSL Errors)
        2. Close Chrome and reopen
        3. Do not open sites"
    }While (1, 2, 3 -notcontains $ChromeSwitch)
    Switch ($ChromeSwitch) {
        1 {
            foreach ($Site in $WebList) {
                start-process "$env:HOMEDRIVE\Program Files (x86)\Google\Chrome\Application\chrome.exe" "$Site", '--profile-directory="Default" --ignore-certificate-errors'
            }
        }
        2 {
            Get-Process "chrome" | Stop-Process
            foreach ($Site in $WebList) {
                start-process "$env:HOMEDRIVE\Program Files (x86)\Google\Chrome\Application\chrome.exe" "$Site", '--profile-directory="Default" --ignore-certificate-errors'
            }
        }
        3 {
            Break
        }
    }

}

#endregion


#region Start Kaspersky

# Starts local Kaspersky application to check healthy systems. I have not found a way to do this systematically through an api.... YET

start-process "$env:HOMEDRIVE\Program Files (x86)\Kaspersky Lab\Kaspersky Security Center Console\Kaspersky Security Center 11.msc"


#endregion


#region Cleanup

# Deletes all new variables that were created after the script was started.

$NewVariables = Get-Variable | Select-Object -ExpandProperty Name | Where-Object { $ExistingVariables -notcontains $_ -and $_ -ne "ExistingVariables" }
if ($NewVariables) {
    Write-Host "Removing the following variables:`n`n$NewVariables"
    Remove-Variable $NewVariables -ErrorAction SilentlyContinue
}
else {
    Write-Host "No new variables to remove!"
}


#endregion