Add-Type -assembly "Microsoft.Office.Interop.Outlook"
Import-Module ActiveDirectory
$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)
function Mail-Outline {
    $Mail.Bodyformat = 1
    #Address to should be '#Snipped#' for SC Server#
    $Mail.To = '#Snipped#'
    $Mail.CC = '#Snipped#'

    #Subject should be blank for the SC Server to see the message#
    $Mail.Subject = ""}

    function Mail-OutlineA {
    $Mail.Bodyformat = 1
    #Address to should be '#Snipped#' for SC Server#
    $Mail.To = '#Snipped#'
    $Mail.CC = '#Snipped#'
    $Mail.Attachments.Add($fileName)
    #Subject should be blank for the SC Server to see the message#
    $Mail.Subject = ""}

Do{
Do{ $MenuChoice = Read-Host -Prompt "Please Select Ticket Process
1. AMS
2. Tracking
3. New Hire Bundle
4. Termination
5. New / Refresh Request Hardware Asset Management
6. External Tracking Ticket
7. Pending
8. Pending
9. User Lookup"} While (1,2,3,4,5,6,9 -notcontains $MenuChoice)
Switch ($MenuChoice) {

#AMS TICKET TEMPLATE#
1 {
Do{Write-Host "Please Select Email file"
$OpenFolderDialog = New-Object System.Windows.Forms.OpenFileDialog 
# must be true for OpenFileDialog, otherwise it hangs
$OpenFolderDialog.ShowDialog()
Write-Host "FileDialog1 Input:  $OpenFolderDialog.filename"
$fileName = $OpenFolderDialog.filename} While ($fileName -eq "")
$Text = Read-Host -Prompt "Paste Email Text"
Do{
$AssignC = Read-Host -Prompt "Assignment Group:
1. #Snipped#
2. #Snipped#
3. #Snipped#  "
if ($AssignC -eq "admin" ) 
        {$Assign = Read-Host -Prompt "Enter Assignment Group"
            Break}
} While (1,2,3 -notcontains $AssignC)
Do{
$Format = Read-Host -Prompt "Which Format?
1. #Snipped#
2. #Snipped#"
} While (1,2 -notcontains $Format)
Do{
$OperatorC = Read-Host -Prompt "Which Operator?
1. #Snipped#
2. #Snipped#"
if ($OperatorC -eq "admin" ) 
        {$Operator = Read-Host -Prompt "Enter Operator Email"
            Break}
}While (1,2 -notcontains $OperatorC)
$CSV = Import-Csv -path "H:\Ticketform.csv"
$CSV | ForEach-Object {
    $FirstName = $_.FirstName
    $LastName = $_.LastName
    $Program = $_.Program
    $Problem = $_.Problem
    $#Snipped# = $#Snipped#

$Program = $Program.ToUpper()
$#Snipped# = $#Snipped#.ToUpper()
$FirstName = (Get-Culture).TextInfo.ToTitleCase($FirstName)
$LastName = (Get-Culture).TextInfo.ToTitleCase($LastName)

If ($Format -eq 1) {$Title = "#Snipped#"}
elseif ($Format -eq 2) {$Title = "#Snipped#"}

If ($AssignC -eq 1) {$Assign = "#Snipped#"}
elseif ($AssignC -eq 2) {$Assign = "#Snipped#"}
elseif ($AssignC -eq 3) {$Assign = "#Snipped#"}

If ($OperatorC -eq 1) {$Operator = "#Snipped#"}
elseif ($OperatorC -eq 2) {$Operator = "#Snipped#"}

If ($Text -eq "") {$Text = "There is no associated email text relevant"}

        #Email/Ticket to AMS AMS Onsite Support Chattanooga  
    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)

    #Formats as Plain Text for SC Server
    $Mail.Bodyformat = 1
    $Mail.Attachments.Add($fileName)
    #Address to should be 'ServiceCenter@volkswagen.de' for SC Server#
    $Mail.To = '#Snipped#'
    $Mail.CC = '#Snipped#'

    #Subject should be blank for the SC Server to see the message#
    $Mail.Subject = ""

    #Body is the context of the ticket to be made, Description being the ticket description#
    $Mail.Body ="#Snipped#


Send to : $Assign


$Text
>
"
    $Mail.Send()
    }
$firstLine = Get-Content "H:\Ticketform.csv" -First 1
Clear-Content "H:\Ticketform.csv"
$firstLine | Set-Content "H:\Ticketform.csv"
Remove-Item -Path $fileName
Do {$Restart = Read-Host -Prompt "Would you like to start again?
1. Yes
2. No"} while (1,2 -notcontains $Restart)
     #WhileLoop Closure#
    } #SwitchLoop Closure#


#Tracking Ticket Template#
2 {
Write-Host "This is to be used only for tickets going to #Snipped# for tracking purposes!"
$#Snipped# = ""
$FirstName = ""
$LastName = ""
$ADGet = ""
$UserPhone = ""
$#Snipped# = Read-Host -Prompt "Enter #Snipped# of User"
$#Snipped# = $#Snipped#.ToUpper()
if ($DVU -ne "") {    
$ADGet = Get-ADUser $DVU -Properties *
$FirstName = $ADGet.GivenName
$LastName = $ADGet.Surname
$UserEmail = $ADGet.EmailAddress
$OfficeNumber = $ADGet.OfficePhone
$UserPhone = $ADGet.MobilePhone
Do{
$Nameconfirm = Read-Host -Prompt "$FirstName $LastName $DVU,$UserEmail is this correct?
1. Yes
2. No"
  } while (1,2 -notcontains $Nameconfirm) 
 if ($Nameconfirm -eq 2) { 
 Write-Host "Please enter user First / Last for AD Lookup"
 $FirstNameCall = Read-Host -Prompt "Enter First Name of User"
 $LastNameCall = Read-Host -Prompt "Enter Last Name of User"
        Try{
 $ADGet = Get-ADUser -Filter {(GivenName -eq $FirstNameCall -and Surname -eq $LastNameCall -and Enabled -eq "True")} -Properties * 
        }Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Write-Host "No user found, please continue manually"
        Break}
        Catch {Write-Host "There was an error, you may not have Active Directory installed"
               Break}
        
 $ADGet | ForEach-Object { 
 $FirstName = $_.GivenName
 $LastName = $_.Surname
 $#Snipped# = $_.SamAccountName
 $UserEmail = $_.EmailAddress
 Do {$Nameconfirm = Read-Host -Prompt "$FirstName $LastName $#Snipped#, $UserEmail is this correct?
 1. Yes
 2. No"
 } while (1,2 -notcontains $Nameconfirm) }
 }
 }
If ($NameConfirm -eq 2) {
$FirstName = Read-Host -Prompt "Enter First Name of User"
$LastName = Read-Host -Prompt "Enter Last Name of User"
}
$Program = Read-Host -Prompt "Enter HEADER / Program for ticket"
$Problem = Read-Host -Prompt "Enter Problem of User"
$Text = Read-Host -Prompt "Notes for ticket"
If ($Text -eq "") {$Text = "There is no associated text or notes relevant, please see title"}
#Text Cleanup#
$Program = $Program.ToUpper()
$FirstName = (Get-Culture).TextInfo.ToTitleCase($FirstName)
$LastName = (Get-Culture).TextInfo.ToTitleCase($LastName)
        #Email/Ticket to #Snipped# 
    Mail-Outline

    #Body is the context of the ticket to be made, Description being the ticket description#
    $Mail.Body ="#Snipped#
$Text

Please complete, this is solely for tracking purposes

>
"
    $Mail.Send()


if ($MobilePhone -eq "") {$MobilePhone = Write-Host "N/A"}
if ($OfficePhone -eq "") {$OfficePhone = Write-Host "N/A"}
$Time = Get-Date
$FileTime = Get-Date -format ddMM
$FilePathV = "H:\Call and Ticket Log\$FileTime CallLog.txt"
$CallLogText = Write-Output "
Time: $Time
Name: $FirstName $LastName
#Snipped#
Phone: Office Phone: $OfficePhone / Mobile Phone: $MobilePhone
Issue: $Program / $Problem
Ticket: "
$CallLogText | Out-File -Append -FilePath $FilePathV


Do {$Restart = Read-Host -Prompt "Would you like to start again?
1. Yes
2. No"} while (1,2 -notcontains $Restart)
    } #MenuChoice Switch Closure#
    
#New Hire Bundle#
3 {
Write-Host "Please Select New Hire Snip image"
Do{
$OpenFolderDialog = New-Object System.Windows.Forms.OpenFileDialog 
# must be true for OpenFileDialog, otherwise it hangs
$OpenFolderDialog.ShowDialog()
Write-Host "FileDialog1 Input:  $OpenFolderDialog.filename"
$fileName = $OpenFolderDialog.filename
Write-Host $fileName
if ($fileName -eq "") {Write-Host "Please select New Hire Document!"}
} until ($fileName -ne "")
$FirstName = Read-Host -Prompt "Enter First Name of User"
$FirstName = (Get-Culture).TextInfo.ToTitleCase($FirstName)
$LastName = Read-Host -Prompt "Enter Last Name of User"
$StartDate = Read-Host -Prompt "Enter Start Date MM/DD/YEAR"
$LastName = (Get-Culture).TextInfo.ToTitleCase($LastName)
Do { $EmployeeType = Read-Host -Prompt "What kind of an employee is this?
1. Contractor
2. #Snipped#
3. Intern" } While (1,2,3 -notcontains $EmployeeType)

if ($EmployeeType -eq 1) {$EmployeeType = "Contractor"}
elseif ($EmployeeType -eq 2) {$EmployeeType = "#Snipped#"}
elseif ($EmployeeType -eq 3) {$EmployeeType = "Intern"}
else {Write-Host "Well, you broke it somehow. Terminating"
            Exit}

$REQNumber = Read-Host -Prompt "Enter REQ Number"
$REQNumber = $REQNumber.ToUpper()

Do {$WorkflowChoice1 = Read-Host -Prompt "Does this person need a new asset?
1. No
2. Desktop
3. Laptop"} while (1,2,3 -notcontains $WorkflowChoice1)

if ($WorkflowChoice1 -eq 2) {$AssetType = "Desktop"}
elseif ($Workflowchoice1 -eq3) {$AssetType = "Laptop"}

Do { $WorkflowChoice3 = Read-Host -Prompt "Does this person need share access?
1. No
2. Yes
(Please enter additional Manually, the next prompt will ask for the path)" } While (1,2 -notcontains $WorkflowChoice3)
Do {
if ($WorkflowChoice3 -eq 2) {$ShareLink = Read-Host -Prompt "Please Enter Share Address"}
else {break}
    } while ($ShareLink -eq "")
      

Do {
$Text = Read-Host -Prompt "Please Paste Workflow Text"
}While ($Text -eq "") 

     #Email/Ticket to #Snipped# for tracking 
 $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)

    #Formats as Plain Text for SC Server
    $Mail.Bodyformat = 1
    #Address to should be '#Snipped#' for SC Server#
    $Mail.To = '#Snipped#'

    #Subject should be blank for the SC Server to see the message#
    $Mail.Subject = ""
    if ($fileName -ne ""){
    $Mail.Attachments.Add($fileName) 
    }
    else {Write-Host "No File Attached!"}
    #Body is the context of the ticket to be made, Description being the ticket description#
    $Mail.Body ="#Snipped#
"
    $Mail.Send()

    #Email/Ticket to Asset Management for Asset 
 if (1 -notcontains $WorkflowChoice1) { 
    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)
    $Mail.Bodyformat = 1
    #Address to should be '#Snipped#' for SC Server#
    $Mail.To = '#Snipped#'

    #Subject should be blank for the SC Server to see the message#
    $Mail.Subject = ""
    if ($fileName -ne ""){
    $Mail.Attachments.Add($fileName) 
    }
    else {Write-Host "No File Attached!"}
    #Body is the context of the ticket to be made, Description being the ticket description#
    $Mail.Body ="#Snipped#
"
    $Mail.Send()


    }
     #Email/Ticket for Skype Creation
 $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)

    #Formats as Plain Text for SC Server
    $Mail.Bodyformat = 1
    #Address to should be '#Snipped#' for SC Server#
    $Mail.To = '#Snipped#'

    #Subject should be blank for the SC Server to see the message#
    $Mail.Subject = ""
    if ($fileName -ne ""){
    $Mail.Attachments.Add($fileName) 
    }
    else {Write-Host "No File Attached!"}
    #Body is the context of the ticket to be made, Description being the ticket description#
    $Mail.Body ="#Snipped#
>
"
    $Mail.Send()
    
     #Email/Ticket to OPS for Share Drive Creation 
 if ($WorkflowChoice3 -eq 2) {$Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)

    #Formats as Plain Text for SC Server
    $Mail.Bodyformat = 1
    #Address to should be '#Snipped#' for SC Server#
    $Mail.To = '#Snipped#'

    #Subject should be blank for the SC Server to see the message#
    $Mail.Subject = ""
    if ($fileName -ne ""){
    $Mail.Attachments.Add($fileName) 
    }
    else {Write-Host "No File Attached!"}
    #Body is the context of the ticket to be made, Description being the ticket description#
    $Mail.Body ="#Snipped#

$Text

>
"

Remove-Item -path $fileName

    $Mail.Send()
    } #Workflow Choice Closure
Do {$Restart = Read-Host -Prompt "Would you like to start again?
1. Yes
2. No"} while (1,2 -notcontains $Restart)

    } #WhileLoop Closure#
    
    
    

#Termination Ticket Template#
4 {
Write-Host "Please Select Termination PDF or Email"
Do{
$OpenFolderDialog = New-Object System.Windows.Forms.OpenFileDialog 
# must be true for OpenFileDialog, otherwise it hangs
$OpenFolderDialog.ShowDialog()
Write-Host "FileDialog1 Input:  $OpenFolderDialog.filename"
$fileName = $OpenFolderDialog.filename
}while ($fileName -eq "") {Write-Host "Please select Termination PDF or Email!"}
$FirstName = Read-Host -Prompt "Enter First Name of User"
$FirstName = (Get-Culture).TextInfo.ToTitleCase($FirstName)
$LastName = Read-Host -Prompt "Enter Last Name of User"
$LastName = (Get-Culture).TextInfo.ToTitleCase($LastName)
$TerminationDate = Read-Host -Prompt "Enter Termination Date MM/DD/YEAR"

         #Email/Ticket to #Snipped#
 $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)

    #Formats as Plain Text for SC Server
    $Mail.Bodyformat = 1

    #Address to should be '#Snipped#' for SC Server#
    $Mail.To = '#Snipped#'
    $Mail.Attachments.Add($fileName)
    #Subject should be blank for the SC Server to see the message#
    $Mail.Subject = ""

    #Body is the context of the ticket to be made, Description being the ticket description#
    $Mail.Body ="#Snipped#
"
    $Mail.Send()

 #Email/Ticket to #Snipped#  
 $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)

    #Formats as Plain Text for #Snipped#
    $Mail.Bodyformat = 1

    #Address to should be '#Snipped#' for SC Server#
    $Mail.To = '#Snipped#'
    $Mail.Attachments.Add($fileName)
    #Subject should be blank for the SC Server to see the message#
    $Mail.Subject = ""

    #Body is the context of the ticket to be made, Description being the ticket description#
    $Mail.Body ="#Snipped#
"
    $Mail.Send()


    #Email/Ticket to OPS Phone Onsite Support Chattanooga 
 $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)

    #Formats as Plain Text for SC Server
    $Mail.Bodyformat = 1

    #Address to should be 'ServiceCenter@volkswagen.de' for SC Server#
    $Mail.To = '#Snipped#'
    $Mail.Attachments.Add($fileName)
    #Subject should be blank for the SC Server to see the message#
    $Mail.Subject = ""

    #Body is the context of the ticket to be made, Description being the ticket description#
    $Mail.Body ="#Snipped#
"
    $Mail.Send()


        #Email/Ticket to #Snipped#
 $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)

    #Formats as Plain Text for SC Server
    $Mail.Bodyformat = 1

    #Address to should be '#Snipped#e' for SC Server#
    $Mail.To = '#Snipped#'
    $Mail.Attachments.Add($fileName)
    #Subject should be blank for the SC Server to see the message#
    $Mail.Subject = ""

    #Body is the context of the ticket to be made, Description being the ticket description#
    $Mail.Body ="#Snipped#
> 
"
    $Mail.Send()

            #Email/Ticket to #Snipped# 
 $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)

    #Formats as Plain Text for SC Server
    $Mail.Bodyformat = 1

    #Address to should be '#Snipped#' for SC Server#
    $Mail.To = '#Snipped#'
    $Mail.Attachments.Add($fileName)
    #Subject should be blank for the SC Server to see the message#
    $Mail.Subject = ""

    #Body is the context of the ticket to be made, Description being the ticket description#
    $Mail.Body ="#Snipped#
> 
"
    $Mail.Send()

Remove-Item -path $fileName 
Write-Host "You are the

| |                    (_)           | |            
| |_ ___ _ __ _ __ ___  _ _ __   __ _| |_ ___  _ __ 
| __/ _ \ '__| '_ ` _ \| | '_ \ / _` | __/ _ \| '__|
| ||  __/ |  | | | | | | | | | | (_| | || (_) | |   
 \__\___|_|  |_| |_| |_|_|_| |_|\__,_|\__\___/|_|   
 
                    ______
                     <((((((\\\
                     /      . }\
                     ;--..--._|}
  (\                 '--/\--'  )
   \\                | '-'  :'|
    \\               . -==- .-|
     \\               \.__.'   \--._
     [\\          __.--|       //  _/'--.
     \ \\       .'-._ ('-----'/ __/      \
      \ \\     /   __>|      | '--.       |
       \ \\   |   \   |     /    /       /
        \ '\ /     \  |     |  _/       /
         \  \       \ |     | /        /
          \  \      \        /

"
Do {$Restart = Read-Host -Prompt "Would you like to start again?
1. Yes
2. No"} while (1,2 -notcontains $Restart)
    }

    #Hardware Asset Request#
5 {
$FirstName = Read-Host -Prompt "Enter First Name of User"
$FirstName = (Get-Culture).TextInfo.ToTitleCase($FirstName)
$LastName = Read-Host -Prompt "Enter Last Name of User"
$LastName = (Get-Culture).TextInfo.ToTitleCase($LastName)
$#Snipped# = Read-Host -Prompt "Enter #Snipped# of User"
$#Snipped# = $#Snipped#.ToUpper()
Do {$AssetTypeT = Read-Host -Prompt "What kind of Asset is this for?
1. Laptop
2. Desktop
3. Monitor
4. Other"} while (1,2,3,4 -notcontains $AssetTypeT)
if ($AssetTypeT -eq 1) {$AssetTypeT = "Laptop"}
elseif ($AssetTypeT -eq 2) {$AssetTypeT = "Desktop"}
elseif ($AssetTypeT -eq 3) {$AssetTypeT = "Monitor"}
elseif ($AssetTypeT -eq 4) {$AssetTypeT = Read-Host -Prompt "What hardware asset is this about?"}
$AssetTypeT = $AssetTypeT.ToUpper()
$AssetReqT = Read-Host -Prompt "What is needing to be done?
1. Refresh Request
2. New Asset"
if ($AssetReqT -eq 1) {$AssetReqT = "Refresh Request"}
elseif ($AssetReqT -eq 2) {$AssetReqT = "New Asset"}
if ($AssetReqT -eq "Refresh Request") {$OldAssetNum = Read-Host -Prompt "Please enter tag of existing / old asset"}
$Text = Read-Host -Prompt "Notes for ticket"
If ($Text -eq "") {$Text = "There is no associated text or notes relevant, please see title"}
if ($OldAssetNum -eq "") {$Title = "#Snipped#"
else {$Title = "#Snipped#" 
        #Email/Ticket to #Snipped#  
 $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)

    #Formats as Plain Text for SC Server
    $Mail.Bodyformat = 1
    #Address to should be '#Snipped#' for SC Server#
    $Mail.To = '#Snipped#'

    #Subject should be blank for the SC Server to see the message#
    $Mail.Subject = ""

    #Body is the context of the ticket to be made, Description being the ticket description#
    $Mail.Body ="#Snipped#
"
    $Mail.Send()

Do {$Restart = Read-Host -Prompt "Would you like to start again?
1. Yes
2. No"} while (1,2 -notcontains $Restart)
    } #MenuChoice Switch Closure#
    }
    }
#Extern Tracking Ticket
6 {
Write-Host "Please Select Relevant Files (if any)"
$OpenFolderDialog = New-Object System.Windows.Forms.OpenFileDialog 
# must be true for OpenFileDialog, otherwise it hangs
$OpenFolderDialog.ShowDialog()
Write-Host "FileDialog1 Input:  $OpenFolderDialog.filename"
$fileName = $OpenFolderDialog.filename
if ($fileName -eq "") {Write-Host "No File Attached"}
Write-Host "This is to be used only for tickets going to #Snipped# for tracking purposes!"
Try{
$#Snipped# = Read-Host -Prompt "Enter #Snipped# of User"
$#Snipped# = $#Snipped#.ToUpper()
if ($#Snipped# -eq "") {$No#Snipped# = "1"}
if ($No#Snipped# -ne "1") {$ADGet = Get-ADUser $#Snipped# -Properties GivenName , Surname
$FirstName = $ADGet.GivenName
$LastName = $ADGet.Surname
$Nameconfirm = Read-Host -Prompt "$FirstName $LastName $#Snipped#, is this correct?
1. Yes
2. No"
if ($Nameconfirm -eq 1) {$NameConfirmV = "1"}
elseif ($Nameconfirm -eq2) {Write-Host "Please enter user information manually"}}
}Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {Write-Host "No user found with this #Snipped#! Please enter manually"}
 Catch {Write-Host "Active Directory might not be installed on this computer, or you did not put a #Snipped#. please enter manually"}
If ($NameConfirmV -ne 1) {
$FirstName = Read-Host -Prompt "Enter First Name of User"
$LastName = Read-Host -Prompt "Enter Last Name of User"
}
$Program = Read-Host -Prompt "Enter HEADER / Program for ticket"
$Problem = Read-Host -Prompt "Enter Problem of User"
$REQNumber = Read-Host -Prompt "Enter REQ Number"
$Text = Read-Host -Prompt "Notes for ticket"
If ($Text -eq "") {$Text = "There is no associated text or notes relevant, please see title"}
#Text Cleanup#
$REQNumber = $REQNumber.ToUpper()
$Program = $Program.ToUpper()
$FirstName = (Get-Culture).TextInfo.ToTitleCase($FirstName)
$LastName = (Get-Culture).TextInfo.ToTitleCase($LastName)

     #Email/Ticket to OPS SD Service Desk Chattanooga for tracking 
 $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)

    #Formats as Plain Text for SC Server
    $Mail.Bodyformat = 1
    #Address to should be '#Snipped#' for SC Server#
    $Mail.To = '#Snipped#'

    #Subject should be blank for the SC Server to see the message#
    $Mail.Subject = ""
    if ($fileName -ne "") {
    $Mail.Attachments.Add($fileName)
    }
    #Body is the context of the ticket to be made, Description being the ticket description#
    $Mail.Body ="#Snipped#
>
"
    $Mail.Send()

Do {$Restart = Read-Host -Prompt "Would you like to start again?
1. Yes
2. No"} while (1,2 -notcontains $Restart)

    } ## END OF EXTERN TICKET CREATION
    
 9{
   Do {

Write-Host "User Lookup"
$ADUserNotes = ""
$WindowsLockAlert = ""
$WindowsPasswordAlert = ""
$#Snipped# = ""
$FirstName = ""
$LastName = ""
$ADGet = ""
$UserPhone = ""
$UserLockoutS = ""
$UserPasswordExpired = ""
$WindowsEnabledAlert = ""
$UserEnabledFlag = ""
$MenuChoiceL = Read-Host -Prompt "Lookup with DVU or Name?
1. #Snipped#
2. Name
3. Close"
Switch ($MenuChoiceL) {

    1{
Do{
$#Snipped# = Read-Host -Prompt "Enter #Snipped# of User"
$#Snipped# = $#Snipped#.ToUpper()
if ($#Snipped# -eq "") {Write-Host "Please input #Snipped#!"}
   } while ($#Snipped# -eq "")
Try{
 $ADGet = Get-ADUser $#Snipped# -Properties * 
        }Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Write-Host "No user found with this #Snipped#, try name lookup"}       

$FirstName = $ADGet.GivenName
$LastName = $ADGet.Surname
$UserEmail = $ADGet.EmailAddress
$OfficeNumber = $ADGet.OfficePhone
$UserPhone = $ADGet.MobilePhone
$ADUserNotes = $ADGet.info
$UserLockoutS = $ADGet.LockedOut
$UserPasswordExpired = $ADGet.PasswordExpired
$UserEnabledFlag = $ADGet.Enabled
if ($UserEnabledFlag -eq "False") {$WindowsEnabledAlert = "This users account is currently DISABLED in AD!"}
if ($UserLockoutS -eq "True") {$WindowsLockAlert = "This users account is currently locked out of AD!"}
if ($UserPasswordExpired -eq "True") {$WindowsPasswordAlert = "This users password is expired in AD!"}
Write-Host "$FirstName $LastName
$DVU, $UserEmail
Office #: $OfficeNumber
Mobile #: $UserPhone
Notes: $ADUserNotes
$WindowsLockAlert
$WindowsPasswordAlert
$WindowsEnabledAlert
"

    
} #Switch 1 Close

  2{

 Write-Host "Please enter user First / Last for AD Lookup"
 $FirstNameCall = Read-Host -Prompt "Enter First Name of User"
 $LastNameCall = Read-Host -Prompt "Enter Last Name of User"
        Try{
 $ADGet = Get-ADUser -Filter {(GivenName -eq $FirstNameCall -and Surname -eq $LastNameCall -and Enabled -eq "True")} -Properties * 
        }Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Write-Host "No user found, please continue manually"
        Break}
        Catch {Write-Host "There was an error, you may not have Active Directory installed"
               Break}
        
 $ADGet | ForEach-Object { 
 $FirstName = $_.GivenName
 $LastName = $_.Surname
 $#Snipped# = $_.SamAccountName
 $UserEmail = $_.EmailAddress
 $OfficeNumber = $_.OfficePhone
 $UserPhone = $_.MobilePhone
 $ADUserNotes = $_.info
 $UserLockoutS = $_.LockedOut
 $UserPasswordExpired = $_.PasswordExpired
 if ($UserEnabledFlag -eq "False") {$WindowsEnabledAlert = "This users account is currently DISABLED in AD!"}
 if ($UserLockoutS -eq "True") {$WindowsLockAlert = "This users account is currently locked out of AD!"}
 if ($UserPasswordExpired -eq "True") {$WindowsPasswordAlert = "This users password is expired in AD!"}

Write-Host "$FirstName $LastName
$#Snipped#, $UserEmail
Office #: $OfficeNumber
Mobile #: $UserPhone
Notes: $ADUserNotes
$WindowsLockAlert
$WindowsPasswordAlert
$WindowsEnabledAlert
"
 }
 }
 3 {
   Write-Host "Closing to Main Menu"
   $MenuChoiceL = "3"
   $Restart = "1"
   } # Switch 3 Close



 } #Main Switch Close
 } while ($MenuChoiceL -ne "3")#Repeat Close


   } 
    } #Switch Close
    } while (1 -contains $Restart) # Repeat Close#
