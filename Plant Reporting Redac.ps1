Add-Type -assembly "Microsoft.Office.Interop.Outlook"

$Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)
    $Mail.SentOnBehalfOfName = '#Snipped#'
    $Mail.To = '#Snipped#'
    $Mail.CC = '#Snipped#'
    #Subject should be blank for the SC Server to see the message#
    $Time = Get-Date -f H:mm
    $Mail.Subject = "Production Cockpit Overview - $Time"
    Invoke-WebRequest #Snipped# -OutFile "#Snipped#"
    $Mail.Attachments.Add("#Snipped#")
    #Attachment to be added to Email which Adds the attachment to the Ticket

    #Body is the context of the ticket to be made, Description being the ticket description#
    $Mail.Body ="#Snipped#
"
    $Mail.Send()
    Remove-Item "#Snipped#"