###
#HTML Report Style Input
$Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@

###



do{

Try{

$access = Invoke-WebRequest -Uri "https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml" -TimeoutSec 15 -UseDefaultCredentials
##https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml- change this URL to the required URL
$check = ($access.AllElements  | ?{$_.tagName -like "IFRAME"}).src | select -First 1

$temp = "" | select Status,Action
$temp.Status = $check


if($check -like "*block*")
{
$smtpServer = "smtp"
$smtpFrom = "Exchange@domain.com"
$smtpTo = "admin1@domain.com,admin2_URLTeam@dmain.com"
$message = New-Object System.Net.Mail.MailMessage $smtpfrom, $smtpto
$message.Subject = "Action Required : URL is blocked"
$message.IsBodyHTML = $true
$action = "Contact URL Team"
$temp.Action = $action
$body =  $temp | select "Status","Action" | ConvertTo-Html -head $header
$message.Body = $null
$message.Body = $body
$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($message) 
}

else
{
$smtpServer = "smtp"
$smtpFrom = "Exchange@domain.com"
$smtpTo = "admin1@domain.com,admin2_Exchange@dmain.com"
$messageSubject = "Different error message that is not blocked"
$message = New-Object System.Net.Mail.MailMessage $smtpfrom, $smtpto
$message.Subject = $messageSubject
$message.IsBodyHTML = $true
$message.Body = $null
$action = "New Status - Check manually"
$temp.Action = $action
$body =  $temp | select "Status","Action" | ConvertTo-Html -head $header
$message.Body = $body 
$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($message)
}
}

catch
{
$ErrorMessage = $_.Exception.Message
$smtpServer = "smtp"
$smtpFrom = "Exchange@domain.com"
$smtpTo = "admin1@domain.com,admin2_Exchange@dmain.com"
$messageSubject = "Working as expected - The url is requesting credentials"
$message = New-Object System.Net.Mail.MailMessage $smtpfrom, $smtpto
$message.Subject = $messageSubject
$message.IsBodyHTML = $true
$message.Body = $null
$temp = "" | select Status,Action
$temp.Status = $ErrorMessage 
$action = "No action required"
$temp.Action = $Action
$body =  $temp | select "Status","Action" | ConvertTo-Html -head $header
$message.Body = $body
$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($message)
}

start-sleep -seconds 3600
$start = get-date
$end = $start.adddays(+7)
}until($start -ge $end)