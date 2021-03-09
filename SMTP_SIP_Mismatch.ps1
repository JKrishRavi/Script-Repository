######################################################

<#
Author: Jayakrishna Ravi
Version: 1.0
Date: 7/18/2019

.SYNOPSIS
get-smtpsipmismatch.Ps1 - Script to get smtp and sip address mismatch

#>


###########################################################
#Parameters
<#
$stringPrimaryAddress - Stores primary email address, converted to a string so that we can compare this with sip address
$stringSIP - Stores SIP address, converted to a string so that we can compare this with email address
$temp - temporarily stores mismatching addresses while going through the looping statements
$mismatches - stores the mismatching email and sip addresses with the respective alias
$e - Error capturing
#>
Try
{

[String] $stringPrimaryAddress
[String] $stringSIP
[Array] $mismatches = @()
$Usermailboxes = Get-mailbox -ResultSize unlimited | ?{$_.RecipientTypeDetails -match "UserMailbox"}
ForEach ($mailbox in $Usermailboxes){
$temp = "" | Select "Name", "Alias", "Primary SMTP Address", "SIP"
$stringPrimaryAddress = $NULL
$stringSIP = $NULL

$mailbox.EmailAddresses | ForEach{
If ($_.IsPrimaryAddress -and $_.Prefix -match "SMTP") 
{ 
$stringPrimaryAddress = $_.SmtpAddress 
}
If ($_.PrefixString -eq "sip") 
{ 
$stringSIP = $_.AddressString }
}

If (($stringPrimaryAddress -ne $stringSIP) -and ($stringSIP -ne $null ))
{
$temp."Name" = $mailbox.DisplayName
$temp."Alias" = $mailbox.alias
$temp."Primary SMTP Address" = $stringPrimaryAddress
$temp."SIP" = $stringSIP
$mismatches += $temp

}
}


######################
#HTML Report Style Input
$Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@

##################################
#Email Report Generation

$Emailbody = $NULL
 
$Emailbody = $mismatches | select-object "Name", "Alias", "Primary SMTP Address", "SIP" | ConvertTo-Html -head $Header


$smtpServer = "smtp"
$smtpFrom = "user@domain.com"
$smtpTo = "DL-grp@domain.com"
$messageSubject = "SMTP and SIP Address mismatch details"
$message = New-Object System.Net.Mail.MailMessage $smtpfrom, $smtpto
$message.Subject = $messageSubject
$message.IsBodyHTML = $true
$message.Body = $null
$message.Body = $Emailbody 
$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($message)



}

Catch

{
$time = Get-Date -Format MMddyyyyHHmm
Add-content "$time $_.Exception.Message" -Path Mismatch$time.txt
$attachment = "Mismatch$time.txt"
Send-MailMessage -From "user@domain.com" -To "DL-grp@domain.com" -Subject "SMTP and SIP address mismatch - ERROR" -Body "Please review Error Log in the attachment" -Attachments $attachment -SmtpServer "smtp"
}
