## place the recipient domains with prefix "*@" , it should be in the following format *@domain.com
$recipients = import-csv recipients.csv | select -expandproperty recipients
foreach($recipient in $recipients)
{
$traceresults = @()

#use $r = $recipient if you are searching for individual recipients and not domains
$r = $recipient.substring(2)
$r

$st = read-host "Enter start date in the following format MM/DD/YYYY"
$ed = read-host "Enter end date in the following format MM/DD/YYYY"

$page = 1
do
{
$page
$check = Get-EOPMessageTrace  -RecipientAddress $recipient -StartDate get-date($st) -EndDate get-date($ed) -Page $page -PageSize 5000
$checkcount = $check.count
$traceresults += $check
$checkcount
$page = $page + 1

}
until($checkcount -le 4999)

$traceresults | export-csv trace$r.csv -notypeinformation
}






