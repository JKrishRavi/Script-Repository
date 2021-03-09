###
Place the mailbox alias and the secondary addres to be removed under header "Alias" and "Email" Respectively in the csv file named "list.csv" in the current directory

###

$lists  = import-csv .\list.csv

foreach($list in $lists)
{
$A = $list.Alias
$B = $list.Email
write-host "Removing emailaddress $B from mailbox $A"
Set-Mailbox $A -EmailAddresses @{add="$B"}
}