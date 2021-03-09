###
Place the Dl alias and the secondary addres to be removed under header "Alias" and "Address" Respectively in the csv file named "list.csv" in the current directory

###

$lists  = import-csv .\list.csv

foreach($list in $lists)
{
$Alias = $list.Alias
$Address = $list.Email
write-host "Removing emailaddress $Address from the DL $Alias"
Set-DistributionGroup "$Alias" -EmailAddresses @{remove=$Address}
}

