$dep = read-host "Enter department name"

$full = @()
$temp1= $null
$temp2 = $null
$temp3 = $null
$users = $null
$users = $null
$temp1 = Get-ADUser -Filter {department -like $dep} -server admin | select -ExpandProperty Name
$temp2 = Get-ADUser -Filter {department -like $dep} -server branch | select -ExpandProperty Name
$temp3 = Get-ADUser -Filter {department -like $dep} -server corp | select -ExpandProperty Name
$users = $temp1 + $temp2 + $temp3

$fn = $dep -replace (' ')
foreach($user in $users)
{

if(get-mailbox $user -ErrorAction SilentlyContinue)
{
$temp = get-mailbox $user | select alias,@{n="department";e={$dep}}
$full += $temp
}
}

$full | export-csv mailboxthatbelongto$fn.csv -notypeinformation