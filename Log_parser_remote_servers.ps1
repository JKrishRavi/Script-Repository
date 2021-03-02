[string[]] $Servers =@()
$Servers = read-host "Enter server names"
$Servers = $Servers.Split(',')
$Filepath = Read-Host "Enternetwork file path start from the drive letter followed by $ sign"
$starttime = Read-Host "Enter start time in mm/dd/yyyy hh:mm:ss format"
$endtime = Read-Host "Enter end time in mm/dd/yyyy hh:mm:ss format"
$path_of_files = @()
$FPs =@()




foreach($Server in $Servers)
{
$FPs += "\\" + $Server + "\" + $Filepath
}
foreach($FP in $FPs)
{
$filtered = $FP | Get-ChildItem | ?{$_.Lastwritetime -ge $starttime } | select -expandproperty name

foreach($filtere in $filtered)
{
$filtered_total = $FP + "\" +  $filtere
$path_of_files += $filtered_total
}
}
$full = @()
[string[]] $Keys =@()
$Keys = read-host "Enter keywords"
$Keys = $Keys.Split(',')

Foreach($path_of_file in $path_of_files)
{

foreach($key in $keys)
{

$temp = $null
$temp =select-string -Path $path_of_file -Pattern $Key
$full += $temp
}
}

$full | select line,pattern |export-csv parsedoutput.csv -notypeinformation