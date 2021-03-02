$root = New-Object System.DirectoryServices.DirectoryEntry("GC://RootDSE")
$entry = New-Object System.DirectoryServices.DirectoryEntry("GC://" + $root.get("rootDomainNamingContext"))
$search = New-Object System.DirectoryServices.DirectorySearcher
$search.SearchRoot = $entry
$search.Filter = $searchfilter
$search.PropertiesToLoad.Add("displayname")
$search.PropertiesToLoad.Add("samaccountname")
$search.SearchScope = "subtree"
$search.PageSize = 1000



$results = $search.FindAll()
$temp_results = @()
foreach ($result in $results) 
{

$check_variable = $result.properties["samaccountname"]

try
{
$check = get-remotemailbox ([string]$check_variable ) -erroraction stop
}

catch
{
$temp_results += $result
}

}

$grid = $null
foreach ($temp_result in $temp_results ) {
  	$grid += "<tr>"
  	$grid += "<td>" + $temp_result.properties["displayname"] + "</td>"
  	$grid += "<td>" + $temp_result.properties["samaccountname"] + "</td>"
  	$grid += "</tr>"
}

$smtp = New-Object Net.Mail.SmtpClient("smtp")
$msg = New-Object Net.Mail.MailMessage
$msg.From = "ad@domain.com"
$msg.To.Add("recipient@domain.com")
$msg.IsBodyHTML = 'true'
$msg.Subject = "External Mail Enabled AD Accounts"
$msg.Body =
 	"<div style='font-family:tahoma;font-size:10pt;'>
 	The following external mail enabled AD accounts have been detected:<br /><br />
 	<table border='1' cellpadding='5' cellspacing='2' style='font-size:10pt;'>
 	<tr style='background:#dddddd;font-weight:bold'><td>Account Name</td><td>Account ID</td></tr>
 	$grid
 	</table>
 	</div>"
$smtp.Send($msg)