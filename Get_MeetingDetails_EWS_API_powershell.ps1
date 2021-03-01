

function Get-CalInforUsingEWSpowershell
{
     [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        $Identity,
  
        [Parameter()]
        $Days,
  
        [Parameter()]
        [System.Management.Automation.CredentialAttribute()]
        [pscredential]
        $Credential
    )
     
    begin
    {
        Import-Module 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
    }
     
    process
    {
        $mbx = get-mailbox $Identity

        $Service = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2)
        if($PSBoundParameters.ContainsKey('Credential'))
        {
            $Service.Credentials = [System.Net.NetworkCredential]::new($Credential.UserName,$Credential.Password)
        }
        else
        {
            $Service.UseDefaultCredentials = $true
        }
        $Service.Url = "https://access.citizensbank.com/EWS/Exchange.asmx"
        $Service.AutoDiscoverUrl($mbx.primarysmtpaddress)
        
        $SMTP = (Get-mailbox $Identity).primarysmtpaddress
        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$SMTP)
  

        $Folder = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($Service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
        $View = [Microsoft.Exchange.WebServices.Data.CalendarView]::new([datetime]::Now,[datetime]::Now.AddDays($Days))
       
       
  
 
        $frCalendarResult = $Folder.FindAppointments($View) 
        $objectall = @()
        foreach ($apApointment in $frCalendarResult.Items){
             $obj = New-Object -TypeName psobject
             $obj | Add-Member -MemberType NoteProperty -Name "Meeting Room Name" -Value $a.Displayname
             $obj | Add-Member -MemberType NoteProperty -Name "Meeting Room Capacity" -Value $a.ResourceCapacity
             $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
             $apApointment.load($psPropset)
             $obj = New-Object -TypeName psobject
             Write-host "`n"
             Write-host "***"
             Write-host "`n"
             
             "Appointment : " + $apApointment.Subject.ToString()
             $obj | Add-Member -MemberType NoteProperty -Name Appointment -Value $apApointment.Subject.ToString()
             
             $apApointmentDuration = $apApointment.End - $apApointment.Start
             "Duration of the meeting in minutes : " + $apApointmentDuration.TotalMinutes
             $obj | Add-Member -MemberType NoteProperty -Name "Duration in minutes" -Value $apApointmentDuration.TotalMinutes

             "Organizer : " + $apApointment.Organizer.ToString()
             $obj | Add-Member -MemberType NoteProperty -Name "Organizer" -Value $apApointment.Organizer.ToString()
             
             "Required Attendees :"
              $countrequired = 0
             foreach($attendee in $apApointment.RequiredAttendees){
                          " " + $attendee.Address
                     $countrequired += 1
                    }
              if($countrequired -ne 0)
              {
              $obj | Add-Member -MemberType NoteProperty -Name "Required Attendees" -Value $countrequired
              }
              else
              {
              " 0 Attendees "
              $obj | Add-Member -MemberType NoteProperty -Name "Required Attendees" -Value $countrequired
              }
              $countoptional = 0
             "Optional Attendees :"
             foreach($attendee in $apApointment.OptionalAttendees){
                          " " + $attendee.Address
                     $countoptional += 1
             }
              if($countoptional -ne 0)
              {
              $obj | Add-Member -MemberType NoteProperty -Name "Optional Attendees" -Value $countoptional
              }
              else
              {
              " 0 Attendees "
              $obj | Add-Member -MemberType NoteProperty -Name "Optional Attendees" -Value $countoptional
              }
             
             $count = $countrequired + $countoptional
             "Total Count of Recipients - includes required and optional : " + $count
             $obj | Add-Member -MemberType NoteProperty -Name "Total Count of Recipients - includes required and optional" -Value $count
             $objectall = $obj+$objectall
             return $objectall
             

         } 
$objectall | export-csv export$a.csv -notypeinformation
    }
     
    end
    {
     
    }
}

$d = Read-host "Enter the number of days you want the Details for"
$cred = get-credential

$rooms = import-csv room.csv | select -ExpandProperty rooms

foreach($room in $rooms)
{
Write-host "===================================================================="
$a = get-mailbox $room
write-host "`n Meeting Room Name : "  $a.Displayname

write-host "`n Meeting Room Capacity : "  $a.ResourceCapacity

Get-CalInforUsingEWSpowershell -Identity $room -Days $d -Credential $cred

}
$objectall | export-csv ewspowershellexport.csv -notypeinformation


