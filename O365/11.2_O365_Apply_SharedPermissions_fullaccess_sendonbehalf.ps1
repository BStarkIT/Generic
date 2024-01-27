#Using an input csv file, apply the full access and sendonbehalf access to users with their shared mailboxes
#sharedmailbox fullaccess0 fullaccess1 sendas0 sendas1 
#sharedmailbox@scotcourts.gov.uk jbloggs@scotcourts.gov.uk jbloggs2@scotcourts.gov.uk
#set-mailbox -identity sharedmbx@scotcourt.gov.uk -grantsendonbehalfto @{add="email1","emailuser2","emailuser3"}

function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
}
$inputs = Import-Csv "c:\scripts\o365\shared\daily_input_shared_28022022.csv"
$logfile = "C:\scripts\o365\shared\daily_input_shared_" + "$(Get-TimeStamp)" +".txt"
foreach ($inp in $inputs)
{
    $sharedid = $inp.mail 
    $facc =  $inp | Get-Member | Where-Object {$_.Name -like "full*"} | Select Definition
    foreach($fa in $facc)
    {
        $tempaccess = $($fa.Definition -Split "=")[1]
        if ($tempaccess.Length -gt 5)
        {
            Add-MailboxPermission -id $sharedid -User $tempaccess -AccessRights fullaccess -confirm:$false
            Write-Output "$(Get-TimeStamp) : Info : Full access provided to $tempaccess on $sharedid" | Out-File $logfile -Append
            Write-Output "$(Get-TimeStamp) : Info : Full access provided to $tempaccess on $sharedid"
        }
    }

     $sacc =  $inp | Get-Member | Where-Object {$_.Name -like "sendon*"} | Select Definition
    foreach($sa in $sacc)
    {
        $tempaccess = $($sa.Definition -Split "=")[1]
        if ($tempaccess.Length -gt 5)
        {
            set-mailbox -identity $sharedid -grantsendonbehalfto @{add="$tempaccess"}
            Write-Output "$(Get-TimeStamp) : Info : SendOnBehalf access provided to $tempaccess on $sharedid" | Out-File $logfile -Append
            Write-Output "$(Get-TimeStamp) : Info : SendOnBehalf access provided to $tempaccess on $sharedid"
        }
    }
}