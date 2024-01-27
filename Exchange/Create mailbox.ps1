param
(
    [Parameter(Mandatory)]$tentativeSAM
)
$sam = $tentativeSAM
if (Get-ADUser -Filter { SamAccountName -eq $tentativeSAM }) {    
    do {
        $inc ++
        $tentativeSAM = $sam + [string]$inc
    } 
    until (-not (Get-ADUser -Filter { SamAccountName -eq $tentativeSAM }))
}
Enable-MailBox -Identity $sam@scotcourts.gov.uk