param
(
    [Parameter(Mandatory)]$OldSAM,
    [Parameter(Mandatory)]$NewSAM
)
$UserEmail = get-aduser -Identity $OldSAM  -Properties * | Select-Object -ExpandProperty mail
$tentativeSAM = $NewSAM
if (Get-ADUser -Filter { SamAccountName -eq $NewSAM }) {    
    do {
        $inc ++
        $NewSAM = $tentativeSAM + [string]$inc
    } 
    until (-not (Get-ADUser -Filter { SamAccountName -eq $NewSAM }))
}
$Mainemail = "SMTP:" + $UserEmail
if ($UserEmail -like '*@scotcourtstribunals.gov.uk*') { 
    $NewEmail1 = "$NewSAM@scotcourtstribunals.gov.uk"
    $NewEmail2 = "$NewSAM@scotcourts.gov.uk"
    $NewEmail3 = "$NewSAM@scotcourts.pnn.gov.uk"
    $NewEmail4 = "$NewSAM@scotcourtstribunals.pnn.gov.uk"
}
elseif ($UserEmail -like '*@scotcourts.gov.uk*') { 
    $NewEmail1 = "$NewSAM@scotcourts.gov.uk"
    $NewEmail2 = "$NewSAM@scotcourtstribunals.gov.uk"
    $NewEmail3 = "$NewSAM@scotcourts.pnn.gov.uk"
    $NewEmail4 = "$NewSAM@scotcourtstribunals.pnn.gov.uk"
}
Write-Output "Changing to $NewEmail1"
Set-ADUser -Identity $OldSAM -emailaddress $NewEmail1
Set-ADUser -Identity $OldSAM -remove @{ProxyAddresses = $Mainemail }
$newProxy1 = "smtp:" + $NewEmail2
$newProxy2 = "smtp:" + $NewEmail3
$newProxy3 = "smtp:" + $NewEmail4
$NewPrimary = "SMTP:" + $NewEmail1
$newProxy4 = "smtp:" + $UserEmail
Set-ADUser -identity $OldSAM -Add @{proxyAddresses = ($newProxy1) }
Set-ADUser -identity $OldSAM -Add @{proxyAddresses = ($newProxy2) }
Set-ADUser -identity $OldSAM -Add @{proxyAddresses = ($newProxy3) }
Set-ADUser -identity $OldSAM -Add @{proxyAddresses = ($newProxy4) }
Set-ADUser -identity $OldSAM -Add @{proxyAddresses = ($NewPrimary) }