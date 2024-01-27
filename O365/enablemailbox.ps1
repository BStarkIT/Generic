$Problem = 'RRay', 'dprebish', 'slowson', 'smorrison'
foreach ($sam in $Problem) {
    $name = Get-ADUser $sam | Select-Object -ExpandProperty Name
    $FirstName = Get-ADUser $sam | Select-Object -ExpandProperty GivenName
    $Lastname = Get-ADUser $sam | Select-Object -ExpandProperty Surname
    $Ticket = "205210"
    Set-ADUser -Identity $SAM -add @{"extensionattribute2" = "USR-PERS" }
    Set-ADUser -Identity $SAM -add @{"extensionattribute3" = $Ticket }
    $route = "$SAM@scotcourtsgovuk.mail.onmicrosoft.com"
    $newProxy = "smtp:" + "$SAM@scotcourtstribunals.gov.uk"
    $NewPrimary = "SMTP:" + "$SAM@scotcourts.gov.uk"
    $newProxy1 = "smtp:" + "$SAM@scotcourts.pnn.gov.uk"
    $newProxy2 = "smtp:" + "$SAM@scotcourtstribunals.pnn.gov.uk"
    $newProxy3 = "smtp:" + "$SAM@scotcourtsgovuk.mail.onmicrosoft.com"
    $newProxy4 = "X400:C=GB;A=CWMAIL;P=SCS;O=SCOTTISH COURTS;S=" + $Lastname + ";G=" + $FirstName + ";"
    Set-ADUser -identity $SAM -EmailAddress "$SAM@scotcourts.gov.uk"
    Set-ADUser -identity $SAM -replace @{proxyAddresses = ($NewPrimary) }
    Set-ADUser -identity $SAM -add @{proxyAddresses = ($newProxy) }
    Set-ADUser -identity $SAM -add @{proxyAddresses = ($newProxy1) }
    Set-ADUser -identity $SAM -add @{proxyAddresses = ($newProxy2) }
    Set-ADUser -identity $SAM -add @{proxyAddresses = ($newProxy3) }
    Set-ADUser -identity $SAM -add @{proxyAddresses = ($newProxy4) }
    Enable-RemoteMailbox $name -RemoteRoutingAddress $route
}