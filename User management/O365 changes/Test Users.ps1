$Users = Get-ADUser -filter * -searchbase 'ou=User Accounts (Testing),ou=scts,DC=scotcourts,DC=local' -Properties EmailAddress | Select-Object EmailAddress | Select-Object -ExpandProperty EmailAddress
foreach ($User in $Users) {
    if ($User -like "*@ScotCourtsTribunals.gov.uk") {
        $SAM = $User -ireplace [regex]::Escape("@scotcourtstribunals.gov.uk"), ""
        Write-Output $SAM
        Set-ADUser -Identity $SAM -clear "extensionattribute2"
        Set-ADUser -Identity $SAM -add @{"extensionattribute2" = "USR-PERT" }
    }
    else {
        $SAM = $User -ireplace [regex]::Escape("@scotcourts.gov.uk"), ""
        Write-Output $SAM
        Set-ADUser -Identity $SAM -clear "extensionattribute2"
        Set-ADUser -Identity $SAM -add @{"extensionattribute2" = "USR-PERS" }
    }
}
