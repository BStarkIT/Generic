$Users = Get-ADUser -filter * -searchbase 'ou=Shared Mailboxes,ou=Resource Accounts,ou=scts,DC=scotcourts,DC=local' -Properties EmailAddress,SamAccountName | Select-Object EmailAddress,SamAccountName 
foreach ($User in $Users) {
    if ($User.EmailAddress -like "*@ScotCourtsTribunals.gov.uk") {
        #$SAM = $User -ireplace [regex]::Escape("@scotcourtstribunals.gov.uk"), ""
        Write-Host $User.SamAccountName
        Set-ADUser -Identity $User.SamAccountName -clear "extensionattribute2"
        Set-ADUser -Identity $User.SamAccountName -add @{"extensionattribute2"="MBT-SHAR"}
    }
    else {
        #$SAM = $User -ireplace [regex]::Escape("@scotcourts.gov.uk"), ""
        Write-Host $User.SamAccountName
        Set-ADUser -Identity $User.SamAccountName -clear "extensionattribute2"
        Set-ADUser -Identity $User.SamAccountName -add @{"extensionattribute2"="MBS-SHAR"}
    }
}
