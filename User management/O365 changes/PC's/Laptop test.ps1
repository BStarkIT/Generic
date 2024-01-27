$User = Get-ADComputer  -Identity "LT024640"
        Set-ADComputer -Identity $User -clear "extensionattribute2"
        Set-ADComputer -Identity $User -add @{"extensionattribute2"="DVC-LAPT"}

