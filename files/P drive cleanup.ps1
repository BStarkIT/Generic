$Paths = Get-ChildItem H:\P |  Where-Object { $_.PSIsContainer } | Foreach-Object { $_.Name }
Foreach ($Path in $Paths) {
    Try {
        $User = Get-ADUser -Identity $Path | Select-Object Name
    }
    Catch {
        #Remove-Item -path "H:\P\$Path" -Force -Recurse
        Write-output $Path
    }
}