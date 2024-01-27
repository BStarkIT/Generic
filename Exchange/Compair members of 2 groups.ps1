
Compare-Object (Get-ADGroupMember 'DSU Staff') (Get-ADGroupMember 'DigitalServicesUnit') -Property 'SamAccountName' -IncludeEqual | Export-csv -path C:\PS\DSUcompar.csv
