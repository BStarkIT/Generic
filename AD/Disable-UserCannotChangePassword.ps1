
filter Disable-UserCannotChangePassword {
    if ($_ -is [ADSI] -and $_.psbase.SchemaClassName -eq 'user') {
        $acl = $_.psbase.ObjectSecurity
        $deny = $acl.GetAccessRules($true, $false, [System.Security.Principal.NTAccount]) | `
            Where-Object { ($_.IdentityReference -eq 'Everyone' -or $_.IdentityReference -eq 'NT AUTHORITY\SELF') `
                -and $_.AccessControlType -eq 'Deny' -and $_.ActiveDirectoryRights -eq 'ExtendedRight' } 
        if ($deny) {
            $deny | Foreach-Object { $null = $acl.psbase.RemoveAccessRule($_) }
            $_.psbase.CommitChanges()
            $_
        }
    }
    else {
        Write-Warning "Invalid object type, only User objects are allowed"
    } 
} 