# Import grouplist from file. One group per line.
$Groups = Get-Content C:\Powershell\Groups\groups.csv
# Iterate through each group.
$Output = @(foreach ($Group in $Groups) {
        # Identify the individual members of the group.
        $Members = Get-ADGroupMember -Identity $Group
        # Identify the username, fullname of each member of the group.
        foreach ($User in $Members) {
            [pscustomobject]@{
                GroupName = $Group
                User      = $User.SamAccountName
                FullName  = $User.Name
            }
        }
    })
# Output results (contained in $Output) to a CSV file.
$Output | Export-CSV C:\Powershell\Groups\GroupMembers.csv -NoTypeInformation