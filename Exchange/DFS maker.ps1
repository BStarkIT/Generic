$Name = Read-host 'DFS Name?'
$Read = "ACL_" + $Name + "_R"
$RWA = "ACL_" + $Name + "_RWA"
$DFSRead = "ACL_DFS_DATA" + $Name + "_R"
$DFSRWA = "ACL_DFS_DATA" + $Name + "_RWA"
New-Item -Path "\\saufs01\F$\" -name "$Name" -ItemType Directory
new-ADGroup -Name $Read -GroupCategory Security -GroupScope DomainLocal -Path "OU=File Server ACL Groups,OU=Domain Local,OU=Groups,OU=SCTS,DC=scotcourts,DC=local" -Description "Read only access to \\scotcourts.local\data\$Name"
new-ADGroup -Name $RWA -GroupCategory Security -GroupScope DomainLocal -Path "OU=File Server ACL Groups,OU=Domain Local,OU=Groups,OU=SCTS,DC=scotcourts,DC=local" -Description "Read write administative access to \\scotcourts.local\data\$Name"
new-ADGroup -Name $DFSRead -GroupCategory Security -GroupScope DomainLocal -Path "OU=DFS Data (N Drive) ACL,OU=Domain Local,OU=Groups,OU=SCTS,DC=scotcourts,DC=local" -Description "Read only access to \\scotcourts.local\data\$Name"
new-ADGroup -Name $DFSRWA -GroupCategory Security -GroupScope DomainLocal -Path "OU=DFS Data (N Drive) ACL,OU=Domain Local,OU=Groups,OU=SCTS,DC=scotcourts,DC=local" -Description "Read write administative access to \\scotcourts.local\data\$Name"
$NewAcl = Get-Acl -Path "\\saufs01\F$\$Name"
$identity = "$RWA"
$identity1 = "$Read"
$fileSystemRights = "FullControl"
$fileSystemRights1 = "Read"
$type = "Allow"
$fileSystemAccessRuleArgumentList = $identity, $fileSystemRights, $type
$fileSystemAccessRuleArgumentList1 = $identity1, $fileSystemRights1, $type
$fileSystemAccessRule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $fileSystemAccessRuleArgumentList
$fileSystemAccessRule1 = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $fileSystemAccessRuleArgumentList1
$NewAcl.SetAccessRule($fileSystemAccessRule)
$NewAcl1.SetAccessRule($fileSystemAccessRule1)
Set-Acl -Path "\\saufs01\F$\$Name" -AclObject $NewAcl
Set-Acl -Path "\\saufs01\F$\$Name" -AclObject $NewAcl1