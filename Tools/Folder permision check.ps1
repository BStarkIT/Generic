$path = read-Host 'Target'
(get-acl $path).access | Format-Table IdentityReference,FileSystemRights,AccessControlType,IsInherited,InheritanceFlags -auto