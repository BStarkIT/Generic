param
(
    [Parameter(Mandatory)]$Path
)
(get-acl $path).access | Format-Table IdentityReference, FileSystemRights, AccessControlType, IsInherited, InheritanceFlags -auto