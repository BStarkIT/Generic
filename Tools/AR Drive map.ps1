$CurrentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$Groups = $CurrentUser.Groups | Foreach-Object {
    $_.Translate([Security.Principal.NTAccount])
}
if ($Groups -contains "local.amazerealise.com\Editors") {
    New-PSDrive -name T -PSProvider FileSystem -Root \\Filestore419\Realise$ -persist
}