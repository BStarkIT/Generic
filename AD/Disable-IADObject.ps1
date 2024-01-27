filter Disable-IADObject {
  if ($_ -is [ADSI] -and $_.psbase.SchemaClassName -match '^(user|computer)$') {
    $null = $_.psbase.invokeSet("AccountDisabled", $true)
    $null = $_.SetInfo()
    $_
  }
  else {
    Write-Warning "Invalid object type. Only 'User' or 'Computer' objects are allowed."
  }
}