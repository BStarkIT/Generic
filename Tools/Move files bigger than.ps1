get-ChildItem $path -recurse -ErrorAction "SilentlyContinue" -include $Extension | 
Where-Object { !($_.PSIsContainer) -and $_.Length / 1MB -gt 10240 } | 
ForEach-Object {
    Write-Output "$($_.fullname) $($_.Length / 1Gb)"
    Move-Item $_.fullname E:\Input
}
