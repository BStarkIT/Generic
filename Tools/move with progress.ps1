
$Extension = "*.mp4", "*.mkv", "*.avi", "*.mov", "*.wmv"
$paths = @(
    "I:\Films2",
    "O:\Films"
)
$drive = Get-PSDrive -Name T  # Replace 'Z' with your mapped drive letter

foreach ($path in $paths) {
    <# $path is the current item #>
    get-ChildItem $path -recurse -ErrorAction "SilentlyContinue" -include $Extension | 
    Where-Object { !($_.PSIsContainer) -and $_.Length / 1GB -gt 10 } | 
    ForEach-Object {
        $freeSpaceGB = [math]::Round($drive.Free / 1GB, 2)
        if ($freeSpaceGB -gt $_.Length / 1GB) {
            Write-Output "Free space on drive $($drive.Name): $freeSpaceGB GB"
            Write-Output "Moving file: $($_.fullname)"
            Move-Item $_.fullname T:\Handbrake-in
        } 
        else {
            Write-Output "Not enough free space on drive $($drive.Name): $freeSpaceGB GB"
            Write-Output "File too large to move: $($_.fullname)"
        }
        
    }
}