# Install all AR Fonts
Write-Host "Installing Fonts"
$FONTS = 0x14
$Path="c:\fonts"
$objShell = New-Object -ComObject Shell.Application
$objFolder = $objShell.Namespace($FONTS)
New-Item $Path -type directory
Copy-Item "\\filestore419.ad.realise.com\brians$\Builds\Fonts\*" $Path
$Fontdir = dir $Path
foreach($File in $Fontdir) {
  $objFolder.CopyHere($File.fullname, 16)
}
remove-item $Path -recurse


# Installer, copy files to local machine then install in batch


New-Item -ItemType directory -Path "C:\temp"


    Copy-Item "\\Filestore419.ad.realise.com\brians$\Builds\*.exe" "C:\temp\" -Recurse
    Copy-Item "\\Filestore419.ad.realise.com\brians$\Builds\*.txt" "C:\temp\" -Recurse
    Copy-Item "\\Filestore419.ad.realise.com\brians$\Builds\AR VPN.MSI" "C:\temp\" -Recurse
# initialize the items variable with the
# contents of a directory defined in $Dir
$Dir="c:\temp"
$items = Get-ChildItem -Path $Dir


# enumerate the items array
foreach ($item in $items)
{
     # if the item is NOT a directory and it's an executable, then process it.
     if ( ($item.Attributes -ne "Directory") -and ($item.Extension -eq ".exe") )
     {
           $arguments = "/q"                                       # Set default arguments list
           $Base = $item.BaseName                                  # Get filename without extension
           Write-Host "Installing $Base"
           if( Test-Path "$Dir\$Base.txt" )                        # Is there an associated text file?
           {
               Write-Host "Reading arguments from $Dir\$Base.txt"
               $arguments = Get-Content "$Dir\$Base.txt" -First 1  # Read arguments from file.
           }


           if(! $arguments)
           {
               Write-Host "Running $item.Fullname -Wait"
               Invoke-Command -ScriptBlock { Start-Process $item.FullName -Wait }
           } else {
               Write-Host "Running $item.Fullname -Argumentlist $arguments -Wait"
               Invoke-Command -ScriptBlock { Start-Process $item.FullName -ArgumentList $arguments -Wait }
           }
     }
}
    Write-Host "Installing VPN"
    Invoke-Command -ScriptBlock {Start-Process "c:\temp\AR VPN.MSI" -ArgumentList "/q" -Wait} 
    Write-Host "Adding Printer"
    Add-Printer -ConnectionName \\steiger.ad.realise.com\FOLLOW_ME


# Remove temporary files and folder


    Write-Host "Removing Temporary files"
    $RemovalPath = "c:\temp\"
    Get-ChildItem  -Path $RemovalPath -Recurse  | Remove-Item -Force -Recurse
    Remove-Item $RemovalPath -Force -Recurse
    #
    Write-Host "Applying Group Policies"
    gpupdate /force
    gpupdate /force
    Write-Host "Complete"

