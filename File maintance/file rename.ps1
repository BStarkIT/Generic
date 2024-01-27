<#
.SYNOPSIS
This PowerShell script is to batch rename files ADDING text within the filename.

.DESCRIPTION
written by BStark

rename AAABCCC.file to AAABXXXBCCC.file
file-123.txt -> file-old-123.txt
use: $Newname = $Prefix + $cutpoint + $add + $cutpoint + $surfix
OR
rename AAABCCC.file to AAAXXXBCCC.file
file-123.txt -> fileold-123.txt
use: $Newname = $Prefix + $add + $cutpoint + $surfix

.NOTES
Script written by Brian Stark
#>
$Path = "C:\PS\Test"                                # change to pairent location
$add = "Name"                                       # test to be added to file
$cutpoint = "_"                                     # where to cut the name
$Batch = Get-ChildItem -Path $Path -Recurse -File   # remove -file if you want folders to be renamed as well
foreach ($currentItemName in $Batch) {
    $Filename = $currentItemName.Name
    if ($currentItemName.Name -match $cutpoint) {
        $Prefix = $Filename.Split($cutpoint )[0]
        $surfix = $Filename.Split($cutpoint )[-1]
        $Newname = $Prefix + $cutpoint + $add + $cutpoint + $surfix
        Rename-Item -path $currentItemName -NewName $Newname
        Write-Output "$currentItemName renamed $Newname"
    }
    else {
        # when the cut marker not in the file name
        Write-Output "$currentItemName not changed"
    }
}