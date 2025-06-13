<#
.SYNOPSIS
Duplicate file finder

.NOTES
Script written by Brian Stark
Date:
Reviewed by:
Date:

Stored in Project Repo:

.DESCRIPTION
written by BStark
#>
#Grab Input Directories
$Input_1 = $args[0]
$Input_2 = $args[1]

#Create Log
$Output = "C:\PS\Log.csv"
Add-Content $Output 'Original,Status,Duplicate'

#Directory Check
$ExtChk_1 = [System.IO.Path]::GetExtension($Input_1)
$ExtChk_2 = [System.IO.Path]::GetExtension($Input_2)

#Check if inputs are directories
If ($ExtChk_1 -eq '' -And $ExtChk_2 -eq '') {
    #Grab list of files from each input path
    $Files_1 = Get-ChildItem -path $Input_1 -include $Off_Array -recurse -file
    $Files_2 = Get-ChildItem -path $Input_2 -include $Off_Array -recurse -file
    #Grab next file from each input path
    ForEach ($File_1 in $Files_1) {
        $FileCheck = 0
        $Hash_1 = Get-FileHash $File_1.FullName
        ForEach ($File_2 in $Files_2) {
            $Hash_2 = Get-FileHash $File_2.FullName
            #Compare file hashes
            If ($Hash_1.Hash -eq $Hash_2.Hash) {
                #Output results to csv
                If ($FileCheck -eq 0) {
                    $Filecheck = 1
                    $Text = $File_1.FullName + ',' + 'duplicate' + ',' + $File_2.FullName
                    Add-Content $Output $Text
                }
                else {
                    $Text = ',' + 'duplicate' + ',' + $File_2.FullName
                    Add-Content $Output $Text 
                }
            }
        }     
    }
}