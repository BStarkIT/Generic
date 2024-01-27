<#
.DESCRIPTION
Script to remove Users from ICMS security group if not on excel sheet
.FUNCTION
Internal Audit (Fraud Prevention and Asset Management) required a review of permissions granted on non-financial systems.  

Lansweeper was used to generate a user report for each court and Sheriff Clerks asked to review the permissions.  

26 of 42 responses have been received.  So far approx. 2,000 user permisions needs to be removed.  

We have security groups from AD 'ICMSGroupName' and 'Username' in Excel spearesheets.

We are looking to run a script to remove the obsolete permissions e.g. remove users from AD groups.

This exercise should be run at least annually.
.INPUTS
Excel file of users to be located C:\PS\ICMS.xls
colum to be called 'Username' to contain the SAM of the users on sheet1 of the excel file
.VERSION
1.0
.DATE
13/08/2021
.HISTORY
13/08/2021  BStark  1.0 Inital request.

#>

$Site = "Wick"


$users1 = @()
[System.Collections.ArrayList]$ArrayList = $users1
$ArrayList.GetType()
$users2 = @()
[System.Collections.ArrayList]$ArrayList = $users2
$ArrayList.GetType()
$users3 = @()
[System.Collections.ArrayList]$ArrayList = $users3
$ArrayList.GetType()
$users4 = @()
[System.Collections.ArrayList]$ArrayList = $users4
$ArrayList.GetType()
$users5 = @()
[System.Collections.ArrayList]$ArrayList = $users5
$ArrayList.GetType()

$file = "N:\IT\Civil Lab\00. Civil Lab Overall Delivery\Entitlement Attestation\ICMS\Completed Audits\reformatted\$Site.xlsx"
$sheetName1 = "Admin Clerk" 
$group1 = "ICMS $Site Admin Clerk"
$sheetName2 = "Adoption Admin Clerk" 
$group2 = "ICMS $Site Adoption Admin Clerk"
$sheetName3 = "Clerk of Court" 
$group3 = "ICMS $Site Clerk of Court"
$sheetName4 = "Judiciary" 
$group4 = "ICMS $Site Judiciary"
$sheetName5 = "Office Manager" 
$group5 = "ICMS $Site Office Manager"
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)
$sheet1 = $workbook.Worksheets.Item($sheetName1)
$sheet2 = $workbook.Worksheets.Item($sheetName2)
$sheet3 = $workbook.Worksheets.Item($sheetName3)
$sheet4 = $workbook.Worksheets.Item($sheetName4)
$sheet5 = $workbook.Worksheets.Item($sheetName5)
$objExcel.Visible = $false
$rowMax1 = ($sheet1.UsedRange.Rows).count
$rowMax2 = ($sheet2.UsedRange.Rows).count
$rowMax3 = ($sheet3.UsedRange.Rows).count
$rowMax4 = ($sheet4.UsedRange.Rows).count
$rowMax5 = ($sheet5.UsedRange.Rows).count
$rowName, $colName = 1, 4
for ($i = 1; $i -le $rowMax1 - 1; $i++) {
    $name1 = $sheet1.Cells.Item($rowName + $i, $colName).text
    $users1 += $name1
}
for ($i = 1; $i -le $rowMax2 - 1; $i++) {
    $name2 = $sheet2.Cells.Item($rowName + $i, $colName).text
    $users2 += $name2
}
for ($i = 1; $i -le $rowMax3 - 1; $i++) {
    $name = $sheet3.Cells.Item($rowName + $i, $colName).text
    $users3 += $name
}
for ($i = 1; $i -le $rowMax4 - 1; $i++) {
    $name = $sheet4.Cells.Item($rowName + $i, $colName).text
    $users4 += $name
}
for ($i = 1; $i -le $rowMax5 - 1; $i++) {
    $name = $sheet5.Cells.Item($rowName + $i, $colName).text
    $users5 += $name
}
Write-Host "Checking group $group1 for unapproved members" 
ForEach ($CurrentUser in (Get-ADGroupMember $group1 | Select-Object -exp samaccountname)) {
    if ($CurrentUser -NotIn $users1) {
        Write-Output "user $CurrentUser Missing, removing from group"
        Remove-ADGroupMember -Identity $group1 -Members $CurrentUser
    }  
}
$GroupMembers = get-ADGroupMember $group1 | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
foreach ($user in $users1) {
    if ($user -ne "") {
        if ($user -Notin $GroupMembers) {
            Write-Output "$user added to $group1"
            Add-ADGroupMember -Identity $group1 -Members $CurrentUser
        } 
    }
}
Write-Host "Checking group $group2 for unapproved members" 
ForEach ($CurrentUser in (Get-ADGroupMember $group2 | Select-Object -exp samaccountname)) {
    if ($CurrentUser -NotIn $users2) {
        Write-Output "user $CurrentUser Missing, removing from group"
        Remove-ADGroupMember -Identity $group2 -Members $CurrentUser
    }  
}
$GroupMembers2 = get-ADGroupMember $group2 | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
foreach ($user in $users2) {
    if ($user -ne "") {
        if ($user -Notin $GroupMembers2) {
            Write-Output "$user added to $group2"
            Add-ADGroupMember -Identity $group2 -Members $CurrentUser
        } 
    }
}
Write-Host "Checking group $group3 for unapproved members" 
ForEach ($CurrentUser in (Get-ADGroupMember $group3 | Select-Object -exp samaccountname)) {
    if ($CurrentUser -NotIn $users3) {
        Write-Output "user $CurrentUser Missing, removing from group"
        Remove-ADGroupMember -Identity $group3 -Members $CurrentUser
    }  
}
$GroupMembers3 = get-ADGroupMember $group3 | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
foreach ($user in $users3) {
    if ($user -ne "") {
        if ($user -Notin $GroupMembers3) {
            Write-Output "$user added to $group3"
            Add-ADGroupMember -Identity $group3 -Members $CurrentUser
        }
    } 
}
Write-Host "Checking group $group4 for unapproved members" 
ForEach ($CurrentUser in (Get-ADGroupMember $group4 | Select-Object -exp samaccountname)) {
    if ($CurrentUser -NotIn $users4) {
        Write-Output "user $CurrentUser Missing, removing from group"
        Remove-ADGroupMember -Identity $group4 -Members $CurrentUser
    }  
}
$GroupMembers4 = get-ADGroupMember $group4 | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
foreach ($user in $users4) {
    if ($user -ne "") {
        if ($user -Notin $GroupMembers4) {
            Write-Output "$user added to $group4"
            Add-ADGroupMember -Identity $group4 -Members $CurrentUser
        } 
    }
}
Write-Host "Checking group $group5 for unapproved members" 
ForEach ($CurrentUser in (Get-ADGroupMember $group5 | Select-Object -exp samaccountname)) {
    if ($CurrentUser -NotIn $users5) {
        Write-Output "user $CurrentUser Missing, removing from group"
        Remove-ADGroupMember -Identity $group5 -Members $CurrentUser -Confirm:$false
    }  
}
$GroupMembers5 = get-ADGroupMember $group5 | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
foreach ($user in $users5) {
    if ($user -ne "") {
        if ($user -Notin $GroupMembers5) {
            Write-Output "$user added to $group5"
            Add-ADGroupMember -Identity $group5 -Members $CurrentUser
        } 
    }
    
}

$objExcel.quit()
