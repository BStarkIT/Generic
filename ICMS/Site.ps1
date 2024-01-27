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
$Site = "Test"
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
$file = "\\scotcourts.local\data\IT\Civil Lab\00. Civil Lab Overall Delivery\Entitlement Attestation\ICMS\Completed Audits\$Site.xlsx"
$group1 = "ICMS $Site Judiciary"
$group2 = "ICMS $Site Office Manager"
$group3 = "ICMS $Site Admin Clerk"
$group4 = "ICMS $Site Clerk of Court"
$group5 = "ICMS $Site Adoption Admin Clerk"
$tab1 = "Judiciary"
$tab2 = "Office Manager"
$tab3 = "Admin Clerk"
$tab4 = "Clerk of Court"
$tab5 = "Adoption Admin Clerk"
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)
$sheet1 = $workbook.Worksheets.Item($tab1)
$sheet2 = $workbook.Worksheets.Item($tab2)
$sheet3 = $workbook.Worksheets.Item($tab3)
$sheet4 = $workbook.Worksheets.Item($tab4)
$sheet5 = $workbook.Worksheets.Item($tab5)
$objExcel.Visible = $false
$rowMax1 = ($sheet1.UsedRange.Rows).count
$rowMax2 = ($sheet2.UsedRange.Rows).count
$rowMax3 = ($sheet3.UsedRange.Rows).count
$rowMax4 = ($sheet4.UsedRange.Rows).count
$rowMax5 = ($sheet5.UsedRange.Rows).count
$rowName1, $colName1 = 1, 4
$rowName2, $colName2 = 1, 4
$rowName3, $colName3 = 1, 4
$rowName4, $colName4 = 1, 4
$rowName5, $colName5 = 1, 4
for ($i1 = 1; $i1 -le $rowMax1 - 1; $i1++) {
    $name1 = $sheet1.Cells.Item($rowName1 + $i1, $colName1).text
    $users1 += $name1
}
for ($i2 = 1; $i2 -le $rowMax2 - 1; $i2++) {
    $name2 = $sheet2.Cells.Item($rowName2 + $i2, $colName2).text
    $users2 += $name2
}
for ($i3 = 1; $i3 -le $rowMax3 - 1; $i3++) {
    $name3 = $sheet3.Cells.Item($rowName3 + $i3, $colName3).text
    $users3 += $name3
}
for ($i4 = 1; $i4 -le $rowMax4 - 1; $i4++) {
    $name4 = $sheet4.Cells.Item($rowName4 + $i4, $colName4).text
    $users4 += $name4
}
for ($i5 = 1; $i5 -le $rowMax5 - 1; $i5++) {
    $name5 = $sheet5.Cells.Item($rowName5 + $i5, $colName5).text
    $users5 += $name5
}
Write-Host "Checking group $group1 for unapproved members" 
ForEach ($CurrentUser in (Get-ADGroupMember $group1 | Select-Object -exp samaccountname)) {
    if ($CurrentUser -NotIn $users1) {
        Write-Output "user $CurrentUser missing from approved list, removing from group"
        Remove-ADGroupMember -Identity $group1 -Member $CurrentUser
    }  
}
$GroupMembers1 = get-ADGroupMember $group1 | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
foreach ($user in $users1) {
    if ($user -Notin $GroupMembers1) {
        Add-ADGroupMember -Identity $group1 -Members $user
        Write-Output "$user added to $group1"
    } 
}
Write-Host "Checking group $group2 for unapproved members" 
ForEach ($CurrentUser in (Get-ADGroupMember $group2 | Select-Object -exp samaccountname)) {
    if ($CurrentUser -NotIn $users2) {
        Write-Output "user $CurrentUser missing from approved list, removing from group"
        Remove-ADGroupMember -Identity $group2 -Member $CurrentUser
    }  
}
$GroupMembers2 = get-ADGroupMember $group2 | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
foreach ($user in $users2) {
    if ($user -Notin $GroupMembers2) {
        Add-ADGroupMember -Identity $group2 -Members $user
        Write-Output "$user added to $group2"
    } 
}
Write-Host "Checking group $group3 for unapproved members" 
ForEach ($CurrentUser in (Get-ADGroupMember $group3 | Select-Object -exp samaccountname)) {
    if ($CurrentUser -NotIn $users3) {
        Write-Output "user $CurrentUser missing from approved list, removing from group"
        Remove-ADGroupMember -Identity $group3 -Member $CurrentUser
    }  
}
$GroupMembers3 = get-ADGroupMember $group3 | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
foreach ($user in $users3) {
    if ($user -Notin $GroupMembers3) {
        Add-ADGroupMember -Identity $group3 -Members $user
        Write-Output "$user added to $group3"
    } 
}
Write-Host "Checking group $group4 for unapproved members" 
ForEach ($CurrentUser in (Get-ADGroupMember $group4 | Select-Object -exp samaccountname)) {
    if ($CurrentUser -NotIn $users4) {
        Write-Output "user $CurrentUser missing from approved list, removing from group"
        Remove-ADGroupMember -Identity $group4 -Member $CurrentUser
    }  
}
$GroupMembers4 = get-ADGroupMember $group4 | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
foreach ($user in $users4) {
    if ($user -Notin $GroupMembers4) {
        Add-ADGroupMember -Identity $group4 -Members $user
        Write-Output "$user added to $group4"
    } 
}

Write-Host "Checking group $group5 for unapproved members" 
ForEach ($CurrentUser in (Get-ADGroupMember $group5 | Select-Object -exp samaccountname)) {
    if ($CurrentUser -NotIn $users5) {
        Write-Output "user $CurrentUser missing from approved list, removing from group"
        Remove-ADGroupMember -Identity $group5 -Member $CurrentUser
    }  
}
$GroupMembers5 = get-ADGroupMember $group5 | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
foreach ($user in $users5) {
    if ($user -Notin $GroupMembers5) {
        Add-ADGroupMember -Identity $group5 -Members $user
        Write-Output "$user added to $group5"
    } 
}
$objExcel.quit()