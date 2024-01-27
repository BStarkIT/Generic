$GroupObj = Get-AzureADGroup -Filter "DisplayName eq 'O365-MAIL'"
#Read group members from CSV file
$GroupMembers = Import-CSV "C:\scripts\O365\groups\update_group_members.csv"
$i = 0;
$TotalRows = $GroupMembers.Count
#Iterate members one by one and add to group
Foreach($GroupMember in $GroupMembers)
{
$User = $GroupMember.UserPrincipalName
Try
{
$UserObj = (Get-AzureADUser -ObjectId $GroupMember.UserPrincipalName).ObjectId
Add-AzureADGroupMember -ObjectId $GroupObj.ObjectId -RefObjectId $UserObj
Write-Host "Adding member" $User -ForegroundColor Green
}
catch
{
Write-Host "Error occurred for $User" -f Yellow
Write-Warning $Error[0]
}
}