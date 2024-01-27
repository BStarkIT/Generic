#This section should be run with standard Admin account, please use below commented command to establish connection
#Ensure you are connected to Exchange online
#Run this script per section as the query takes time to create and run before execution

#Section1
#Connect-IPPSSession -Credential (Get-Credential)

$Sender = "gkhosa_test@domain.com.com"
$Subject = "check phish test april"

#Assign Start date and End date in same format (yyyy-mm-dd)
#https://compliance.microsoft.com/contentsearchv2?viewid=search

$StartDate = "2023-04-24"
$EndDate = "2023-04-26"
$date = $StartDate + ".." + $EndDate
$SearchName = "EmailRemoval_" + (Get-date).ToString("yyyyMMdd_hhmmss")
$SearchActionName = $SearchName + "_Preview"
$SearchActionName = $SearchName + "_Purge"
$Search = New-ComplianceSearch -Name $SearchName -ExchangeLocation all -ContentMatchQuery "(c:c)(senderauthor=$Sender)(subjecttitle=$Subject)(date=$date)"
#$Search = New-ComplianceSearch -Name $SearchName -ExchangeLocation all -ContentMatchQuery "(c:c)(senderauthor=$Sender)(date=$date)"

Start-ComplianceSearch -Identity $SearchName
Get-ComplianceSearch -identity $SearchName
#after 5 mins run this and wait till status is completed
Get-ComplianceSearch -identity $SearchName
#Can be viewed under here https://compliance.microsoft.com/contentsearchv2?viewid=search


#Section2
#For deletion of phising mails

#This section should be run with Global/Compliance/Cloud Admin account, please use below commented command to establish connection
#Connect-IPPSSession -Credential (Get-Credential)

New-ComplianceSearchAction -SearchName "$SearchName" -Purge -PurgeType HardDelete -Confirm:$false
#NOTE correct permission above on account else you will get error purge parameter not found
Get-ComplianceSearchAction -Identity $SearchActionName
#wait for status to change


#Section3 
#For log generation
#check folder paths below

$Identity = (Get-ComplianceSearchAction | Where-Object {$_.SearchName -like "$SearchName"}).Identity
$Results = (Get-ComplianceSearchAction -Identity $Identity).Results
$JobDate = (Get-ComplianceSearchAction -Identity $Identity).JobStartTime

$Results = $Results -split ";" -replace "}","" -replace "{",""
$LogPath = "C:\scripts\o365\phishing_removals\" + $SearchName + "_log.txt"
$Results >> $LogPath
$JobDate >> $LogPath
