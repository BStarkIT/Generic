$Path="C:\scripts\o365\reports\"
$date=Get-Date -f "_dd_MM_yyyy"
$CSVOutputFilePath = $Path+"o365_mailbox_report"+$date+"_Export.csv"
#$mailboxes = get-recipient -ResultSize unlimited
$mailboxes = get-mailbox -ResultSize unlimited
$addresses = $mailboxes | select primarysmtpaddress,displayname,userprincipalname,alias,distinguishedname,database,recipienttype,recipienttypedetails,remoterecipienttype,accountdisabled,legacyexchangedn,customattribute1,customattribute15,customattribute4,customattribute5,externalemailaddress,WindowsEmailAddress,ForwardingAddress,ForwardingSmtpAddress,emailaddresses,samaccountname,usagelocation,mailboxregion,whenmailboxcreated,microsoftonlineservicesid,windowsliveid,archivename,archivestatus,archivestate,acceptmessagesonlyfrom,acceptmessagesonlyfromdlmembers,acceptmessagesonlyfromsendersormembers
$addresses | Export-csv -path $CSVOutputFilePath -NoTypeInformation