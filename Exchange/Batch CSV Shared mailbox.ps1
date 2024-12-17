
Import-CSV C:\PS\EdinburghUsers.csv | ForEach-Object { Add-MailboxPermission  -Identity TestMB@scotcourts.gov.uk -User  $_.SamAccountName  -AccessRights  FullAccess }