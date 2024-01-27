
Import-CSV C:\PS\EdinburghUsers.csv | foreach { Add-MailboxPermission  -Identity TestMB@scotcourts.gov.uk -User  $_.SamAccountName  -AccessRights  FullAccess }