$US = read-Host 'email address of User to be Checked'
Get-MailboxPermission -identity $US