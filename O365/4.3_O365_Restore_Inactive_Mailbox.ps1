#To recover mailbox if in soft deleted status
#$InactiveMailbox = Get-Mailbox -InactiveMailboxOnly -Identity etest2@scotcourts.gov.uk
#$TargetMailbox = Get-Mailbox -Identity etest23@scotcourts.gov.uk


#Connect O365
Get-Mailbox etest2@scotcourts.gov.uk -InactiveMailboxOnly | FL Name,DistinguishedName,LegacyExchangeDN,ExchangeGuid,PrimarySmtpAddress
Get-Mailbox etest2@scotcourts.gov.uk -InactiveMailboxOnly | Select-Object DisplayName,PrimarySmtpAddress, @{Name="EmailAddresses";Expression={($_.EmailAddresses | Where-Object {$_ -clike "smtp*"} | ForEach-Object {$_ -replace "smtp:",""}) -join ","}} | Sort-Object DisplayName | Export-CSV "C:\scripts\O365\inactive\etest2.csv" -NotypeInformation


$InactiveMailbox = Get-Mailbox -InactiveMailboxOnly -Identity etest2@scotcourts.gov.uk
$TargetMailbox = Get-Mailbox -Identity etest23@scotcourts.gov.uk
New-MailboxRestoreRequest -SourceMailbox $InactiveMailbox.DistinguishedName -TargetMailbox $TargetMailbox.ExchangeGuid -AllowLegacyDNMismatch

Get-MailboxRestoreRequest | Get-MailboxRestoreRequestStatistics

#Now restore o365 archive
$InactiveMailbox = Get-Mailbox -InactiveMailboxOnly -Identity etest2@scotcourts.gov.uk
New-MailboxRestoreRequest -SourceMailbox $InactiveMailbox.DistinguishedName -SourceIsArchive -TargetMailbox etest23@scotcourts.gov.uk -TargetIsArchive -AllowLegacyDNMismatch
