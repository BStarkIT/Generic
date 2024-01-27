$TempSharedMailbox = read-Host 'name of service account'
$TicketTrace = read-Host 'Ticket number'
$Mail = Read-Host 'Mail enabled? (y/n)'
$Description = Read-Host 'Description'
$SharedMailbox = $TempSharedMailbox -replace '[^\x30-\x39\x2D\x41-\x5A\x61-\x7A]+', ''
$Ticket = $TicketTrace -replace '[^\x30-\x39]+', ''
write-host $TicketTrace
if ($ticket -eq '') {
    write-host "this is empty"
}
if ($null -eq $Ticket) {
    write-host "this is null"
}
Write-host $Ticket