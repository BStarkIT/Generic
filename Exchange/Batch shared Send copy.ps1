$Shared = 'peterheadcivil@scotcourts.gov.uk','peterheadcommissary@scotcourts.gov.uk','peterheadcriminal@scotcourts.gov.uk','peterheadfines@scotcourts.gov.uk','Peterhead@scotcourts.gov.uk','peterheadsimpprocedure@scotcourts.gov.uk'

ForEach ($Share in $Shared) {
    set-mailbox $Share -MessageCopyForSendOnBehalfEnabled $True
    set-mailbox $Share -MessageCopyForSentAsEnabled $True
}