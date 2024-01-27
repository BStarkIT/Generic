param
(
    [Parameter(Mandatory)]$Room
)
Get-CalendarProcessing -Identity $Room | fl

