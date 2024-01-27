param
(
    [Parameter(Mandatory)]$Target
) 
Test-NetConnection $Target -TraceRoute