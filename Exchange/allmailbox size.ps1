$dir = read-host "C:\Tools\" 
$server = read-host "sauex03" 
 
get-mailboxstatistics -server $server | Where-Object {$_.ObjectClass -eq "Mailbox"} |  
Select-Object DisplayName,TotalItemSize,ItemCount,StorageLimitStatus |  
Sort-Object TotalItemSize -Desc | 
export-csv "$dir\mailbox_size.csv"