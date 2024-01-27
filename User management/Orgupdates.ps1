$Lists = import-csv -path "C:\PS\HRU.csv" |
foreach { 
    $_.User = "{1}, {0}" -f ($_.User -split ' ')
    $_ 
} |
Export-Csv "C:\PS\HRUPass.csv"

$Lists = import-csv -path "C:\PS\HRUpass.csv" |
foreach { 
    $_.Manager = "{1}, {0}" -f ($_.Manager -split ' ')
    $_ 
} |
Export-Csv "C:\PS\Ready\HRU.csv"
$Lists = import-csv -path "C:\PS\FPU.csv" |
foreach { 
    $_.User = "{1}, {0}" -f ($_.User -split ' ')
    $_ 
} |
Export-Csv "C:\PS\FPUPass.csv"

$Lists = import-csv -path "C:\PS\FPUpass.csv" |
foreach { 
    $_.Manager = "{1}, {0}" -f ($_.Manager -split ' ')
    $_ 
} |
Export-Csv "C:\PS\Ready\FPU.csv"
$Lists = import-csv -path "C:\PS\LIU.csv" |
foreach { 
    $_.User = "{1}, {0}" -f ($_.User -split ' ')
    $_ 
} |
Export-Csv "C:\PS\LIUPass.csv"

$Lists = import-csv -path "C:\PS\LIUpass.csv" |
foreach { 
    $_.Manager = "{1}, {0}" -f ($_.Manager -split ' ')
    $_ 
} |
Export-Csv "C:\PS\Ready\LIU.csv"