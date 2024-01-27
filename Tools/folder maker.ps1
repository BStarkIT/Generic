Get-ChildItem -File -Path $directory -Filter "*.jpg" | 
ForEach-Object {
New-Item -ItemType Directory "$directory$($_.Name.Split("_")[0])" -Force;   
Move-Item -Path $_.FullName -Destination  "$directory$($_.Name.Split("_")[0])\$($_.Name)"      
}