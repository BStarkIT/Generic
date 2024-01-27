# Random Password Generator
# Change the x,y in the parenthesees right after GeneratePassword
# x = length in characters
# y = minimum number of non-alphanumeric characters
# If you only want alpha-numeric characters, simply use “x.”
# As soon as you use “y”, you’ll get anywhere from y to x number of non-alphanumeric characters.
param
(
    [Parameter(Mandatory)]$length,
    [Parameter(Mandatory)]$Symbols
)
$Password = [system.web.security.membership]::GeneratePassword($length,$Symbols)

Write-Host $Password