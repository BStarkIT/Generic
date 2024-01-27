$charlist1 = [char]97..[char]122
$charlist2 = [char]65..[char]90
$charlist3 = [char]48..[char]57
$charlist4 = [char]33..[char]38 + [char]40..[char]43 + [char]45..[char]46 + [char]64
        $pwdList = @()
        $pwLength = 2
        For ($i = 0; $i -lt $pwlength; $i++) {
            $pwdList += $charlist1 | Get-Random
            $pwdList += $charlist2 | Get-Random
            $pwdList += $charlist3 | Get-Random
            $pwdList += $charlist4 | Get-Random
            $pwdList += $charlist1 | Get-Random
            $pwdList += $charlist2 | Get-Random
            $pwdList += $charlist3 | Get-Random
        }
        $pass = -join ($pwdList | get-random -count $pwdList.count)
        Write-Output $pass | Clip
        Write-Output $pass