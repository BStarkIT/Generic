$Server = Read-Host -Prompt 'Input your target IP/host' ; 1..65536 | % {$test= new-object system.Net.Sockets.TcpClient; $wait = $test.beginConnect("$Server",$_,$null,$null); ($wait.asyncwaithandle.waitone(250,$false)); if ($test.Connected){echo "$_ open"}else{echo ""}} | select-string " "