$looptoggle = $true
echo "Starting Forever Loop!"
while($looptoggle) {
    Write-Host (Get-Content "./GGBOT_LOGS_$getdate.txt" -Delimiter "`r`n")
    Write-Host "Press 'q' to quit!"
    Start-Sleep 2
    cls
    if ([console]::KeyAvailable) {
        $x = [System.Console]::ReadKey()
        switch ($x.key) {
            "q" { $looptoggle = $false; break }
        }
    }
}

# Try putting in the if statement a thing to read host for a key. It'll wait for input, and when receives the right output
# it'll continue the forever loop.
