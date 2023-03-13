$count = 0
$jobslist = @{jobs=[System.Collections.ArrayList]::new()}
$serverlist = ('server1','server2','server3','server4')

Write-Host "Creating Jobs..."
foreach ($server in $serverlist) {
    $scriptblock = {
        $ret = "Server $($args[0]) started."
        if ($args[0] -eq 'server2'){
            start-sleep 100 # simulating server taking too long to boot
        }
        else {
            Start-Sleep 10 # simulating standard server boot
        }
        return $ret
    }
    $count++
    $jobslist.jobs.Add($(Start-Job -Name "Start-Server-$($server)" -ScriptBlock $scriptblock -ArgumentList $server)) | out-null
}

Write-Host "Created, waiting for jobs to finish..."
#Wait-Job -Name "Start-Server-*" -Timeout 20
$jobslist.jobs | Wait-Job -Timeout 20
Write-Host "All jobs finished."

$result = $jobslist.jobs | Get-Job | Receive-Job

foreach ($job in $jobslist.jobs) {
    if ($job.State -eq "Completed") {
        Write-Host "Task $($job.Name) completed"
    }
    else {
        Write-Warning "Task $($job.Name) not completed"
    }
}