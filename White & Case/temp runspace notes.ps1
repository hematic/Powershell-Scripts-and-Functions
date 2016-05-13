$Parameters = @{

    Param1 = 'Param1'

    Param2 = 'Param2'

}

[runspacefactory]::CreateRunspacePool()

$SessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

$RunspacePool = [runspacefactory]::CreateRunspacePool(

    1, #Min Runspaces

    5 #Max Runspaces

)

$PowerShell = [powershell]::Create()

#Uses the RunspacePool vs. Runspace Property

#Cannot have both Runspace and RunspacePool property used; last one applied wins

$PowerShell.RunspacePool = $RunspacePool

$RunspacePool.Open()




$jobs = New-Object System.Collections.ArrayList

1..50 | ForEach {

    $PowerShell = [powershell]::Create() 

    $PowerShell.RunspacePool = $RunspacePool   

    [void]$PowerShell.AddScript({

        Param (

            $Param1,

            $Param2

        )

        $ThreadID = [appdomain]::GetCurrentThreadId()

        Write-Verbose “ThreadID: Beginning $ThreadID” -Verbose

        $sleep = Get-Random (1..5)       

        [pscustomobject]@{

            Param1 = $param1

            Param2 = $param2

            Thread = $ThreadID

            ProcessID = $PID

            SleepTime = $Sleep

        } 

        Start-Sleep -Seconds $sleep

        Write-Verbose “ThreadID: Ending $ThreadID” -Verbose

    })

    [void]$PowerShell.AddParameters($Parameters)

    $Handle = $PowerShell.BeginInvoke()

    $temp = $Handle | Select PowerShell,Handle

    $temp.PowerShell = $PowerShell

    $temp.handle = $Handle

    [void]$jobs.Add($Temp)   

    Write-Debug (“Available Runspaces in RunspacePool: {0}” -f $RunspacePool.GetAvailableRunspaces())

    Write-Debug (“Remaining Jobs: {0}” -f @($jobs | Where {$_.handle.iscompleted -ne ‘Completed’}).Count)

}



#Verify completed

Write-Debug (“Available Runspaces in RunspacePool: {0}” -f $RunspacePool.GetAvailableRunspaces())

Write-Debug (“Remaining Jobs: {0}” -f @($jobs | Where {

    $_.handle.iscompleted -ne ‘Completed’

}).Count)

$return = $jobs | ForEach {

    $_.powershell.EndInvoke($_.handle)

    $_.PowerShell.Dispose()

}

$jobs.clear()


$return | Group ProcessID | Select Count, Name

$return | Group Thread | Select Count, Name

($return | Group Thread).Count