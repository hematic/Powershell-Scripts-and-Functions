Function Get-RebootHistory {  
    <# 
        .SYNOPSIS 
            Retrieves historical information about shutdown/restart events from one or more remote computers. 
     
        .DESCRIPTION 
            The Get-RebootHistory function uses Windows Management Instrumentation (WMI) to retrieve information about all shutdown events from a remote computer.   
             
            Using this function, you can analyze shutdown events across a large number of computers to determine how frequently shutdown/restarts are occurring, whether unexpected shutdowns are occurring and quickly identify the source of the last clean shutdown/restart. 
             
            Data returned includes date/time information for all available boot history events (e.g. restarts, shutdowns, unexpected shutdowns, etc.), date/time information for unexpected reboots and detailed information about the last clean shutdown including date/time, type, initiating user, initiating process and reason.      
             
            Because Get-RebootHistory uses WMI to obtain shutdown event history from the system event log, it is fully supported against both legacy and current versions of Windows including legacy versions that do not support filtering of event logs through standard methods.  
         
        .PARAMETER ComputerName 
            Accepts a single computer name or an array of computer names separated by commas (e.g. "prod-web01","prod-web02").  
 
            This is an optional parameter, the default value is the local computer ($Env:ComputerName). 
         
        .PARAMETER Credential 
            Accepts a standard credential object. 
             
            This is an optional parameter and is only necessary when the running user does not have access to the remote computer(s). 
 
        .EXAMPLE 
            .\Get-RebootHistory -ComputerName prod-web01,prod-web02 -Credential (Get-Credential) 
         
            Get boot history for multiple remote computers with alternate credentials.  
         
        .EXAMPLE 
            .\Get-RebootHistory -ComputerName prod-web01,prod-web02 -Credential (Get-Credential) | ? { $_.PercentDirty -ge 30 } 
         
            Get a list of computers experiencing a high percentage of unexpected shutdown events. 
         
        .EXAMPLE  
            .\Get-RebootHistory -ComputerName prod-web01,prod-web02 -Credential (Get-Credential) | ? { $_.RecentShutdowns -ge 3 }  
         
            Return information about servers that have been experiencing frequent shutdown/reboot events over the last 30 days. 
         
        .OUTPUTS 
            System.Management.Automation.PSCustomObject  
             
            Return object includes the following properties:  
             
                Computer                 
                BootHistory                : Array of System.DateTime objects for all recorded instances of the system booting (clean or otherwise). 
                RecentShutdowns            : The number of shutdown/restart events in the last 30 days. 
                UnexpectedShutdowns        : Array of System.DateTime objects for all recorded unexpected shutdown events. 
                RecentUnexpected        : The number of unexpected shutdown events in the last 30 days. 
                PercentDirty            : The percentage of shutdown events that were unexpected (UnexpectedShutdowns/BootHistory). 
                LastShutdown            : System.DateTime object of the last clean shutdown event. 
                LastShutdownType        : Type of the last clean shutdown event (Restart | Shutdown). 
                LastShutdownUser        : The user who initiated the last clean shutdown event. 
                LastShutdownProcess        : The process that initiated the last clean shutdown event. 
                LastShutdownReason        : If available, the reason code and comments for the last clean shutdown event. 
                 
        .NOTES 
            N/A
    #> 
     
    Param ( 
        [Parameter(Mandatory = $False, Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)] 
        [Alias("CN","Computer")] 
        [Array]$ComputerName = $Env:ComputerName, 
 
        [Parameter(Mandatory = $False, Position = 1, ValueFromPipeline = $False)] 
        [Alias("Cred")] 
        [ValidateNotNull()] 
        [System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty 
    ) 
     
    Begin { 
        $i = 0 
        $RecentShutdowns = 0 
        $RecentUnexpected = 0 
         
        $BootHistory = @() 
        $ShutdownDetail = @() 
        $UnexpectedShutdowns = @()  
         
        # Store original credential, if we attempt to make a local connection we need to  
        # temporarily empty out the credential object. 
        $Original_Credential = $Credential 
         
        # Select properties defined to ensure proper display order. 
        $BootInformation = @( 
            "Computer" 
            "BootHistory" 
            "RecentShutdowns" 
            "UnexpectedShutdowns" 
            "RecentUnexpected" 
            "PercentDirty" 
            "LastShutdown" 
            "LastShutdownType" 
            "LastShutdownUser" 
            "LastShutdownProcess" 
            "LastShutdownReason" 
        ) 
         
        # Arguments to be passed to our WMI call.  
        $Params = @{ 
            ErrorAction        = 'Stop' 
            ComputerName    = $Computer 
            Credential        = $Credential 
            Class            = 'Win32_NTLogEvent' 
            Filter            = "LogFile = 'System' and EventCode = 6009 or EventCode = 6008 or EventCode = 1074" 
        } 
    } 
 
    Process { 
        ForEach ($Computer In $ComputerName) { 
            $Params.ComputerName = $Computer 
             
            # You can't use credentials when connecting to the local machine so temporarily empty out the credential object. 
            If ($Computer -eq $Env:ComputerName) {  
                $Params.Credential = [System.Management.Automation.PSCredential]::Empty 
            } 
             
            If ($ComputerName.Count -gt 1) {  
                Write-Progress -Id 1 -Activity "Retrieving boot history." -Status ("Percent Complete: {0:N0}" -f $($i / $($ComputerName.Count)*100)) -PercentComplete (($i / $ComputerName.Count)*100); $i++ 
            } Else {  
                Write-Progress -Id 1 -Activity "Retrieving boot history." 
            } 
 
            Try {  
                $d = 0 
                $Events = Get-WmiObject @Params 
                 
                ForEach ($Event In $Events) {  
                    Write-Progress -Id 2 -ParentId 1 -Activity "Processing reboot history." -PercentComplete (($d / $Events.Count)*100); $d++ 
                     
                    # Record the relevant details for the shutdown event. 
                    Switch ($Event.EventCode) {  
                        6009 { $BootHistory += (Get-Date(([WMI]'').ConvertToDateTime($Event.TimeGenerated)) -Format g) } 
                        6008 { $UnexpectedShutdowns += ('{0} {1}' -f ($Event.InsertionStrings[1], $Event.InsertionStrings[0])) } 
                        1074 { $ShutdownDetail += $Event } 
                    } 
                } 
                 
                # We explicitly ignore exceptions originating from this process since some versions of Windows may store dates in invalid formats (e.g. ?11/?16/?2014) in the event log after an unexpected shutdown causing this calculation to fail. 
                Try {  
                    $RecentUnexpected = ($UnexpectedShutdowns | ? { ((Get-Date)-(Get-Date $_)).TotalDays -le 30 }).Count 
                } Catch {  
                    $RecentUnexpected = "Unable to calculate." 
                }  
                 
                # Grab details about the last clean shutdown and generate our return object. 
                $ShutdownDetail | Select -First 1 | ForEach-Object {  
                    New-Object -TypeName PSObject -Property @{ 
                        Computer = $Computer 
                        BootHistory = $BootHistory  
                        RecentUnexpected = $RecentUnexpected 
                        LastShutdownUser = $_.InsertionStrings[6] 
                        UnexpectedShutdowns = $UnexpectedShutdowns 
                        LastShutdownProcess = $_.InsertionStrings[0] 
                        PercentDirty = '{0:P0}' -f (($UnexpectedShutdowns.Count/$BootHistory.Count)) 
                        LastShutdownType = (Get-Culture).TextInfo.ToTitleCase($_.InsertionStrings[4]) 
                        LastShutdown = (Get-Date(([WMI]'').ConvertToDateTime($_.TimeGenerated)) -Format g) 
                        RecentShutdowns = ($BootHistory | ? { ((Get-Date)-(Get-Date $_)).TotalDays -le 30 }).Count 
                        LastShutdownReason = 'Reason Code: {0}, Reason: {1}' -f ($_.InsertionStrings[3], $_.InsertionStrings[2]) 
                    } | Select $BootInformation     
                }             
            } Catch [System.Exception] {  
                # We explicitly ignore exceptions originating from Get-Date since some versions of Windows may store dates in invalid formats in the event log after an unexpected shutdown. 
                If ($_.CategoryInfo.Activity -ne 'Get-Date') {  
                    Write-Warning ("Unable to retrieve boot history for {0}. `nError Details: {1}" -f ($Computer, $_)) 
                } 
            } 
             
            # Reset credential object since we may have temporarily overwrote it to deal with local connections. 
            $Params.Credential = $Original_Credential 
        } 
    } 
} 