#Script Revision History and Notes
<#
    Current Version : 0.9.7
    FileName        : ServerAudit.ps1
    Original Author : Daniel Lange
    Original Date   : 13 July 2009
    Current Author  : Phillip Marshall
    Takeover Date   : 17 May 2016

    .Description 
        Displays and verifies key and important server configuration
	    settings. Settings displayed in green meet current configuration standards.
	    Settings displayed in red are most likely incorrect and should be verified. 
	    Settings displayed in yellow are discretionary and should be checked
	    against the requirements of the particular server build.

    .Parameter Servers
        Should be a single server, but an array will also 
		work
        
    .Parameter CSV
        Filename of csv with list of servers seperated by comma. Should be
		referenced as named switch.

    .Parameter Txt
        filename of txt file with list of servers seperated by newline. 
		Should be referenced as named switch

    .Notes 
        Script can only accept one parameter.

        Script uses a single main custom PSObject '$ServerInfoObject'.
        Current properties of the object:
            OS - String representing the OS of the Server.
            Servername - String representing the name of the Server.
            Domain - String representing the name of the domain.
            Servicepack - String Representing the servicepack level.
            AdressWidth - String representing either a 64 or 32 bit OS.
            RAM - String representing the amount of RAM on the server.
            SmcService - String representing the status of the SMC Service.
            SmsService - String representing the status of the SMS service.
            FireWallData - Object containing all relevant data about the remote firewall.
            RDPStatus - String representing whethere or not RDP is disabled.
            RDPAuthReq - String representing Whether or not NLA is required for RDP.
            TSMaxDisconTime - String representing the max disconnection time for TS.
            TSMaxIdleTime - String representing the max idle time for TS.
            SNMpStatus - String representing the status of the SNMP service.
            AdminName - The name of the local admin on the machine.
            Diskinfo - An object with all of the disk data for the server.
            NetworkInfo - A list of all adapaters that are IP enabled.
            Processors - Information on the processors for the server.


    .Revisions
        --------------------------------------------------------------------------
        |13 Jul 09 - Added progress bar                                          |
        --------------------------------------------------------------------------
        |14 Jul 09 - Added Test-Host to ping RPC prior to making WMI calls.      |
	    |            OS information retrieval and output made functions.         |
	    |            Added ability to pass input from command line               |
        --------------------------------------------------------------------------
        |15 Jul 09 - Added ability to read in servers from CSV or text file      |
	    |            Fixed autoupdate check for 2003                             |
	    |            Fixed firewall check logic for 2003                         |
        --------------------------------------------------------------------------
        |16 Jul 09 - Added HTML output                                           |
        --------------------------------------------------------------------------
        |17 Jul 09 - Changed HTML colors to be more readable                     |
        --------------------------------------------------------------------------
        |23 Jul 09 - Added check for x86 vs x32                                  |
        --------------------------------------------------------------------------
        |03 Aug 09 - Added freespace calculation / report                        |
	    |            Refined disk check to exclude floppies                      |
        --------------------------------------------------------------------------
        |17 May 16 - Reworked script revision history and notes section.         |
        |                                                                        |
        |            Reworked script parameters and verification.                |
        |                                                                        |
        |            Removed Test-Host Function. Replaced with built in          |
        |            Test-Connection commandlet.                                 |
        |                                                                        |
        |            Renamed "ServerInfo" function to "Get-ServerInformation"    |
        |            to comply with naming conventions.                          |
        |                                                                        |
        |            Replaced 'CheckServices' function with built in get-service |
        |            commandlet.                                                 |
        |                                                                        |
        |            Moved all of the data gathering out of the main script body |
        |            to a function called get-serverinformation.                 |
        |                                                                        |
        |            Added a percent complete param to all write-progress calls. |
        |                                                                        |
        |            Reworked the logic for ALL data collection to be cleaner    |
        |            and have good error checking.                               |
        |-------------------------------------------------------------------------
        |18 May 16 - Made functions for each call that uses invoke command to    |
        |             to connect to the remote machine. Rmoved all old remote wmi|
        |             calls.                                                     |
        |                                                                        |
        |            Removed all references to server 2003.                      |
        |                                                                        |
        |            Removed unneccessary items that no longer are relevant.     |
        |                                                                        |
        |            Changed the script to prompt for credentials at the start.  |
        |                                                                        |
        |            Created SendTo-Console function for immediate results.      |
        |                                                                        |
        |            Reworked almost every calculation related to drives.        |
        |                                                                        |
        |            Added Processor Data.                                       |
        |                                                                        |
        |            Added network adapter data.                                 |
        |                                                                        |
        |            Started reworking the html output. By writing some functions|
        |            to simplify outputting multiple objects to one report.      |
        |-------------------------------------------------------------------------

#>

#########################################################################
#Function Declarations

Function Get-OSData
{
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [String]$Server,
        [Parameter(Position=1, Mandatory=$True)]
        [Object]$Credential
    )


    <#
        Query returns an object like:

        SystemDirectory : C:\WINDOWS\system32
        Organization    : White & Case LLP;
        BuildNumber     : 7601
        RegisteredUser  : White & Case LLP;
        SerialNumber    : 00392-918-5000002-85773
        Version         : 6.1.7601
    #>
    $OS = Invoke-command -ScriptBlock {Get-WmiObject Win32_OperatingSystem} -ComputerName $Server -Credential $Credential
    Return $OS
}

Function Get-ComputerSystemData
{
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [String]$Server,
        [Parameter(Position=1, Mandatory=$True)]
        [Object]$Credential
    )

    <#
        Query returns an object like:

        Domain              : WCNET.whitecase.com
        Manufacturer        : LENOVO
        Model               : 20BT003VUS
        Name                : TAMLT00058
        PrimaryOwnerName    : White & Case LLP;
        TotalPhysicalMemory : 8254623744
    #>
	
    $ComputerSystem = Invoke-command -ScriptBlock {Get-WmiObject Win32_ComputerSystem} -ComputerName $Server -Credential $Credential
    Return $ComputerSystem
}

Function Get-ProcessorData
{
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [String]$Server,
        [Parameter(Position=1, Mandatory=$True)]
        [Object]$Credential
    )

    <#
        Query returns an object like:

        Caption           : Intel64 Family 6 Model 61 Stepping 4
        DeviceID          : CPU0
        Manufacturer      : GenuineIntel
        MaxClockSpeed     : 2301
        Name              : Intel(R) Core(TM) i5-5300U CPU @ 2.30GHz
        SocketDesignation : U3E1
    #>

    $Processor = Invoke-command -ScriptBlock {Get-WmiObject Win32_Processor} -ComputerName $Server -Credential $Credential
    
    Return $Processor
}

Function Get-LocalAdminAccountname
{
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [String]$Server,
        [Parameter(Position=1, Mandatory=$True)]
        [Object]$Credential
    )

    <#
	    The local admin account SID always ends in 500.
        Returns a string like 'WCAdmin'
    #>
    $AdminName = Invoke-command -ScriptBlock {(Get-WmiObject Win32_UserAccount -filter "LocalAccount=True AND SID LIKE '%500'").name} -ComputerName $Server -Credential $Credential
    Return $AdminName
}

Function Get-DiskData
{
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [String]$Server,
        [Parameter(Position=1, Mandatory=$True)]
        [Object]$Credential
    )

    <#
        Query returns an object like:

        DiskSize    : 256052966400
        RawSize     : 256058064896
        FreeSpace   : 189591076864
        Disk        : \\.\PHYSICALDRIVE0
        DriveLetter : C:
        DiskModel   : SAMSUNG MZNLN256HCHP-000 SCSI Disk Device
        VolumeName  : Local Disk
        Size        : 256058060800
        Partition   : Disk #0, Partition #0

    #>	
    $Diskinfo = Invoke-Command -ScriptBlock{
	    Get-WmiObject Win32_DiskDrive | % {
            $disk = $_
            $partitions = "ASSOCIATORS OF " +
                "{Win32_DiskDrive.DeviceID='$($disk.DeviceID)'} " +
                "WHERE AssocClass = Win32_DiskDriveToDiskPartition"
        Get-WmiObject -Query $partitions | % {
            $partition = $_
            $drives = "ASSOCIATORS OF " +
                "{Win32_DiskPartition.DeviceID='$($partition.DeviceID)'} " +
                "WHERE AssocClass = Win32_LogicalDiskToPartition"
        Get-WmiObject -Query $drives | % {
            $DiskSize  = '{0:d} GB' -f [int]($Disk.Size / 1GB)
            $RawSize   = '{0:d} GB' -f [int]($partition.Size / 1GB)
            $Freespace = '{0:d} GB' -f [int]($_.FreeSpace / 1GB)
            $DriveSize = '{0:d} GB' -f [int]($_.Size / 1GB)
            New-Object -Type PSCustomObject -Property @{
            Disk        = $disk.DeviceID
            DiskModel   = $disk.Model
            DiskSize    = $disksize
            DriveLetter = $_.DeviceID
            DriveName   = $_.VolumeName
            DriveSize   = $Drivesize
            Partition   = $partition.Name
            RawSize     = $rawsize
            FreeSpace   = $FreeSpace
        }
        }
        }
        }
        } -ComputerName $Server -Credential $Credential



	Return $Diskinfo
}

Function Get-networkAdapterData
{
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [String]$Server,
        [Parameter(Position=1, Mandatory=$True)]
        [Object]$Credential
    )

    <#
        Query returns an object like:

        DHCPEnabled      : True
        IPAddress        : {172.25.64.65}
        DefaultIPGateway : {172.25.64.254}
        DNSDomain        : wcnet.whitecase.com
        ServiceName      : NETwNs64
        Description      : Intel(R) Dual Band Wireless-AC 7265
        Index            : 9
    #>
    $adapters = Invoke-Command -ScriptBlock {Get-WmiObject win32_networkadapterconfiguration | where-object {$_.ipenabled -eq $true}} -ComputerName $Server -Credential $Credential
    Return $Adapters
}

Function Get-SymantecManagementServiceData
{
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [String]$Server,
        [Parameter(Position=1, Mandatory=$True)]
        [Object]$Credential
    )

    $SmcService = Invoke-Command -ScriptBlock {(Get-Service -Name "SmcService").status} -ComputerName $Server -Credential $Credential -ErrorAction SilentlyContinue
    
    If(!$SmcService)
    {
        $SmcService = 'Missing'
	}
    
    Return $SmcService
}

Function Get-SMSAgentHostServiceData
{
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [String]$Server,
        [Parameter(Position=1, Mandatory=$True)]
        [Object]$Credential
    )
    $SmsService = Invoke-Command -ScriptBlock {(Get-Service -Name "CCMExec").status} -ComputerName $Server -Credential $Credential -ErrorAction SilentlyContinue
    If(!$SmsService)
    {
        $SmsService = 'Missing'
	}
    
    Return $SmsService
}

Function Get-RemoteFirewallStatus
{
    <#
    .SYNOPSIS
       Retrieves firewall status from remote systems via the registry.
    .DESCRIPTION
       Retrieves firewall status from remote systems via the registry.
    .PARAMETER ComputerName
       Specifies the target computer for data query.
    .PARAMETER ThrottleLimit
       Specifies the maximum number of systems to inventory simultaneously 
    .PARAMETER Timeout
       Specifies the maximum time in second command can run in background before terminating this thread.
    .PARAMETER ShowProgress
       Show progress bar information
    .EXAMPLE
       PS > (Get-RemoteFirewallStatus).Rules | where {$_.Active -eq 'TRUE'} | Select Name,Dir
       
       Description
       -----------
       Displays the name and direction of all active firewall rules defined in the registry of the
       local system.
    .EXAMPLE
       PS > (Get-RemoteFirewallStatus).FirewallEnabled

       Description
       -----------
       Displays the status of the local firewall.
    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 2.0

       Version History
       1.0.0 - 10/02/2013
        - Initial release
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(HelpMessage="Computer or computers to gather information from",
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias('DNSHostName','PSComputerName')]
        [string[]]
        $ComputerName=$env:computername,
        
        [Parameter(HelpMessage="Maximum number of concurrent runspaces.")]
        [ValidateRange(1,65535)]
        [int32]
        $ThrottleLimit = 32,
 
        [Parameter(HelpMessage="Timeout before a runspaces stops trying to gather the information.")]
        [ValidateRange(1,65535)]
        [int32]
        $Timeout = 120,
 
        [Parameter(HelpMessage="Display progress of function.")]
        [switch]
        $ShowProgress,
        
        [Parameter(HelpMessage="Set this if you want the function to prompt for alternate credentials.")]
        [switch]
        $PromptForCredential,
        
        [Parameter(HelpMessage="Set this if you want to provide your own alternate credentials.")]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    Begin
    {
        # Gather possible local host names and IPs to prevent credential utilization in some cases
        Write-Verbose -Message 'Firewall Information: Creating local hostname list'
        $IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
        $HostNames = $IPAddresses | ForEach-Object {
            try {
                [net.dns]::GetHostByAddress($_)
            } catch {
                # We do not care about errors here...
            }
        } | Select-Object -ExpandProperty HostName -Unique
        $LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
 
        Write-Verbose -Message 'Firewall Information: Creating initial variables'
        $runspacetimers       = [HashTable]::Synchronized(@{})
        $runspaces            = New-Object -TypeName System.Collections.ArrayList
        $bgRunspaceCounter    = 0
        
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        
        Write-Verbose -Message 'Firewall Information: Creating Initial Session State'
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
        {
            Write-Verbose -Message "Firewall Information: Adding variable $ExternalVariable to initial session state"
            $iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
        }
        
        Write-Verbose -Message 'Firewall Information: Creating runspace pool'
        $rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
        $rp.ApartmentState = 'STA'
        $rp.Open()
 
        # This is the actual code called for each computer
        Write-Verbose -Message 'Firewall Information: Defining background runspaces scriptblock'
        $ScriptBlock = {
            [CmdletBinding()]
            Param
            (
                [Parameter(Position=0)]
                [string]
                $ComputerName,

                [Parameter()]
                [int]
                $bgRunspaceID
            )
            $runspacetimers.$bgRunspaceID = Get-Date
            
            try
            {
                Write-Verbose -Message ('Firewall Information: Runspace {0}: Start' -f $ComputerName)
                $WMIHast = @{
                    ComputerName = $ComputerName
                    ErrorAction = 'Stop'
                }
                if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
                {
                    $WMIHast.Credential = $Credential
                }
                $FwProtocols = @{
                    1="ICMPv4"
                    2="IGMP"
                    6="TCP"
                    17="UDP"
                    41="IPv6"
                    43="IPv6Route"
                    44="IPv6Frag"
                    47="GRE"
                    58="ICMPv6"
                    59="IPv6NoNxt"
                    60="IPv6Opts"
                    112="VRRP"
                    113="PGM"
                    115="L2TP"
                }
                  
                #region Firewall Settings
                Write-Verbose -Message ('Firewall Information: Runspace {0}: Gathering registry information' -f $ComputerName)
                $defaultProperties    = @('ComputerName','FirewallEnabled')
                $HKLM = '2147483650'
                $BasePath = 'System\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy'
                $FW_DomainProfile = "$BasePath\DomainProfile"
                $FW_PublicProfile = "$BasePath\PublicProfile"
                $FW_StandardProfile = "$BasePath\StandardProfile"
                $FW_DomainLogPath = "$($FW_DomainProfile)\Logging"
                $FW_PublicLogPath = "$($FW_PublicProfile)\Logging"
                $FW_StandardLogPath = "$($FW_StandardProfile)\Logging"
                
                $reg = Get-WmiObject @WMIHast -Class StdRegProv -Namespace 'root\default' -List:$true
                $DomainEnabled = [bool]($reg.GetDwordValue($HKLM, $FW_DomainProfile, "EnableFirewall")).uValue
                $PublicEnabled = [bool]($reg.GetDwordValue($HKLM, $FW_PublicProfile, "EnableFirewall")).uValue
                $StandardEnabled = [bool]($reg.GetDwordValue($HKLM, $FW_StandardProfile, "EnableFirewall")).uValue
                $FirewallEnabled = $false
                $FirewallKeys = @() 
                $FirewallKeys += 'System\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\FirewallRules'
                $FirewallKeys += 'Software\Policies\Microsoft\WindowsFirewall\FirewallRules'
                $RuleList = @()
                
                foreach ($Key in $FirewallKeys) 
                { 
                    $FirewallRules = $reg.EnumValues($HKLM,$Key)
                    for ($i = 0; $i -lt $FirewallRules.Types.Count; $i++) 
                    {
                        $Rule = ($reg.GetStringValue($HKLM,$Key,$FirewallRules.sNames[$i])).sValue
                        
                        # Prepare hashtable 
                        $HashProps = @{ 
                            NameOfRule = ($FirewallRules.sNames[$i])
                            RuleVersion = ($Rule -split '\|')[0] 
                            Action = $null 
                            Active = $null 
                            Dir = $null 
                            Proto = $null 
                            LPort = $null 
                            App = $null 
                            Name = $null 
                            Desc = $null 
                            EmbedCtxt = $null 
                            Profile = 'All' 
                            RA4 = $null 
                            RA6 = $null 
                            Svc = $null 
                            RPort = $null 
                            ICMP6 = $null 
                            Edge = $null 
                            LA4 = $null 
                            LA6 = $null 
                            ICMP4 = $null 
                            LPort2_10 = $null 
                            RPort2_10 = $null 
                        } 
             
                        # Determine if this is a local or a group policy rule and display this in the hashtable 
                        if ($Key -match 'System\\CurrentControlSet') 
                        { 
                            $HashProps.RuleType = 'Local' 
                        } 
                        else 
                        { 
                            $HashProps.RuleType = 'GPO'
                        } 
             
                        # Iterate through the value of the registry key and fill PSObject with the relevant data 
                        foreach ($FireWallRule in ($Rule -split '\|')) 
                        { 
                            switch (($FireWallRule -split '=')[0]) { 
                                'Action' {$HashProps.Action = ($FireWallRule -split '=')[1]} 
                                'Active' {$HashProps.Active = ($FireWallRule -split '=')[1]} 
                                'Dir' {$HashProps.Dir = ($FireWallRule -split '=')[1]} 
                                'Protocol' {$HashProps.Proto = $FwProtocols[[int](($FireWallRule -split '=')[1])]} 
                                'LPort' {$HashProps.LPort = ($FireWallRule -split '=')[1]} 
                                'App' {$HashProps.App = ($FireWallRule -split '=')[1]} 
                                'Name' {$HashProps.Name = ($FireWallRule -split '=')[1]} 
                                'Desc' {$HashProps.Desc = ($FireWallRule -split '=')[1]} 
                                'EmbedCtxt' {$HashProps.EmbedCtxt = ($FireWallRule -split '=')[1]} 
                                'Profile' {$HashProps.Profile = ($FireWallRule -split '=')[1]} 
                                'RA4' {[array]$HashProps.RA4 += ($FireWallRule -split '=')[1]} 
                                'RA6' {[array]$HashProps.RA6 += ($FireWallRule -split '=')[1]} 
                                'Svc' {$HashProps.Svc = ($FireWallRule -split '=')[1]} 
                                'RPort' {$HashProps.RPort = ($FireWallRule -split '=')[1]} 
                                'ICMP6' {$HashProps.ICMP6 = ($FireWallRule -split '=')[1]} 
                                'Edge' {$HashProps.Edge = ($FireWallRule -split '=')[1]} 
                                'LA4' {[array]$HashProps.LA4 += ($FireWallRule -split '=')[1]} 
                                'LA6' {[array]$HashProps.LA6 += ($FireWallRule -split '=')[1]} 
                                'ICMP4' {$HashProps.ICMP4 = ($FireWallRule -split '=')[1]} 
                                'LPort2_10' {$HashProps.LPort2_10 = ($FireWallRule -split '=')[1]} 
                                'RPort2_10' {$HashProps.RPort2_10 = ($FireWallRule -split '=')[1]} 
                                Default {} 
                            }
                            if ($HashProps.Name -match '\@')
                            {
                                $HashProps.Name = $HashProps.NameOfRule
                            }
                        } 
                     
                        # Create and output object using the properties defined in the hashtable 
                        $RuleList += New-Object -TypeName 'PSCustomObject' -Property $HashProps 
                    }
                }

                if ($DomainEnabled -or $PublicEnabled -or $StandardEnabled)
                {
                    $FirewallEnabled = $true
                }
                $ResultProperty = @{
                    'PSComputerName' = $ComputerName
                    'PSDateTime' = $PSDateTime
                    'ComputerName' = $ComputerName
                    'FirewallEnabled' = $FirewallEnabled
                    'DomainZoneEnabled' = $DomainEnabled
                    'DomainZoneLogPath' = [string]($reg.GetStringValue($HKLM,$FW_DomainLogPath, "LogFilePath").sValue)
                    'DomainZoneLogSize' = [int]($reg.GetDWORDValue($HKLM,$FW_DomainLogPath,"LogFileSize").uValue)
                    'PublicZoneEnabled' = $PublicEnabled
                    'PublicZoneLogPath' = [string]($reg.GetStringValue($HKLM,$FW_PublicLogPath, "LogFilePath").sValue)
                    'PublicZoneLogSize' = [int]($reg.GetDWORDValue($HKLM,$FW_PublicLogPath,"LogFileSize").uValue)
                    'StandardZoneEnabled' = $StandardEnabled
                    'StandardZoneLogPath' = [string]($reg.GetStringValue($HKLM,$FW_StandardLogPath, "LogFilePath").sValue)
                    'StandardZoneLogSize' = [int]($reg.GetDWORDValue($HKLM,$FW_StandardLogPath,"LogFileSize").uValue)
                    'Rules' = $RuleList
                }

                $ResultObject = New-Object PSObject -Property $ResultProperty
                
                # Setup the default properties for output
                $ResultObject.PSObject.TypeNames.Insert(0,'My.RemoteFirewall.Info')
                $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
                $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                $ResultObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers

                Write-Output -InputObject $ResultObject
            }
            catch
            {
                Write-Warning -Message ('Firewall Information: {0}: {1}' -f $ComputerName, $_.Exception.Message)
            }
            Write-Verbose -Message ('Firewall Information: Runspace {0}: End' -f $ComputerName)
        }
 
        function Get-Result
        {
            [CmdletBinding()]
            Param 
            (
                [switch]$Wait
            )
            do
            {
                $More = $false
                foreach ($runspace in $runspaces)
                {
                    $StartTime = $runspacetimers[$runspace.ID]
                    if ($runspace.Handle.isCompleted)
                    {
                        Write-Verbose -Message ('Firewall Information: Thread done for {0}' -f $runspace.IObject)
                        $runspace.PowerShell.EndInvoke($runspace.Handle)
                        $runspace.PowerShell.Dispose()
                        $runspace.PowerShell = $null
                        $runspace.Handle = $null
                    }
                    elseif ($runspace.Handle -ne $null)
                    {
                        $More = $true
                    }
                    if ($Timeout -and $StartTime)
                    {
                        if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
                        {
                            Write-Warning -Message ('Firewall Information: Timeout {0}' -f $runspace.IObject)
                            $runspace.PowerShell.Dispose()
                            $runspace.PowerShell = $null
                            $runspace.Handle = $null
                        }
                    }
                }
                if ($More -and $PSBoundParameters['Wait'])
                {
                    Start-Sleep -Milliseconds 100
                }
                foreach ($threat in $runspaces.Clone())
                {
                    if ( -not $threat.handle)
                    {
                        Write-Verbose -Message ('Firewall Information: Removing {0} from runspaces' -f $threat.IObject)
                        $runspaces.Remove($threat)
                    }
                }
                if ($ShowProgress)
                {
                    $ProgressSplatting = @{
                        Activity = 'Getting installed programs'
                        Status = 'Firewall Information: {0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
                        PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
                    }
                    Write-Progress @ProgressSplatting
                }
            }
            while ($More -and $PSBoundParameters['Wait'])
        }
    }
    Process
    {
        foreach ($Computer in $ComputerName)
        {
            $bgRunspaceCounter++
            $psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
            $null = $psCMD.AddParameter('bgRunspaceID',$bgRunspaceCounter)
            $null = $psCMD.AddParameter('ComputerName',$Computer)
            $null = $psCMD.AddParameter('Verbose',$VerbosePreference)
            $psCMD.RunspacePool = $rp
 
            Write-Verbose -Message ('Firewall Information: Starting {0}' -f $Computer)
            [void]$runspaces.Add(@{
                Handle = $psCMD.BeginInvoke()
                PowerShell = $psCMD
                IObject = $Computer
                ID = $bgRunspaceCounter
           })
           Get-Result
        }
    }
    End
    {
        Get-Result -Wait
        if ($ShowProgress)
        {
            Write-Progress -Activity 'Firewall Information: Getting program listing' -Status 'Done' -Completed
        }
        Write-Verbose -Message "Firewall Information: Closing runspace pool"
        $rp.Close()
        $rp.Dispose()
    }
}

Function Get-RDPStatus
{
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [String]$Server,
        [Parameter(Position=1, Mandatory=$True)]
        [Object]$Credential
    )

    <#
        Query returns a binary which gets converted to:
        Enabled,Disabled or unknown.
    #>

    $RDPValue = Invoke-command -ScriptBlock {Get-ItemProperty "hklm:\System\ControlSet001\Control\Terminal Server" | select -ExpandProperty fdenytsconnections} -ComputerName $Servername -Credential $Credential
    
    Switch ($RDPValue)
    {
        0 {$RDP = 'Enabled'}
        1 {$RDP = 'Disabled'}
        default {$RDP = 'Unknown'}
    }
        
    Return $RDP
}

Function Get-RDPAuthStatus
{
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [String]$Server,
        [Parameter(Position=1, Mandatory=$True)]
        [Object]$Credential
    )

    <#
        Query returns a binary which gets converted to:
        Enabled,Disabled or unknown.
    #>

    $RDPAuthValue = Invoke-command -ScriptBlock {Get-itemproperty "hklm:\System\ControlSet001\Control\Terminal Server\WinStations\RDP-Tcp" | select -ExpandProperty userauthentication} -ComputerName $Servername -Credential $Credential
    
    Switch ($RDPAuthValue)
    {
        0 {$RDP = 'Disabled'}
        1 {$RDP = 'Enabled'}
        default {$RDP = 'Unknown'}
    }
        
    Return $RDP
}

Function Get-MaxDisconnectTime
{
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [String]$Server,
        [Parameter(Position=1, Mandatory=$True)]
        [Object]$Credential
    )

    <#
        Query returns a value in milliseconds.
    #>

    $MaxTime = Invoke-command -ScriptBlock {Get-itemproperty "hklm:\System\ControlSet001\Control\Terminal Server\WinStations\RDP-Tcp" | select -ExpandProperty MaxDisconnectionTime} -ComputerName $Servername -Credential $Credential
    
    $Timespan = [timespan]::FromMilliseconds($maxtime)
 
    Return $($Timespan.TotalHours)
}

Function Get-MaxIdleTime
{
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [String]$Server,
        [Parameter(Position=1, Mandatory=$True)]
        [Object]$Credential
    )

    <#
        Query returns a value in milliseconds.
    #>

    $MaxTime = Invoke-command -ScriptBlock {Get-itemproperty "hklm:\System\ControlSet001\Control\Terminal Server\WinStations\RDP-Tcp" | select -ExpandProperty MaxIdleTime} -ComputerName $Servername -Credential $Credential
    
    $Timespan = [timespan]::FromMilliseconds($maxtime)
 
    Return $($Timespan.TotalHours)
}

Function Get-ServerInformation 
{
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [String]$Servername,
        [Parameter(Position=1, Mandatory=$True)]
        [Object]$Credential
	)
    $ServerInfoObject = New-Object –TypeName PSObject

    #############
    #Get OS Data#
    #############
	Write-Progress -CurrentOperation "Gathering OS Data" -Activity "Server Audit : $ServerName" -Status "Querying Remote Machine" -PercentComplete 10
	$OS = Get-OSData -Server $Servername -Credential $Credential

    switch -wildcard ($($OS.caption))
    {
        '*2008*' {$ServerInfoObject | Add-Member –MemberType NoteProperty –Name OS –Value 2008}

        '*2012*' {$ServerInfoObject | Add-Member –MemberType NoteProperty –Name OS –Value 2012}

        Default  {$ServerInfoObject | Add-Member –MemberType NoteProperty –Name OS –Value Unknown}
    }
    
    ##########################
    #Get Computer System Data#
    ##########################
	Write-Progress -CurrentOperation "Gathering Computersystem Data" -Activity "Server Audit : $ServerName" -Status "Querying Remote Machine" -PercentComplete 20
	$ComputerSystem = Get-ComputerSystemData -Server $Servername -Credential $Credential
	
    ####################
    #Get Processor Data#
    ####################
    Write-Progress -CurrentOperation "Gathering Processor Data" -Activity "Server Audit : $ServerName" -Status "Querying Remote Machine" -PercentComplete 30
	$Processor = Get-ProcessorData -Server $Servername -Credential $Credential
    $ServerInfoObject | Add-Member –MemberType NoteProperty –Name Processors –Value $Processor
	
    ##############################
    #Get Local Admin Account Name#
    ##############################
    Write-Progress -CurrentOperation "Gathering Local Admin Name" -Activity "Server Audit : $ServerName" -Status "Querying Remote Machine" -PercentComplete 40
	$AdminName = Get-LocalAdminAccountname -Server $Servername -Credential $Credential
    $ServerInfoObject | Add-Member –MemberType NoteProperty –Name LocalAdmin –Value $Adminname

    ###############
    #Get Disk Data#
    ###############
    Write-Progress -CurrentOperation "Gathering Disk Information" -Activity "Server Audit : $ServerName" -Status "Querying Remote Machine" -PercentComplete 50
	$Disks = Get-DiskData -Server $Servername -Credential $Credential
    $ServerInfoObject | Add-Member –MemberType NoteProperty –Name Diskinfo –Value $Disks
	
    ##########################
    #Get network Adapter Data#
    ##########################
    Write-Progress -CurrentOperation "Gathering Network Adapter Information" -Activity "$ServerName" -Status "Querying Remote Machine" -PercentComplete 60
    $Adapters =  Get-networkAdapterData -Server $Servername -Credential $Credential
    $ServerInfoObject | Add-Member –MemberType NoteProperty –Name Networkinfo –Value $Adapters

    ####################################
    #Set Servername and Domain Property#
    ####################################
    $ServerInfoObject | Add-Member –MemberType NoteProperty –Name ServerName –Value $OS.CSName
    $ServerInfoObject | Add-Member –MemberType NoteProperty –Name Domain –Value $ComputerSystem.Domain

    #########################
    #Set ServicePack Version#
    #########################
    $ServerInfoObject | Add-Member –MemberType NoteProperty –Name Servicepack –Value $OS.ServicePackMajorVersion

    ################################
	#Determine and Set AddressWidth#
    ################################
	If ($Processor -is [System.Array]) 
    {
        $ServerInfoObject | Add-Member –MemberType NoteProperty –Name AddressWidth –Value $Processor[0].Addresswidth
	}
 
    Else 
    {
		$ServerInfoObject | Add-Member –MemberType NoteProperty –Name AddressWidth –Value $Processor.Addresswidth
	}

    #######################
    #Determine and Set Ram#
    #######################
    $Ram = ($OS.TotalVisibleMemorySize / 1kb).tostring("F00")
    $ServerInfoObject | Add-Member –MemberType NoteProperty –Name Ram –Value $Ram
	
    ######################################
    #Get Symantec Management Service Data#
    ######################################
	Write-Progress -CurrentOperation "Gathering information on Symantec Services" -Activity "Server Audit : $ServerName" -Status "Querying Remote Machine" -PercentComplete 70
	$SmcService = Get-SymantecManagementServiceData -Server $Servername -Credential $Credential
    $ServerInfoObject | Add-Member –MemberType NoteProperty –Name SmcService –Value $SmcService

    #################################
    #Get SMS Agent Host Service Data#
    #################################
    Write-Progress -CurrentOperation "Gathering Information on Symantec Services" -Activity "Server Audit : $ServerName" -Status "Querying Remote Machine" -PercentComplete 80
    $SmsService = Get-SMSAgentHostServiceData -Server $Servername -Credential $Credential
    $ServerInfoObject | Add-Member –MemberType NoteProperty –Name SmsService –Value $SmcService

    ############################
    #Get Domain Firewall Status#
    ############################
    Write-Progress -CurrentOperation "Gathering Information on Firewall Configuration" -Activity "Server Audit : $ServerName" -Status "Querying Remote Machine" -PercentComplete 85
    $FirewallData = Get-RemoteFirewallStatus -ComputerName $Servername -Credential $Credential
    $ServerInfoObject | Add-Member –MemberType NoteProperty –Name FireWallData –Value $FirewallData
	
    ##############################
    #Get Terminal Services Status#
    ##############################
	Write-Progress -CurrentOperation "Gathering Information on Terminal Services" -Activity "Server Audit : $ServerName" -Status "Querying Remote Machine" -PercentComplete 90
    $RDPStatus = Get-RDPStatus -Server $Servername -Credential $Credential
    $RDPAuthReq = Get-RDPAuthStatus -Server $Servername -Credential $Credential
    $TSMaxDisconTime = Get-MaxDisconnectTime -Server $Servername -Credential $Credential
    $TSMaxIdleTime = Get-MaxIdleTime -Server $Server -Credential $Credential
    
    $ServerInfoObject | Add-Member –MemberType NoteProperty –Name RDPStatus –Value $RDPStatus
    $ServerInfoObject | Add-Member –MemberType NoteProperty –Name RDPAuthReq –Value $RDPAuthReq
    $ServerInfoObject | Add-Member –MemberType NoteProperty –Name TSMaxDisconTime –Value $TSMaxDisconTime
    $ServerInfoObject | Add-Member –MemberType NoteProperty –Name TSMaxIdleTime –Value $TSMaxIdleTime

    #################
    #Get SNMP status#
    #################
    Write-Progress -CurrentOperation "Gathering SNMP Service Status" -Activity "Server Audit : $ServerName" -Status "Querying Remote Machine" -PercentComplete 100
	$SNMPService = Invoke-Command -ScriptBlock {(Get-Service -Name 'SNMP').status} -ComputerName $Servername -Credential $Credential
    $ServerInfoObject | Add-Member –MemberType NoteProperty –Name SNMPStatus –Value $SNMPService

    Return $ServerInfoObject

}

Function Send-ToConsole 
{

	Param
    (
        [Parameter(Position=0,ValueFromPipeline=$true,Mandatory=$true)]
        [Object]$ServerInfo
        
    )
	
	####################################
	#Neatly Format and output to screen#
	####################################
	
    $Information = @"
****************************
*General Server Information*
****************************

Server Name : $($ServerInfo.servername) 
Domain      : $($ServerInfo.domain)  
OS          : $($ServerInfo.OS)
ServicePack : $($ServerInfo.servicepack)
Architecture: $($ServerInfo.Addresswidth)
Ram         : $($ServerInfo.ram) MB

***********************
*Processor Information*
***********************
$($ServerInfo.Processors | select -ExcludeProperty pscomputername | fl  | out-string)
***********
*AV Status*
***********

Symantec Service : $($ServerInfo.Smcservice)
SMS Service      : $($ServerInfo.Smsservice)

*************************
*Windows Firewall Status*
*************************

Public Zone Enabled   : $($Serverinfo.firewalldata.PublicZoneEnabled)
Standard Zone Enabled : $($Serverinfo.firewalldata.StandardZoneEnabled)
Domain Zone Enabled   : $($Serverinfo.firewalldata.DomainZoneEnabled)

****************************
*Terminal Services Settings*
****************************

RDP Status             : $($Serverinfo.RDPStatus)
RDP NLA Required       : $($ServerInfo.RDPAuthReq)
TS Max Disconnect Time : $($ServerInfo.TSMaxDiscontime) Hours
TS Max Idle Time       : $($ServerInfo.TSMaxDiscontime) Hours

****************
*Disk Settings*
****************
$($Serverinfo.diskinfo | fl | out-string)
******************
*Network Settings*
******************
$($Serverinfo.networkinfo | fl | out-string)
****************
*Other Settings*
****************

SNMP Status : $($ServerInfo.SNMPStatus.value)
Local Admin Name : $($ServerInfo.localadmin)

"@
    
    Write-Host $Information

    Return $Information
}

Function Export-HtmlReport
{

<#
.SYNOPSIS
	Creates a HTML report
.DESCRIPTION
	Creates an eye-friendly HTML report with an inline cascading style sheet from given PowerShell objects and saves it to a file.
.PARAMETER InputObject
	Hashtable containing data to be converted into a HTML report.
	
	HashTable Proberties:
	
	[array]  .Object       Any PowerShell object to be converted into a HTML table.

	[string] .Property     Select a set of properties from the object. Default is "*".
	
	[sting]  .As           Use "List" to create a vertical table instead of horizontal alignment.
			
	[string] .Title        Add a table headline.
	
	[string] .Description  Add a table description.

.PARAMETER Title
	Title of the HTML report.
.PARAMETER OutputFileName
	Full path of the output file to create, e.g. "C:\temp\output.html".
.EXAMPLE
	@{Object = Get-ChildItem "Env:"} | Export-HtmlReport -OutputFile "HtmlReport.html" | Invoke-Item

	Creates a directory item HTML report and opens it with the default browser.
.EXAMPLE

	$ReportTitle = "HTML-Report"
	$OutputFileName = "HtmlReport.html"
	
	$InputObject =  @{ 
	                   Title  = "Directory listing for C:\";
	                   Object = Get-Childitem "C:\" | Select -Property FullName,CreationTime,LastWriteTime,Attributes
					},
					@{
	                   Title  = "PowerShell host details";
					   Object = Get-Host;
					   As     = 'List'
					},
					@{
	                   Title       = "Running processes";
					   Description = 'Information about the first 2 running processes.'
					   Object      = Get-Process | Select -First 2
					},
					@{
					   Object = Get-ChildItem "Env:"
					}
					
	Export-HtmlReport -InputObject $InputObject -ReportTitle $ReportTitle -OutputFile $OutputFileName

	Creates a HTML report with separated tables for each given object.
.INPUTS
	Data object, title and alignment parameter
.OUTPUTS
	File object for the created HTML report file.
#>

	[CmdletBinding()]
	Param(
		[Parameter(ValueFromPipeline=$True, Mandatory=$True)]
		[ValidateNotNullOrEmpty()]
		[Array]$InputObject,

		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[String]$ReportTitle = 'Generic HTML-Report',

		[Parameter(Mandatory=$True)]
		[ValidateNotNullOrEmpty()]
		[String]$OutputFileName
	)

	BEGIN
	{
		$HtmlTable		= ''
	}

	PROCESS
	{
		ForEach ($InputElement in $InputObject)
		{
			If ($InputElement.ContainsKey('Title') -eq $False)
			{
				$InputElement.Title = ''
			}

			If ($InputElement.ContainsKey('As') -eq $False)
			{
				$InputElement.As = 'Table'
			}

			If ($InputElement.ContainsKey('Property') -eq $False)
			{
				$InputElement.Property = '*'
			}

			If ($InputElement.ContainsKey('Description') -eq $False)
			{
				$InputElement.Description = ''
			}

			$HtmlTable += $InputElement.Object | ConvertTo-HtmlTable -Title $InputElement.Title -Description $InputElement.Description -Property $InputElement.Property -As $InputElement.As
			$HtmlTable += '<br>'
		}
	}

	END
	{
		$HtmlTable | New-HtmlReport -Title $ReportTitle | Set-Content $OutputFileName
		Get-Childitem $OutputFileName | Write-Output
	}
}

Function ConvertTo-HtmlTable
{

<#
.SYNOPSIS
	Converts a PowerShell object into a HTML table
.DESCRIPTION
	Converts a PowerShell object into a HTML table.	Then use "New-HtmlReport" to create an eye-friendly HTML report with an inline cascading style sheet.
.PARAMETER InputObject
	Any PowerShell object to be converted into a HTML table.
.PARAMETER Property
	Select object properties to be used for table creation. Default is "*"
.PARAMETER As
	Use "List" to create a vertical table. All other values will create a horizontal table.
.PARAMETER Title
	Adds an additional table with a title. Very useful for multi-table-reports!
.PARAMETER Description
	Adds an additional table with a description. Very useful for multi-table-reports!
EXAMPLE
	Get-Process | ConvertTo-HtmlTable
	
	Returns a HTML table as a string.
.EXAMPLE
	Get-Process | ConvertTo-HtmlTable | New-HtmlReport | Set-Content "HtmlReport.html"

	Returns a HTML report and saves it as a file.
.EXAMPLE
	$body =	ConvertTo-HtmlTable -InputObject $(Get-Process) -Property Name,ID,Path -As "List" -Title "Process list" -Description "Shows running processes as a list"
	New-HtmlReport -Body $body | Set-Content "HtmlReport.html"

	Returns a HTML report and saves it as a file.
.INPUTS
	Any PowerShell object
.OUTPUTS
	HTML table as String
#>

	[CmdletBinding()]
	Param(
		[Parameter(ValueFromPipeline=$True, Mandatory=$True)]
		[ValidateNotNullOrEmpty()]
		[PSObject]$InputObject,

		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[Object[]]$Property = '*',

		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[String]$As = 'TABLE',
		
		[Parameter()]
		[String]$Title,

		[Parameter()]
		[String]$Description
	)
	
	BEGIN
	{
		$InputObjectList = @()
		
		If ($As -ne 'LIST')
		{
			$As = 'TABLE'
		}
	}

	PROCESS
	{
		$InputObjectList += $InputObject
	}

	END
	{
		$ofs = "`r`n"	# Set separator for string-convertion to carrige return
		[String]$HtmlTable = $InputObjectList | ConvertTo-HTML -Property $Property -As $As -Fragment
		Remove-Variable ofs -force

		# Add description table
		If ($Description)
		{
			$Html		= '<table id="TableDescription"><tr><td>' + "$Description</td></tr></table>`n"
			$Html 		+= '<table id="TableSpacer"></table>' + "`n"
			$HtmlTable	= $Html + $HtmlTable
		}
		Else
		{
			$Html 		= '<table id="TableSpacer"></table>' + "`n"
			$HtmlTable	= $Html + $HtmlTable
		}
		
		# Add title table
		If ($Title)
		{
			$Html		= '<table id="TableHeader"><tr><td>' + "$Title</td></tr></table>`n"
			$HtmlTable	= $Html + $HtmlTable
		}

		# Add missing data separator tag <hr> to second column (on list-tables only)
		$HtmlTable = $HtmlTable -Replace '<hr>', '<hr><td><hr></td>'
				
		Write-Output $HtmlTable	
	}
}

Function New-HtmlReport
{

<#
.SYNOPSIS
	Creates a HTML report
.DESCRIPTION
	Creates an eye-friendly HTML report with an inline cascading style sheet for a given HTML body.
	Usage of "ConvertTo-HtmlTable" is recommended to create the HTML body.
.PARAMETER Body
	Any HTML body, e.g. a table. Usage of "ConvertTo-HtmlTable" is recommended
	to create an according table from any PowerShell object.
.PARAMETER Title
	Title of the HTML report.
.PARAMETER Head
	Any HTML code to be inserted into the head-tag, e.g. scripts or meta-information.
.PARAMETER CssUri
	Path to a CSS-File to be included as an inline css.
	If CssUri is invalid or not provided, a default css is used instead.
.EXAMPLE
	Get-Process | ConvertTo-HtmlTable | New-HtmlReport
	
	Returns a HTML report as a string.
.EXAMPLE
	Get-Process | ConvertTo-HtmlTable | New-HtmlReport -Title "HTML Report with CSS" | Set-Content "HtmlReport.html"

	Returns a HTML report and saves it as a file.
.EXAMPLE
	$body =	Get-Process | ConvertTo-HtmlTable
	New-HtmlReport -Body $body -Title "HTML Report with CSS" -Head '<meta name="author" content="Thomas Franke">' -CssUri "stylesheet.css" | Set-Content "HtmlReport.html"

	Returns a HTML report with an alternative CSS and saves it as a file.
.INPUTS
	HTML body as String
.OUTPUTS
	HTML page as String
#>

	[CmdletBinding()]
	Param(
		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[String]$CssUri,

		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[String]$Title = 'HTML Report',

		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[String]$Head = '',

		[Parameter(ValueFromPipeline=$True, Mandatory=$True)]
		[ValidateNotNullOrEmpty()]
		[Array]$Body
	)

	# Add title to head because -Title parameter is ignored if -Head parameter is used
	If ($Title)
	{
		$Head = "<title>$Title</title>`n$Head`n"
	}

	# Add inline stylesheet
	If (($CssUri) -And (Test-Path $CssUri))
	{
		$Head += "<style>`n" + $(Get-Content $CssUri | ForEach {"`t$_`n"}) + "</style>`n"
	}
	Else
	{
		$Head += @'
<style>
	table
		{
			Margin: 0px 0px 0px 4px;
			Border: 1px solid rgb(190, 190, 190);
			Font-Family: Tahoma;
			Font-Size: 8pt;
			Background-Color: rgb(252, 252, 252);
		}

	tr:hover td
		{
			Background-Color: rgb(150, 150, 220);
			Color: rgb(255, 255, 255);
		}

	tr:nth-child(even)
		{
			Background-Color: rgb(242, 242, 242);
		}
	  
	th
		{
			Text-Align: Left;
			Color: rgb(150, 150, 220);
			Padding: 1px 4px 1px 4px;
		}


	td
		{
			Vertical-Align: Top;
			Padding: 1px 4px 1px 4px;
		}

	#TableHeader
		{
			Margin-Bottom: -1px;
			Background-Color: rgb(255, 255, 225);
			Width: 30%;
		}
		
	#TableDescription
		{
			Background-Color: rgb(252, 252, 252);
			Width: 30%;
		}

	#TableSpacer
		{
			Height: 6px;
			Border: 0px;
		}
</style>
'@
	}

	$HtmlReport = ConvertTo-Html -Head $Head
	
	# Delete empty table that is created because no input object was given
	$HtmlReport = $($HtmlReport -Replace '<table>', '') -Replace '</table>', $Body

	Write-Output $HtmlReport
}

#########################################################################
#Variable Declarations
[Array]$ServerDataArray = @()
#$Credential = Get-Credential
$secpasswd = ConvertTo-SecureString $ENV:ADMPassword -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential ($ENV:ADMUsername, $secpasswd)
$Server = $ENV:Targetmachine

#########################################################################
#Main Script

$ErrorActionPreference = "Stop"

If(!$ENV:Targetmachine)
{
    Write-Output "No server data was passed to the script. Please provide a servername or a path to a txt or csv file."
    exit;
}

#Process each Server


If(!(Test-connection -Computername $Server -Quiet))
{
    Write-Output "$Server not found!"
    exit;
}

#Pull all of the data from the server.
$ServerData = Get-ServerInformation -ServerName $Server -Credential $Credential
    
#Just to easily reference it later.
$ServerDataArray += $ServerData

#Display some simple results to console.
Send-ToConsole -Serverinfo $ServerData

#Break apart the main objects in smaller objects for ease of creating separate tables.

$GeneralInfo_OBJ = New-Object –TypeName PSObject
    $GeneralInfo_OBJ | Add-Member –MemberType NoteProperty –Name "Server Name"  –Value $Serverdata.ServerName
    $GeneralInfo_OBJ | Add-Member –MemberType NoteProperty –Name "Domain Name"  –Value $Serverdata.Domain
    $GeneralInfo_OBJ | Add-Member –MemberType NoteProperty –Name "OS"           –Value $Serverdata.OS
    $GeneralInfo_OBJ | Add-Member –MemberType NoteProperty –Name "Service Pack" –Value $Serverdata.Servicepack
    $GeneralInfo_OBJ | Add-Member –MemberType NoteProperty –Name "32 or 64 bit" –Value $Serverdata.AddressWidth
    $GeneralInfo_OBJ | Add-Member –MemberType NoteProperty –Name "Memory"       –Value "$($Serverdata.Ram) MB"
    $GeneralInfo_OBJ | Add-Member –MemberType NoteProperty –Name "Local Admin"  –Value $Serverdata.LocalAdmin

$ServiceInfo_OBJ = New-Object –TypeName PSObject
    $ServiceInfo_OBJ | Add-Member –MemberType NoteProperty –Name "SMC Service"  –Value $Serverdata.SmcService
    $ServiceInfo_OBJ | Add-Member –MemberType NoteProperty –Name "SMS Service"  –Value $Serverdata.SmsService
    $ServiceInfo_OBJ | Add-Member –MemberType NoteProperty –Name "SNMP Service" –Value $Serverdata.SNMPStatus.Value
    
$FirewallInfo_OBJ  = New-Object –TypeName PSObject
    $FirewallInfo_OBJ | Add-Member –MemberType NoteProperty –Name "Public Zone Enabled"   –Value $Serverdata.FireWallData.PublicZoneEnabled   
    $FirewallInfo_OBJ | Add-Member –MemberType NoteProperty –Name "Standard Zone Enabled" –Value $Serverdata.FireWallData.StandardZoneEnabled 
    $FirewallInfo_OBJ | Add-Member –MemberType NoteProperty –Name "Domain Zone Enabled"   –Value $Serverdata.FireWallData.DomainZoneEnabled 
    
$RemotingInfo_OBJ = New-Object –TypeName PSObject    
    $RemotingInfo_OBJ | Add-Member –MemberType NoteProperty –Name "RDP Status"             –Value $Serverdata.RDPStatus
    $RemotingInfo_OBJ | Add-Member –MemberType NoteProperty –Name "RDP NLA Required"       –Value $Serverdata.RDPAuthReq
    $RemotingInfo_OBJ | Add-Member –MemberType NoteProperty –Name "TS Max Disconnect Time" –Value $Serverdata.TSMaxDisconTime
    $RemotingInfo_OBJ | Add-Member –MemberType NoteProperty –Name "TS Max Idle Time"       –Value $Serverdata.TSMaxIdleTime
    
    
$ProcessorsInfo_OBJ = $($($Serverdata.Processors) | select caption,deviceid,manufacturer,maxclockspeed,name,socketdesignation)
$DiskInfo_OBJ       = $($ServerData.Diskinfo | select DiskSize, Disk, DiskModel, DriveName, DriveLetter, DriveSize, FreeSpace)
$NetworkInfo_OBJ    = $($Serverdata.networkinfo | select DHCPEnabled, IPAddress, DefaultIPGateway,ServiceName,Description)

#OutputThe results to HTML.
    
$OutputFileName = "$Env:windir\temp\$Server" + '.html'
$ReportTitle = "Server Audit Report for $Server"
$InputObject =  @{
                    Title  = "General Server Information";
                    Description = 'The following table contains General Server Information'
                    Object = $GeneralInfo_OBJ
                },
                @{
                    Title  = "Procesor Information";
                    Description = 'The following table contains Processor Information.'
                    Object = $ProcessorsInfo_OBJ
                },
                @{
                    Title  = "Storage Information";
                    Description = 'The following table contains Disk, Drive, and Partition Information.'
                    Object = $DiskInfo_OBJ
                },
                @{
                    Title  = "Network Adapter Information";
                    Description = 'The following table contains information on all IP-Enabled Network Adapters.'
                    Object = $NetworkInfo_OBJ
                },
                @{
                    Title  = "Firewall Information";
                    Description = 'The following table contains the status for each Windows Firewall Zone.'
                    Object = $FirewallInfo_OBJ
                },
                @{
                    Title  = "RDP and Terminal Services Information";
                    Description = 'The following table contains information about remote access.'
                    Object = $RemotingInfo_OBJ
                },
                @{
                    Title  = "Services Information";
                    Description = 'The following table contains information about key services.'
                    Object = $ServiceInfo_OBJ
                }

Export-HtmlReport -InputObject $InputObject -ReportTitle $ReportTitle -OutputFile $OutputFileName

