Param 
(
     [parameter(Mandatory = $true,Position = 0)] $PFXPassword,
     [parameter(Mandatory = $true,Position = 1)] $WindowsPassword,
     [parameter(Mandatory = $false,Position = 2)] $IP2
)

#region Do Function Declaration

Function Write-Log
{
	<#
	.SYNOPSIS
		A function to write ouput messages to a logfile.
	
	.DESCRIPTION
		This function is designed to send messages to a logfile of your choosing.
		Use it to replace something like write-host for a more long term log.
	
	.PARAMETER Message
		The message being written to the log file.
	
	.EXAMPLE
		PS C:\> Write-Log -Message 'This is the message being written out to the log.' 
	
	.NOTES
		N/A
#>
	
	Param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[String]$Message
	)

    
	add-content -path $LogFilePath -value $Message
    write-log $Message
}

#endregion

#region Do Variable Set
[String]$DNS = "209.244.0.3" #level3
[String]$PFXThumbprint = "E31975030B72A245878BBC336BA87CF1DEF0C3D8"
[String]$PFXUrl = "https://cloud.screenconnect.com/downloads/ScreenConnectComWildcard.pfx"
[String]$cloudServiceUrl = "https://cloud.screenconnect.com/downloads/CloudService.zip"
[HashTable]$cloudServices = @{"ScreenConnect Router" = "Router"; "ScreenConnect Instance Manager" = "InstanceManager" }
[String]$windowsUserName = "screenconnect"
[Array]$rdpIPs = @("96.10.31.0/24", "63.145.136.0/24", "70.46.245.0/24")
[String]$LogFilePath = "$env:windir\temp\SCServerSetup.txt"
$ProgressPreference = 'SilentlyContinue'
#endregion

#region Do Get Ip1 Information

[Object]$IP1Adapter = (Get-NetAdapter | Where-Object {$_.status -eq 'Up'} | Select -First 1)
[Object]$IP1Config = Get-NetIPConfiguration | Where-Object {$_.InterfaceIndex -eq $IP1Adapter.ifindex}
[String]$SubnetMask = (Get-WmiObject Win32_NetworkAdapterConfiguration | ? { $_.IPSubnet.Length } | % { $_.IPSubnet[0] })
[String]$IP1ExternalIP = 

#endregion

#region Do Get IP2 Information
if (!$ip2) 
{
    write-log "IP2 Was NOT passed. Going to check for two LAN IPs"
    [Array]$Ip1Split = ($IP1Config.IPv4Address.ipaddress).split(".")

    Foreach($Interface in Get-NetIPConfiguration)
    {
        $InterfaceIpSplit = ($Interface.ipv4address.ipaddress).split(".")
        If($Ip1Split[0] -eq $InterfaceIpSplit[0] -and $($Interface.ipv4address.ipaddress) -ne $($IP1Config.ipv4address.ipaddress))
        {
            Write-Log "Second Lan IP Address Identified : Alias - $($Interface.ipv4address.alias) IP - $($Interface.ipv4address.ipaddress)"
            $IP2Config = $Interface
        }
    }

    If(!$IP2Config)
    {
        Write-Log "Unable to find a second lan IP with the same first Octet."
        exit;
    }


} 

write-log "Adding IP2=$ip2 and configuring static with /24 subnet and ADAPTER=$adapterName, IP1=$ip1, SUBNETMASK=$subnetMask, GATEWAY=$gateway, DNS=$dns"
netsh interface ip set address $adapterName static $ip1 $subnetMask $gateway
netsh interface ip add address $adapterName $ip2 $subnetMask

#endregion

#region Do Configure Time
Write-Log "Beginning Time Server Configuration Region"
w32tm /config /manualpeerlist:pool.ntp.org /syncfromflags:MANUAL 2>&1 | Out-Null
Stop-Service w32time

If((Get-Service w32time).status -ne 'Stopped')
{
    Write-Log "Unable to stop the w32time service."
    exit;
}

Start-Service w32time

If((Get-Service w32time).status -ne 'Running')
{
    Write-Log "Unable to start the w32time service."
    exit;
}
#endregion

#region Do Configure PageFile
Write-Log "Setting page file to auto-managed"
$System = gwmi Win32_ComputerSystem -EnableAllPrivileges
$System.AutomaticManagedPagefile = $True
$System.Put() | Out-Null
#endregion

#region Do Stop Services
Write-Log "Beginning the Stop Services Region"
Foreach($Service in $($cloudServices.GetEnumerator() | Select -ExpandProperty Name))
{ 
    Write-Log "Stopping the $Service service."
    Stop-Service $Service -ErrorAction SilentlyContinue 

    If((Get-Service $Service).status -ne 'Stopped')
    {
        Write-Log "Unable to stop the $Service service."
        exit;
    }
    Else
    {
        Write-Log "Service Stopped Successfully."
    }
}
#endregion

#region Do Configure Windows Firewall
Write-Log "Beginning the Configure Windows Firewall Region"

if (Get-NetFirewallRule | ? { $_.DisplayName -eq "Allow 80 Inbound" }) 
{
    Write-Log "Firewall already configured."
} 

else 
{
    Write-Log "Configuring firewall main ports..."
    New-NetFirewallRule -DisplayName "Allow 80 Inbound" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 80 | Out-Null
    Write-Log "Port 80 Rule Created."
    New-NetFirewallRule -DisplayName "Allow 443 Inbound" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 443 | Out-Null
    Write-Log "Port 443 Rule Created."
    New-NetFirewallRule -DisplayName "Allow 5986 Inbound" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 5986 | Out-Null
    Write-Log "Port 5986 Rule Created."
}

if (!(Get-NetFirewallRule | ? { $_.DisplayName -eq "Allow 3389 Inbound" })) 
{
    New-NetFirewallRule -DisplayName "Allow 3389 Inbound" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 3389 -RemoteAddress $rdpIPs | Out-Null
    Write-Log "Port 3389 Rule Created."
}

Write-Log "Disabling blanket RDP and management firewall rules..."
Get-NetFirewallRule | ? { $_.DisplayName.StartsWith("Remote Desktop") } | Disable-NetFirewallRule
Get-NetFirewallRule | ? { $_.DisplayName.StartsWith("Windows Remote Management") } | Disable-NetFirewallRule
Get-NetFirewallRule | ? { $_.DisplayName.StartsWith("Hyper-V") } | Disable-NetFirewallRule
Get-NetFirewallRule | ? { $_.DisplayName.StartsWith("RDP") } | Disable-NetFirewallRule
Get-NetFirewallRule | ? { $_.DisplayName.StartsWith("SSH") } | Disable-NetFirewallRule
#endregion

#region Do Download Debuggers
Write-Log "Beginning Download debuggers Region"

Try
{
    Invoke-WebRequest "https://cloud.screenconnect.com/downloads/Mdbg4_x86.exe" -OutFile \ScreenConnect\Mdbg4_x86.exe
}

Catch
{
    $DownloadErr = $_.Exception.Message;
    Write-Log $DownloadErr
    exit;
}

Try
{
    Invoke-WebRequest "https://cloud.screenconnect.com/downloads/Mdbg4_x64.exe" -OutFile \ScreenConnect\Mdbg4_x64.exe
}

Catch
{
    $DownloadErr = $_.Exception.Message;
    Write-Log $DownloadErr
    exit;
}
#endregion

#region Do Add Reg Key and File Structure
Write-Log "Beginning Add Reg Key and File Structure Region"
New-Item 'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Force | New-ItemProperty -Name LocalAccountTokenFilterPolicy -Value 1 -Force | Out-Null

if (Test-Path \ScreenConnect) 
{
    Write-Log "Directory structure already configured."
} 

else 
{
    Write-Log "Creating directory structure..."
    New-Item \ScreenConnect -ItemType Directory -Force | Out-Null
    New-Item \ScreenConnect\Instances -ItemType Directory -Force | Out-Null

    If(Test-Path \ScreenConnect\Instances)
    {
        Write-Log "Directory Structure created successfully!"
        $acl = new-object System.Security.AccessControl.DirectorySecurity
        $acl.SetAccessRuleProtection($true, $false)
        $acl.AddAccessRule((new-object System.Security.AccessControl.FileSystemAccessRule("Administrators", "FullControl", "ObjectInherit, ContainerInherit", "None", "Allow")))
        $acl.AddAccessRule((new-object System.Security.AccessControl.FileSystemAccessRule("SYSTEM", "FullControl", "ObjectInherit, ContainerInherit", "None", "Allow")))
        Set-Acl -Path \ScreenConnect -AclObject $acl
    }

    Else
    {
        Write-Log "Failed to create directory structure."
        exit;
    }
}
#endregion

#region Do Add PFX Certificate
Write-Log "Beginning Do Add PFX Certificate Region"

if (Get-ChildItem -Path cert:\LocalMachine\My | ? { $_.Thumbprint -eq $pfxThumbprint }) 
{
    Write-Log "Certificate already installed."
} 

else 
{
    Write-Log "Downloading and installing certificate..."
    
    Try
    {
        Invoke-WebRequest $pfxUrl -OutFile $env:TEMP\cert.pfx
    }

    Catch
    {
        $DownloadErr = $_.Exception.Message;
        Write-Log $DownloadErr
        exit;
    }

    Import-PfxCertificate -FilePath $env:TEMP\cert.pfx -Password (ConvertTo-SecureString $pfxPassword -AsPlainText -Force) cert:\localMachine\my -ErrorVariable $PFXError
    netsh http delete sslcert "ipport=0.0.0.0:443"
    netsh http add sslcert "ipport=0.0.0.0:443" "certhash=$pfxThumbprint" "appid={00000000-0000-0000-0000-000000000000}"
    del $env:TEMP\cert.pfx

    If(!$PFXError)
    {
        Write-Log "Cert Installed successfully!"
    }

    Else
    {
        Write-Log "Cert failed to install! Error was $PfxError"
        exit;
    }
}
#endregion

#region Do Enable PowerShell Remoting
Write-Log "Configuring powershell remoting..."
winrm set winrm/config/service/auth "@{Basic=`"true`"}" -remote:$env:COMPUTERNAME 2>&1 | Out-Null
winrm delete "winrm/config/Listener?Address=*+Transport=HTTPS" -remote:$env:COMPUTERNAME 2>&1 | Out-Null
winrm create "winrm/config/Listener?Address=*+Transport=HTTPS" "@{Port=`"5986`";CertificateThumbprint=`"$pfxThumbprint`"}" -remote:$env:COMPUTERNAME 2>&1 | Out-Null
#endregion

#region Do Create Windows Account
if ([adsi]::Exists("WinNT://" + (hostname) + "/$windowsUserName")) 
{
    Write-Log "Skipping creating windows account."
} 

else 
{
    Write-Log "Creating windows account..."
    net accounts /maxpwage:unlimited

    $windowsUser = ([adsi]("WinNT://" + (hostname))).Create("User", $windowsUserName)
    $windowsUser.SetInfo()

    $administratorsGroup = ([adsi]("WinNT://" + (hostname) + "/Administrators"))
    $administratorsGroup.Add("WinNT://" + (hostname) + "/$windowsUserName")
    $administratorsGroup.SetInfo()
}

if ($windowsPassword -eq "") 
{
    Write-Log "Skipping setting windows password."
} 

else 
{
    Write-Log "Setting windows password..."
    ([adsi]("WinNT://" + (hostname) + "/$windowsUserName")).SetPassword($windowsPassword)
}
#endregion

#region Do Install Cloud Service Files
Write-Log "Installing cloud service files..."

Try
{
    Invoke-WebRequest $cloudServiceUrl -OutFile $env:TEMP\CloudService.zip
}

Catch
{
    $DownloadErr = $_.Exception.Message;
    Write-Log $DownloadErr
    exit;
}

[Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null

Foreach($ServiceDirectory in $($cloudServices.GetEnumerator() | Select -ExpandProperty Value))
{
    Remove-Item "\ScreenConnect\$ServiceDirectory" -Recurse -ErrorAction SilentlyContinue
    New-Item "\ScreenConnect\$ServiceDirectory" -ItemType Directory | Out-Null
    [System.IO.Compression.ZipFile]::ExtractToDirectory("$env:TEMP\CloudService.zip", "\ScreenConnect\$ServiceDirectory")
}
    
Remove-Item $env:TEMP\CloudService.zip

Write-Log "Configuring services..."
foreach ($cloudService in $cloudServices.GetEnumerator()) 
{
    sc.exe delete $cloudService.Name
    sc.exe create $cloudService.Name type= "share" start= "auto" binPath= "C:\ScreenConnect\$($cloudService.Value)\ScreenConnect.CloudService.exe"
    sc.exe failure $cloudService.Name  actions= restart/20000/restart/20000/restart/20000 reset= 86400
}
#endregion

#region Do IP2 Config
if (!$ip2) 
{
    write-log "IP2 Was NOT passed. Going to check for two LAN IPs"
    [Array]$Ip1Split = ($IP1Config.IPv4Address.ipaddress).split(".")

    Foreach($Interface in Get-NetIPConfiguration)
    {
        $InterfaceIpSplit = ($Interface.ipv4address.ipaddress).split(".")
        If($Ip1Split[0] -eq $InterfaceIpSplit[0])
        {
            Write-Log "Second Lan IP Address Identified : Alias - $($Interface.ipv4address.alias) IP - $($Interface.ipv4address.ipaddress)"
            $IP2Config = $Interface
        }
    }

    If(!$IP2Config)
    {
        Write-Log "Unable to find a second lan IP with the same first Octet."
        exit;
    }


} 

write-log "Adding IP2=$ip2 and configuring static with /24 subnet and ADAPTER=$adapterName, IP1=$ip1, SUBNETMASK=$subnetMask, GATEWAY=$gateway, DNS=$dns"
netsh interface ip set address $adapterName static $ip1 $subnetMask $gateway
netsh interface ip add address $adapterName $ip2 $subnetMask

#endregion

#region Do Set DNS and Start Services

Write-Log "Setting DNS..."
netsh interface ip set dns $adapterName static $dns

write-log "Starting services..."

Foreach($Service in $($cloudServices.GetEnumerator() | Select -ExpandProperty Name))
{
    Start-Service $Service

    If((Get-Service $Service).status -ne 'Running')
    {
        write-log "Failed to start the $Service service."
        exit;
    }
}
#endregion