#Set-ExecutionPolicy Unrestricted -Force; (New-Object System.Net.WebClient).DownloadString("https://cloud.screenconnect.com/downloads/ConfigureServer.ps1") > \ConfigureServer.ps1; \ConfigureServer.ps1

Param 
(
  #  [parameter(Mandatory = $true)] $ip2, This may not be needed
  #  [parameter(Mandatory = $true)] $pfxPassword,
     [parameter(Mandatory = $true)] $windowsPassword
)

#region Do Variable Set

Write-Output "Beginning Variable Set Region"
[String]$AdapterName = (Get-NetAdapter | Select -First 1 -ExpandProperty Name)
[String]$Gateway = (Get-WmiObject -Class Win32_IP4RouteTable | ? { $_.destination -eq '0.0.0.0' -and $_.mask -eq '0.0.0.0'} | Sort-Object metric1 | % {$_.nexthop })
[String]$dns = "209.244.0.3" #level3
[Array]$ip1 = (Get-NetIPConfiguration | ? { $_.IPv4DefaultGateway.NextHop -eq $gateway } | % { $_.IPv4Address.IPAddress })
[Array]$subnetMask = (Get-WmiObject Win32_NetworkAdapterConfiguration | ? { $_.IPSubnet.Length } | % { $_.IPSubnet[0] })
[String]$pfxThumbprint = "E31975030B72A245878BBC336BA87CF1DEF0C3D8"
[String]$pfxUrl = "https://cloud.screenconnect.com/downloads/ScreenConnectComWildcard.pfx"
[String]$cloudServiceUrl = "https://cloud.screenconnect.com/downloads/CloudService.zip"
[HashTable]$cloudServices = @{"ScreenConnect Router" = "Router"; "ScreenConnect Instance Manager" = "InstanceManager" }
[String]$windowsUserName = "screenconnect"
[Array]$rdpIPs = @("96.10.31.0/24", "63.145.136.0/24", "70.46.245.0/24")
#Set-PSDebug -Trace 1
# this makes downloads 100x faster
$progressPreference = 'SilentlyContinue'
#endregion

#region Do Configure Time
Write-Output "Beginning Time Server Configuration Region"
w32tm /config /manualpeerlist:pool.ntp.org /syncfromflags:MANUAL 2>&1 | Out-Null
Stop-Service w32time

If((Get-Service w32time).status -ne 'Stopped')
{
    Write-output "Unable to stop the w32time service."
    exit;
}

Start-Service w32time

If((Get-Service w32time).status -ne 'Running')
{
    Write-output "Unable to start the w32time service."
    exit;
}
#endregion

#region Do Configure PageFile
# TODO this is kind of a problem because it really needs a reboot for this
Write-Output "Setting page file to auto-managed"
$System = gwmi Win32_ComputerSystem -EnableAllPrivileges
$System.AutomaticManagedPagefile = $True
$System.Put() | Out-Null
#endregion

#region Do Stop Services
Write-Output "Beginning the Stop Services Region"
Foreach($Service in $($cloudServices.GetEnumerator() | Select -ExpandProperty Name))
{ 
    Stop-Service $Service -ErrorAction SilentlyContinue 

    If((Get-Service $Service).status -ne 'Stopped')
    {
        Write-output "Unable to stop the $Service service."
        exit;
    }
}
#endregion

#region Do Configure Windows Firewall
Write-Output "Beginning the Configure Windows Firewall Region"

if (Get-NetFirewallRule | ? { $_.DisplayName -eq "Allow 80 Inbound" }) 
{
    Write-Output "Firewall already configured."
} 

else 
{
    Write-Output "Configuring firewall main ports..."
    New-NetFirewallRule -DisplayName "Allow 80 Inbound" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 80 | Out-Null
    New-NetFirewallRule -DisplayName "Allow 443 Inbound" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 443 | Out-Null
    New-NetFirewallRule -DisplayName "Allow 5986 Inbound" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 5986 | Out-Null
}

if (!(Get-NetFirewallRule | ? { $_.DisplayName -eq "Allow 3389 Inbound" })) 
{
    Write-Output "Configuring firewall RDP..."
    New-NetFirewallRule -DisplayName "Allow 3389 Inbound" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 3389 -RemoteAddress $rdpIPs | Out-Null
}

Write-Output "Disabling blanket RDP and management firewall rules..."
Get-NetFirewallRule | ? { $_.DisplayName.StartsWith("Remote Desktop") } | Disable-NetFirewallRule
Get-NetFirewallRule | ? { $_.DisplayName.StartsWith("Windows Remote Management") } | Disable-NetFirewallRule
Get-NetFirewallRule | ? { $_.DisplayName.StartsWith("Hyper-V") } | Disable-NetFirewallRule
Get-NetFirewallRule | ? { $_.DisplayName.StartsWith("RDP") } | Disable-NetFirewallRule
Get-NetFirewallRule | ? { $_.DisplayName.StartsWith("SSH") } | Disable-NetFirewallRule
#endregion

<# we worked around this
Write-Output "Adding desktop experience for VFW encoding..."
Install-WindowsFeature Desktop-Experience
#>

#region Do Download Debuggers
Write-Output "Beginning Download debuggers Region"

Try
{
    Invoke-WebRequest "https://cloud.screenconnect.com/downloads/Mdbg4_x86.exe" -OutFile \ScreenConnect\Mdbg4_x86.exe
}

Catch
{
    $DownloadErr = $_.Exception.Message;
    Write-output $DownloadErr
    exit;
}

Try
{
    Invoke-WebRequest "https://cloud.screenconnect.com/downloads/Mdbg4_x64.exe" -OutFile \ScreenConnect\Mdbg4_x64.exe
}

Catch
{
    $DownloadErr = $_.Exception.Message;
    Write-output $DownloadErr
    exit;
}
#endregion

<# this would be nice, but OVH doesn't have it
Install-WindowsFeature DNS
#>

#region Do Add Reg Key and File Structure
# not 100% sure this will help with anything but to keep servers consistent
Write-Output "Beginning Add Reg Key and File Structure Region"
New-Item 'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Force | New-ItemProperty -Name LocalAccountTokenFilterPolicy -Value 1 -Force | Out-Null

if (Test-Path \ScreenConnect) 
{
    Write-Output "Directory structure already configured."
} 

else 
{
    Write-Output "Creating directory structure..."
    New-Item \ScreenConnect -ItemType Directory -Force | Out-Null
    New-Item \ScreenConnect\Instances -ItemType Directory -Force | Out-Null

    If(Test-Path \ScreenConnect\Instances)
    {
        Write-Output "Directory Structure created successfully!"
        $acl = new-object System.Security.AccessControl.DirectorySecurity
        $acl.SetAccessRuleProtection($true, $false)
        $acl.AddAccessRule((new-object System.Security.AccessControl.FileSystemAccessRule("Administrators", "FullControl", "ObjectInherit, ContainerInherit", "None", "Allow")))
        $acl.AddAccessRule((new-object System.Security.AccessControl.FileSystemAccessRule("SYSTEM", "FullControl", "ObjectInherit, ContainerInherit", "None", "Allow")))
        Set-Acl -Path \ScreenConnect -AclObject $acl
    }

    Else
    {
        Write-Output "Failed to create directory structure."
        exit;
    }
}
#endregion

<#
if (Get-ChildItem -Path cert:\LocalMachine\My | ? { $_.Thumbprint -eq $pfxThumbprint }) {
    Write-Output "Certificate already installed."
} else {
    Write-Output "Downloading and installing certificate..."
    Invoke-WebRequest $pfxUrl -OutFile $env:TEMP\cert.pfx
    Import-PfxCertificate -FilePath $env:TEMP\cert.pfx -Password (ConvertTo-SecureString $pfxPassword -AsPlainText -Force) cert:\localMachine\my | Out-Null
    netsh http delete sslcert "ipport=0.0.0.0:443"
    netsh http add sslcert "ipport=0.0.0.0:443" "certhash=$pfxThumbprint" "appid={00000000-0000-0000-0000-000000000000}"
    del $env:TEMP\cert.pfx
}
#>

#region Do Enable PowerShell Remoting
Write-Output "Configuring powershell remoting..."
#Enable-PSRemoting -Force -ErrorAction SilentlyContinue
#goddamn -remote:computername seems to be required and I don't know why
winrm set winrm/config/service/auth "@{Basic=`"true`"}" -remote:$env:COMPUTERNAME 2>&1 | Out-Null
winrm delete "winrm/config/Listener?Address=*+Transport=HTTPS" -remote:$env:COMPUTERNAME 2>&1 | Out-Null
winrm create "winrm/config/Listener?Address=*+Transport=HTTPS" "@{Port=`"5986`";CertificateThumbprint=`"$pfxThumbprint`"}" -remote:$env:COMPUTERNAME 2>&1 | Out-Null
#endregion

#region Do Create Windows Account
if ([adsi]::Exists("WinNT://" + (hostname) + "/$windowsUserName")) 
{
    Write-Output "Skipping creating windows account."
} 

else 
{
    Write-Output "Creating windows account..."
    net accounts /maxpwage:unlimited

    $windowsUser = ([adsi]("WinNT://" + (hostname))).Create("User", $windowsUserName)
    $windowsUser.SetInfo()

    $administratorsGroup = ([adsi]("WinNT://" + (hostname) + "/Administrators"))
    $administratorsGroup.Add("WinNT://" + (hostname) + "/$windowsUserName")
    $administratorsGroup.SetInfo()
}

if ($windowsPassword -eq "") 
{
    Write-Output "Skipping setting windows password."
} 

else 
{
    Write-Output "Setting windows password..."
    ([adsi]("WinNT://" + (hostname) + "/$windowsUserName")).SetPassword($windowsPassword)
}
#endregion

#region Do Install Cloud Service Files
Write-Output "Installing cloud service files..."

Try
{
    Invoke-WebRequest $cloudServiceUrl -OutFile $env:TEMP\CloudService.zip
}

Catch
{
    $DownloadErr = $_.Exception.Message;
    Write-output $DownloadErr
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

Write-Output "Configuring services..."
foreach ($cloudService in $cloudServices.GetEnumerator()) 
{
    sc.exe delete $cloudService.Name
    sc.exe create $cloudService.Name type= "share" start= "auto" binPath= "C:\ScreenConnect\$($cloudService.Value)\ScreenConnect.CloudService.exe"
    sc.exe failure $cloudService.Name  actions= restart/20000/restart/20000/restart/20000 reset= 86400
}
#endregion

<# Don't think this is needed. Must test.
if ($ip2 -eq "") {
    Write-Output "Skipping setting IP."
} else {
    Write-Output "Adding IP2=$ip2 and configuring static with /24 subnet and ADAPTER=$adapterName, IP1=$ip1, SUBNETMASK=$subnetMask, GATEWAY=$gateway, DNS=$dns"
    netsh interface ip set address $adapterName static $ip1 $subnetMask $gateway
    netsh interface ip add address $adapterName $ip2 $subnetMask
}
#>

#region Do Set DNS and Start Services

Write-Output "Setting DNS..."
netsh interface ip set dns $adapterName static $dns


Write-Output "Starting services..."

Foreach($Service in $($cloudServices.GetEnumerator() | Select -ExpandProperty Name))
{
    Start-Service $Service

    If((Get-Service $Service).status -ne 'Running')
    {
        Write-output "Failed to start the $Service service."
        exit;
    }
}
#endregion

<#
    #New-Service -Name "ScreenConnect Router" -BinaryPathName "" -StartupType
    #New-Service -Name "ScreenConnect Instance Manager" -BinaryPathName "" -StartupType


    # doens't work, have to add it for each site (which we do in code now)
    #netsh http add urlacl url=https://+:443/ user=\Everyone

    # powershell stuff here has ... problems ... to say the least
    #$adapter = Get-NetAdapter -Name Ethernet
    #$adapter | Set-NetIPInterface -DHCP Enabled
    #Remove-NetRoute -DestinationPrefix 0.0.0.0/0 -Confirm:$false
    #$adapter | Remove-NetIPAddress -Confirm:$false
    #$adapter | Set-NetIPInterface -DHCP Disabled
    #$adapter | New-NetIPAddress -IPAddress $ip1 -PrefixLength 24
    #$adapter | New-NetIPAddress -IPAddress $ip2 -PrefixLength 24
    #$adapter | New-NetRoute -DestinationPrefix 0.0.0.0/0 -NextHop 192.168.2.1

    #netsh firewall set opmode disable
    #Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False
    #$dns = (Get-NetIPConfiguration | ? { $_.IPv4DefaultGateway.NextHop -eq $gateway } | % { $_.DNSServer } | % { $_.ServerAddresses }),
#>