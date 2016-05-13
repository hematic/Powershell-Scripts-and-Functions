<#######################################################################################>
#Function Declarations
Function Check-Location
{
	Param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[String]$Servername
	)

    $LocationCheck = $False

    $LocCode = $Servername.substring(0,3)
    $LocCode = $LocCode.ToUpper()
    
    If($LocationCodes.containsvalue($LocCode) -eq $True)
    {
        $LocationCheck = $True
    }

    Return $LocationCheck
}

Function Check-Role
{
	Param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[String]$Servername
	)

    $RoleCheck = $False

    $RoleCode = $ServerName.substring(3,2)
    $RoleCode = $RoleCode.ToUpper()

    If($RoleCodes.containsvalue($RoleCode) -eq $True)
    {
        $RoleCheck = $True
    }
    
    eLSE
    {
        Write-output ""
    }
    Return $RoleCheck
}

Function Check-Function
{
	Param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[String]$Servername
	)
    $FunctionCheck = $False

    $FunctionCode = $ServerName.substring(5,2)
    $FunctionCode = $FunctionCode.ToUpper()

    $RoleCode = $ServerName.substring(3,2)
    $RoleCode = $RoleCode.ToUpper()

    If($($FunctionCodes.code) -contains $FunctionCode)
    {
        $Matches = $FunctionCodes | where-object {$_.code -eq $FunctionCode -and $_.appliesto -contains $RoleCode}
        [Int]$Count = $Matches | measure-object | select -ExpandProperty count
        If($Count -gt 0)
        {
            $FunctionCheck = $true 
        }    
    }
    Return $FunctionCheck
}

Function Check-Index
{
	Param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[String]$Servername
	)
    $IndexCheck = $False

    $IndexCode = $ServerName.substring(7,3)
    
    If([Microsoft.VisualBasic.Information]::IsNumeric($IndexCode) -eq $True)
    {
        $IndexCheck = $True
    }
    
    Return $IndexCheck
}

Function Build-FunctionCodes
{
    [Array]$FunctionCodes = @()

    $ServerFunction = New-Object PSObject -Property @{
        Code = 'AM'
        Description = 'Alert Management'
        AppliesTo = @('ap')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
    Code = 'AP'
    Description = 'Application Server'
    AppliesTo = @('BK', 'HR', 'FM', 'FN', 'DS', 'MF', 'DM', 'TC')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'AR'
        Description = 'Attorney Resources'
        AppliesTo = @('ap')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'AV'
        Description = 'AntiVirus Servers'
        AppliesTo = @('AP', 'SC', 'ST')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'BH'
        Description = 'Bridgehead Servers'
        AppliesTo = @('MS')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
    Code = 'BK'
    Description = 'Backup Servers'
    AppliesTo = @('AP', 'FM')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'CS'
        Description = 'Certificate Services'
        AppliesTo = @('AP', 'DS')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
    Code = 'CC'
    Description = 'Content Crawler/ Communication Server'
    AppliesTo = @('IS')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'CO'
        Description = 'Console'
        AppliesTo = @('MO')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'DB'
        Description = 'Database License Engine or Component'
        AppliesTo = @('AP', 'CL', 'HR', 'FM', 'FN', 'DM', 'TC', 'MO')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
    Code = 'DC'
    Description = 'Domain Controller'
    AppliesTo = @('DS')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'DH'
        Description = 'Data Process'
        AppliesTo = @('DS')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'DP'
        Description = 'DHCP'
        AppliesTo = @('AP','DM')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'DR'
        Description = 'Disaster Recovery'
        AppliesTo = @('AP', 'DM', 'MS', 'TC')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
    Code = 'EM'
    Description = 'Enterprise Mobility'
    AppliesTo = @('MO')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'EX'
        Description = 'Extranet Servers'
        AppliesTo = @('AP', 'CM', 'HR', 'FM', 'FN', 'DM')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'FS'
        Description = 'File Servers'
        AppliesTo = @('CL', 'HR', 'FM', 'FN', 'DM', 'TC')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'FT'
        Description = 'FTP Servers'
        AppliesTo = @('AP')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
    Code = 'GW'
    Description = 'Gateway Servers'
    AppliesTo = @('MS', 'TC')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'HO'
        Description = 'Home Areas'
        AppliesTo = @('HR', 'FM', 'FN', 'FS')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
    Code = 'IM'
    Description = 'Instant Messaging'
    AppliesTo = @('MS')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'IN'
        Description = 'Intranet'
        AppliesTo = @('AP', 'CM', 'DM')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'IT'
        Description = 'Information Technology'
        AppliesTo = @('AP')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
    Code = 'IX'
    Description = 'Indexer'
    AppliesTo = @('AP','CM','DM')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'MB'
        Description = 'Mailbox Servers'
        AppliesTo = @('CL','MS')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'MD'
        Description = 'Monitoring Device'
        AppliesTo = @('AP')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
    Code = 'NM'
    Description = 'Network Management'
    AppliesTo = @('AP')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'PA'
        Description = 'Public Access'
        AppliesTo = @('FS')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'PF'
        Description = 'Public Folders'
        AppliesTo = @()
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'PC'
        Description = 'User Sighting Systems'
        AppliesTo = @('AP')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
    Code = 'PM'
    Description = 'Physical Machine'
    AppliesTo = @()
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'PR'
        Description = 'Proxy Servers'
        AppliesTo = @('AP')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
    Code = 'PS'
    Description = 'Print Servers'
    AppliesTo = @('CL','AP','FM')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'PV'
        Description = 'Provisioning'
        AppliesTo = @('RA','AP')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'PX'
        Description = 'Proxy'
        AppliesTo = @('MO')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
    Code = 'RA'
    Description = 'Remote Applications'
    AppliesTo = @('AP', 'HR', 'FM', 'FN', 'DM', 'TC')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'SA'
        Description = 'Security Access'
        AppliesTo = @('AP')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'SC'
        Description = 'Security Control'
        AppliesTo = @('AP', 'TC')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
    Code = 'SD'
    Description = 'Software Distribution'
    AppliesTo = @('AP','TC')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'SM'
        Description = 'Software Management'
        AppliesTo = @('AP', 'TC', 'RA')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'SS'
        Description = 'Support Services'
        AppliesTo = @('AP')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'SV'
        Description = 'Service'
        AppliesTo = @()
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
    Code = 'TC'
    Description = 'Telecommunications'
    AppliesTo = @('AP', 'MS')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'TM'
        Description = 'Time Management'
        AppliesTo = @('AP')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
    Code = 'TS'
    Description = 'Terminal Servers'
    AppliesTo = @('AP', 'HR', 'FM', 'FN', 'DM')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'WR'
        Description = 'Wireless Handheld Server'
        AppliesTo = @('AP', 'MS')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'VM'
        Description = 'Virtual Machine Host/vCenter'
        AppliesTo = @('AP','MF')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
    Code = 'WA'
    Description = 'Wireless Application'
    AppliesTo = @('AP')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'WB'
        Description = 'Web Systems'
        AppliesTo = @('AP', 'HR', 'FM', 'FN', 'MS', 'DM', 'TC')
    }
    $FunctionCodes += $ServerFunction
    $ServerFunction = New-Object PSObject -Property @{
        Code = 'WS'
        Description = 'Web Services'
        AppliesTo = @('AP')
    }
    $FunctionCodes += $ServerFunction
    
    Return $FunctionCodes    
}

<#######################################################################################>
#Variable Declarations
Add-Type -Assembly Microsoft.VisualBasic
[Array]$BadServers = @()
[Hashtable]$LocationCodes = @{
    "Abu Dhabi"        = "ABU"
    "Almaty"           = "ALM"
    "Americas"         = "AM1"
    "Ankara"           = "ANK"
    "Asia Pacific"     = "AP1"
    "Astana"           = "AST"
    "Beijing"          = "BEI"
    "Berlin"           = "BER"
    "Bratislava"       = "BRA"
    "Brussels"         = "BRU"
    "Doha"             = "DOH"
    "Dresden"          = "DRE"
    "Dubai"            = "DUB"
    "Düsseldorf"       = "DUS"
    "EMEA"             = "EM1"
    "Flensburg"        = "FLE"
    "Frankfurt"        = "FRA"
    "Geneva"           = "GEN"
    "Hamburg"          = "HAM"
    "Helsinki"         = "HEL"
    "Hong Kong"        = "HON"
    "Istanbul"         = "IST"
    "Jakarta"          = "JAK"
    "Johannesburg"     = "JOH"
    "London"           = "LON"
    "Los Angeles"      = "LOS"
    "Madrid"           = "MAD"
    "Manila"           = "MAN"
    "Mexico City"      = "MEX"
    "Miami"            = "MIA"
    "Milan"            = "MIL"
    "Moscow"           = "MOS"
    "Munich"           = "MUN"
    "New York City"    = "NYC"
    "Palo Alto"        = "PAL"
    "Paris"            = "PAR"
    "Prague"           = "PRA"
    "Riyadh"           = "RIY"
    "Sao Paulo"        = "SAO"
    "Seoul"            = "SEO"
    "Shanghai"         = "SHA"
    "Singapore"        = "SIN"
    "Stockholm"        = "STO"
    "Tampa"            = "TAM"
    "Tokyo"            = "TOK"
    "Warsaw"           = "WAR"
    "Washington D.C."  = "WDC"
}
[Hashtable]$RoleCodes = @{
    "Applications"                            = "AP"
    "Backup Systems"                          = "BK"
    "Conflicts / New Business"                = "CF"
    "Cluster Systems"                         = "CL"
    "Content Management"                      = "CM"
    "Document Management"                     = "DM"
    "Directory Services"                      = "DS"
    "Datawarehouse"                           = "DW"
    "Extended Storage"                        = "ES"
    "Finance / Marketing"                     = "FM"
    "Finance"                                 = "FN"
    "File Share Systems"                      = "FS"
    "Human Resources"                         = "HR"
    "Integration"                             = "IG"
    "Index Search"                            = "IS"
    "Log management"                          = "LM"
    "Litigation Support"                      = "LS"
    "Multifunction Systems"                   = "MF"
    "Mobility"                                = "MO"
    "Messaging Systems"                       = "MS"
    "Network Monitoring / System Monitoring"  = "NM"
    "Paper Less"                              = "PL"
    "Quality Management"                      = "QM"
    "Remote Access"                           = "RA"
    "Reporting Services"                      = "RS"
    "Security Control Systems"                = "SC"
    "Support Services"                        = "SS"
    "Storage Systems"                         = "ST"
    "Thin Client"                             = "TC"
    "Web Systems"                             = "WB"
}
$FunctionCodes = Build-FunctionCodes

<#######################################################################################>
#Main Script
$Servers = Get-ADComputer -LDAPFilter “(&(objectcategory=computer)(OperatingSystem=*server*))”

If(!$Servers)
{
    Write-Output "Gathering list of servers from Active Directory failed."
    exit;  
}

Foreach($Server in $Servers)
{
    $ServerName = $Server.Name

    
    If($($ServerName.substring(0,3)) -eq 'DR-' -or $($ServerName.substring(0,3)) -eq 'dr-')
    {
        $ServerName = $ServerName.TrimStart('DR-')
        $ServerName = $ServerName.TrimStart('dr-')
    }

    $Location = Check-Location -Server $ServerName
    If($Location -eq $False)
    {
        $BadServers += $Server
    Write-output "server : $Servername"
        Write-output "Failed Location"
        continue;
    }

    $Role =     Check-Role -ServerName $ServerName
    If($Role -eq $False)
    {
        $BadServers += $Server
    Write-output "server : $Servername"
        Write-output "Failed Role"
        continue;
    }


    $Function = Check-Function -ServerName $ServerName
    If($Function -eq $False)
    {
        $BadServers += $Server
    Write-output "server : $Servername"
        Write-output "Failed Function"
        continue;
    }

    $Index =    Check-Index -ServerName $ServerName
    If($Index -eq $False)
    {
        $BadServers += $Server
    Write-output "server : $Servername"
        Write-output "Failed Index"
        continue;
    }
}
