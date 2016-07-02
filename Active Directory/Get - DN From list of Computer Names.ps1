#Requires -version 3
#requires -module ActiveDirectory

<#
.Synopsis
	Script gets the DN/ADSPath, IP Address and online state of a list of computers.  
	It accepts a list of NETBIOS names or FQDNs. Searches Forest root by default.

.PARAMETER PCNameOrFile
	Required input object, list of computers.

.PARAMETER UseMyDomain
	If present, searches AD for computers in current domain.

.PARAMETER UseParentDomain
	If present, searches AD for computers in parent of the current domain.

.DESCRIPTION
   This script gets the DistinguishedName (less CN=) and online state of a list of computers. It is multi-threaded using workflows. Active Directory queries can be directed to the current domain, parent domain or AD Root.

.NOTES
	N/A

.EXAMPLE
	$computers = "C:\Users\MyName\Desktop\ping.txt"
	$data = Get-ComputerDN -PCNameOrFile $computers 
	$data | Export-Csv -NoTypeInformation -Path $logfile -force
#>

WorkFlow Get-Info{
	param
	(
		[string[]]$computers,
		$server
    )
	
	ForEach -parallel ($Computer in $computers)
	{
		if ($computer.Length -gt 1)
		{
			InLineScript
			{
             	#begin by checking to see if list is FQDNs
				if (($using:Computer).Contains("."))
				{
                	$DNShostname = ($using:Computer).Trim()
                	$NBTName = (($using:Computer).Split(".")[0]).Trim()
				}
				
				ELSE
				{
					$NBTName = ($using:Computer).trim()
					$DNSHostName = ""
				}
				
				$ErrorActionPreference = "SilentlyContinue"
            	#this is much faster then loading AD Module in each spawned process
	            $de = New-Object System.DirectoryServices.DirectoryEntry("LDAP://"+$using:Server)
	            $ds = New-Object System.DirectoryServices.DirectorySearcher
	            $ds.SearchRoot = $de
	            $ds.Filter = "(&(objectCategory=computer)(objectClass=computer)(Name=$NBTName))"
	            $ds.SearchScope = "SubTree"
	            $retval = $ds.FindOne()

            	#Always use name from AD if available
				if ($retval)
				{
					$dnsHostName = ($retval.Properties).dnshostname
				}
				
				#If IP Address returned, use NBTName instead
				if ([Net.IPAddress]::TryParse($dnshostname, [ref]$null))
				{
					$dnshostname = $NBTName
				}
				
				If ($retval.path.Length -gt 0)
				{
                	$adsPath = ($retval.Path).replace("LDAP://"+$Server+"/CN=","")
				}
				
				ELSE
				{
					$adsPath = "Not Found"
				}
				
				if ($DNShostname)
				{
					$PingName = $DNShostname
				}
				
				Else
				{
					$PingName = $NBTName
				}
				
				$objPing = get-wmiobject win32_pingstatus -filter "Address = '$PingName' AND ResolveAddressNames='true'"
            	$IPAddress = ($objPing.IPV4Address).IPAddressToString
				
				if ($IPAddress -eq $null)
				{
					$IPAddress = "Not Found"
				}
				
				if ($objPing.StatusCode -eq $null)
				{
                	$PingStatusCode = $objPing.PrimaryAddressResolutionStatus
				}
				
				ELSE
				{
					$PingStatusCode = $objPing.StatusCode
				}
				
				$PingStatusText = switch ($PingStatusCode)
				{
	                      0 {"Online"}
	                  11001 {"Buffer Too Small"}
	                  11002 {"Destination Net Unreachable"}
	                  11003 {"Destination Host Unreachable"}
	                  11004 {"Host Not Found"} # Destination Protocol Unreachable
	                  11005 {"No Host Record"} # Destination Port Unreachable
	                  11006 {"No Resources"}
	                  11007 {"Bad Option"}
	                  11008 {"Hardware Error"}
	                  11009 {"Packet Too Big"}
	                  11010 {"Request Timed Out"}
	                  11011 {"Bad Request"}
	                  11012 {"Bad Route"}
	                  11013 {"TimeToLive Expired Transit"}
	                  11014 {"TimeToLive Expired Reassembly"}
	                  11015 {"Parameter Problem"}
	                  11016 {"Source Quench"}
	                  11017 {"Option Too Big"}
	                  11018 {"Bad Destination"}
	                  11032 {"Negotiating IPSEC"}
	                  11050 {"General Error"}
	                default {"Host Not Found"}
                }

        		#if we have the DNSName from AD, use it, else from Ping, else set to NBTName
				if (($DNSHostName).length -eq 0)
				{
					$DNSHostName = $objPing.ProtocolAddressResolved.toString().ToUpper()
				}
				
				if (($DNSHostName).length -eq 0)
				{
					$DNShostname = $NBTName
				}

		        $output = [PSCustomObject]@{
		            SystemName = $NBTName
		            DNSName = $DNSHostName
		            IPAddress = $IPAddress
		            ADSPath= $adsPath
		            PingReply = $PingStatusText
		            PSComputerName = $PSComputerName
		            PSSourceJobInstanceId = $PSSourceJobInstanceId
		            }
        		Return $output

        	}#End InlineScript
      	} #End check for empty lines
    }#End for Each Parallel
}#End Get-Info 

function Get-ComputerDN
{
    [CmdletBinding(DefaultParameterSetName="RootDomain")]
    Param
    (
        # PCNameOrFile is the input object.  Accepts list, text file, filter from AD
        [Parameter(Mandatory=$True,
                   ValueFromPipeline=$true)]
        $PCNameOrFile,
       
       #using ParameterSet to force choice but default is defined at top to be RootDomain
       [Parameter(ParameterSetName='MyDomain')]
          [switch] $UseMyDomain,

       [Parameter(ParameterSetName='ParentDomain')]
          [switch]$UseParentDomain,

         [Parameter(ParameterSetName='RootDomain')]
          [switch]$UseRootDomain
    )

    switch ($PCNameOrFile)
    {
        {$PCNameOrFile -is [array]}
          {$PCList = $PCNameOrFile.Split(","); Break}
        {$PCNameOrFile.contains(",")} 
          {$PCList = $PCNameOrFile.Split(","); Break}
        {Test-path $PCNameOrFile} 
          {$PCList = Get-Content $PCNameOrFile | sort ;Break}
    }
    
    cls
    Import-Module ActiveDirectory
    $MyDomain = Get-ADDomain

    #Use DNSRoot of local domain as server name for UseMyDomain
    if($UseMyDomain) {$Server = $MyDomain.DNSRoot.ToString()}
    ELSE{
        #The AD forest domain is used by default
        $SearchRoot = $($MyDomain).Forest.ToString()

        #UseParentDomain: Search using a GC in parent domain
        if($UseParentDomain){$SearchRoot = $($MyDomain).ParentDomain.ToString()}

        Write-Host "Getting list of all GC Server(s) for $SearchRoot"
        $gc = get-addomaincontroller -server $SearchRoot -Filter { isGlobalCatalog -eq $true}
        if ($gc.count -eq 0){
            Write-Warning "No Global Catalog Server found in $SearchRoot"
            $server = $SearchRoot
        }ELSE{
            #This syntax handles result whether or not it is an array
            $server = Get-ClosestServer $gc.hostname
            #3268 is port for Global Catalog
            $Server = $server+ ":3268"
            Write-Host "Using $Server for LDAP queries"`n
        }
    }

    $script:iPCcount = $PCList.count
    Write-Host $iPCcount multi-threaded queries running...
    $Start = Get-Date
    #remove the workflow generated properties from returned data
    $outObj = Get-Info $PCList $server| select * -ExcludeProperty PS* 
    $End = Get-Date
    $ElapsedMin = [System.Math]::Round(($End - $Start).totalMinutes,1)
    Write-Host Elapsed time in minutes $ElapsedMin minutes for $script:IPCcount computers
    Return $outObj
}


Workflow Test-WFConnection {
  param(
    [string[]]$Computers
  )
  foreach -parallel ($computer in $computers) {
    Test-Connection -ComputerName $computer -Count 1 -ErrorAction SilentlyContinue
  }
}

Function Get-ClosestServer{  param(
    [string[]]$Computers
  )
    Write-Host "... Now Finding Closest GC Server for $SearchRoot"
    $PingInfo = Test-WFConnection $Computers
 
    #get system with lowest responsetime
    $PingInfo | sort-object ResponseTime | select -expandproperty address -First 1
}

######### Example Usage  #######

$logfile = "$env:userprofile\Desktop\ADSInfo.csv" 
$computers = "$env:userprofile\Desktop\ping.txt"

$data = Get-ComputerDN -PCNameOrFile $computers  |
Export-Csv -NoTypeInformation -Path $logfile -force
Write-host You can open the log with the command
Write-host `ii $LogFile