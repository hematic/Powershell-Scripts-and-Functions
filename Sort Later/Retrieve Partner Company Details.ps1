<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2015 v4.2.89
	 Created on:   	8/11/2015 4:07 PM
	 Created by:   	Marcus Bastian
	 Organization: 	LabTech Software / ConnectWise
	 Filename:     Retrieve Partner Company Details.ps1	
	===========================================================================
	.DESCRIPTION
		The purpose of this script is to take in API details and a partner CD key.

		The script will make the necessary API calls to the CP/API in order to
		output the following information:

		1. Company Rec ID
		2. Company ID
		3. Company Name
		4. City

		
#>

$TimeNow = Get-Date
#$VerbosePreference = 'Inquire'

function Get-PartnerCompanyData
{
		param
		(
			[parameter(Mandatory = $true)]
			[string]
			$base64AuthInfo
	)
	
	$RestArgs = @{
	Uri = "https://$API_FQDN/v1/licensing/accounts/$CDKey"
	ContentType = "application/json"
	method = 'get'
	Headers = @{ Authorization = ("Basic {0}" -f $base64AuthInfo) }
	TimeoutSec = 60
	errorVariable = 'RestError'
	ea = 'silentlyContinue'
	}
	
	try
	{
		$CompanyInfo = Invoke-RestMethod @RestArgs;
	}
	catch
	{
		$StatusCode = $_.Exception.Response.StatusCode.value__
		$thisCaughtException = $_.Exception.Message
		return "Status Code was:  $StatusCode`nException Encountered: $thisCaughtException";
	}
	
	return $CompanyInfo;
}

function Get-PartnerRecID
{
	param
	(
		[parameter(Mandatory = $true)]
		[string]
		$RecID,
		[parameter(Mandatory = $true)]
		[string]
		$base64AuthInfo
	)
	
	$RestArgs = @{
		Uri = "https://$API_FQDN/v1/lumi/companies/$RecID"
		ContentType = "application/json"
		method = 'get'
		Headers = @{ Authorization = ("Basic {0}" -f $base64AuthInfo) }
		TimeoutSec = 60
		errorVariable = 'RestError'
		ea = 'silentlyContinue'
	}
	
	try
	{
		$CompanyID = Invoke-RestMethod @RestArgs;
	}
	catch
	{
		$StatusCode = $_.Exception.Response.StatusCode.value__
		$thisCaughtException = $_.Exception.Message
		return "Status Code was:  $StatusCode`nException Encountered: $thisCaughtException";
	}
	
	return $CompanyID;
	
}

Function Write-Log
{
	# i.e:    write-log $msg [1]    <---- optional last param
	Param
		(
		[Parameter(Mandatory = $True, Position = 0)]
		[String]$Message,
		[Parameter(Mandatory = $False, Position = 1)]
		[INT]$Severity
	)
	
	$Note = "[NOTE]"
	$Warning = "[WARNING]"
	$Problem = "[ERROR]"
	
	[string]$Date = get-date
	
	switch ($Severity)
	{
		1 { $Msg = ($Date + "`t:`t" + $Note + $Message); }
		2 { $Msg = ($Date + "`t:`t" + $Warning + $Message); }
		3 { $Msg = ($Date + "`t:`t" + $Problem + $Message); }
		default { $Msg = ($Date + "`t:`t" + $Message); }
	}
	
	add-content -path $LogFilePath -value $Msg
	write-output $Msg;
	
}

$ErrorActionPreference = 'SilentlyContinue';

$API_FQDN 		= $args[0];
$CDKey 			= $args[1];
$API_Usr 		= $args[2];
$Auth_Pass		= $args[3];

#$test = $true;

if ($Test -eq $true)
{
	$API_FQDN 	= "api-discovery-doc.labtechsoftware.com"
	$CDKey = "7992110634414005"
	$CDKey_Hash 	= "e3a99cb4602f7dbdef629130eabfc8998b1111525c2596b5b6f7837b"
	$API_Usr 	= "infrastructure"
	$Auth_Pass = "ECko40n6U5rC8V9nZI0jCqre"
}

# log file used for output
$logFilePath= "$env:windir\temp\$cdkey-GetCompanyInfoError.log";
if (test-path $logFilePath) { Remove-Item $logFilePath }


# build authorization header
$authString = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $API_Usr, $Auth_Pass)))

$PartnerCompanyData = Get-PartnerCompanyData $authString;
	
if (!$PartnerCompanyData)
{
	Write-Log "Failed to gather company data. Response was $PartnerCompanyData" 3;
	exit;
}
elseif ($PartnerCompanyData.GetType().Name -eq 'String')
{
	# passed exception back	as string
	Write-Log "Exception encountered... $PartnerCompanyData" 3;
	exit;
}
else
{
	# json returned. Let's drill into its properties
	write-log "Company RecID: $($PartnerCompanyData.Data.Company_RecID)" 1;
	write-log "City: $($PartnerCompanyData.Data.City)" 1 
	write-log "Company Name: $($PartnerCompanyData.Data.company)" 1;
	
	$RecID 	 = $PartnerCompanyData.Data.Company_RecID;
	$City 	 = $PartnerCompanyData.Data.City;
	$Company = $PartnerCompanyData.Data.Company;
	#$PartnerCompanyData.data
	
	Write-Log "Successfully pulled data from API/CP!" 1;
	
	if (!$City)
	{
		$City = "Partner Location"
		write-log "City value was defaulted to 'Partner Location' because no city value was returned." 2
	}
}

$CompanyID = Get-PartnerRecID $RecID $authString;

if (!$CompanyID)
{
	write-log "Failed to gather company data" 3;
	exit;
}
elseif ($CompanyID.GetType().Name -eq 'String')
{
	# passed exception back	as string
	write-log "Exception encountered while trying to retrieve CompanyId from LUMI endpoint... `n`n`t$CompanyID" 3;
	exit;
}
else
{
	write-log "Company ID returned was:  $($CompanyID.Data.CompanyID)" 1;
	$CompanyID = $CompanyID.Data.CompanyID;
}
#[[RecID:103276]][[CompanyName:The WebMaster's Inc.]][[CompanyID:WedmastersInc]][[City:Tampa]]
$outputFormat = "[[RecID:{0}]][[CompanyName:{1}]][[CompanyID:{2}]][[City:{3}]]" -f $RecID, $Company, $CompanyID, $City
write-log $outputFormat 1;

#notepad $logFilePath
