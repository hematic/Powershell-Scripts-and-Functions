#region Function Declarations
######################################
Function Write-Log
{
	Param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[String]$Message
	)

    
	add-content -path $($Paths.logfilepath) -value ($Message)
    Write-Output $Message
}

function Get-LabTechConnection
{
	param 
    (
	    [Parameter(Mandatory = $true, Position = 0)]
	    [string]$Phrase
	)
	
	$connectionObject = New-Object PSObject -Property @{
		host = "localhost"
		User = "root"
		pass = ""
		ltversion = 0
	}
	
	$Version = (Get-ItemProperty "HKLM:\Software\Wow6432Node\LabTech\Agent" -Name Version -ea SilentlyContinue).Version;
	
	#Looks like in 10.5 they decided to remove the version key ...
	$LTAgentVersion = Get-ItemProperty "C:\Program Files\LabTech\ltagent.exe"
	
	if ($LTAgentVersion)
	{
		$Version = $LTAgentVersion.VersionInfo.FileVersion.Substring(0, 7);
	}
	
	if (-not $Version)
	{
		#Try 10.5+ path
		$Version = (Get-ItemProperty "HKLM:\Software\LabTech\Agent" -Name Version -ea SilentlyContinue).Version;
		
		if (-not $Version)
		{
			write-error "Failed to retrieve version."
			return $null;
		}
	}
	
	$LTVersion = [double]$Version
	$connectionObject.ltversion = $LTVersion;
	
	# Check version
	if ($LTVersion -lt [double]105.210)
	{
		write-log "Version is pre 10.5";
		$DatabaseHost = (Get-ItemProperty "HKLM:\Software\Wow6432Node\LabTech\Agent" -Name SQLServer).SQLServer;
		
		if ($DatabaseHost)
		{
			$connectionObject.host = $DatabaseHost;
		}
		
		$connectionObject.pass = (Get-ItemProperty "HKLM:\Software\Wow6432Node\LabTech\Setup" -Name RootPassword -ea SilentlyContinue).RootPassword;
		return $connectionObject;
	}
	else
	{
		write-log "Version is 105 or greater";
		
		$DatabaseUser = (Get-ItemProperty "HKLM:\Software\LabTech\Agent" -Name User -ea SilentlyContinue).User;
		$DatabaseHost = (Get-ItemProperty "HKLM:\Software\LabTech\Agent" -Name SQLServer -ea SilentlyContinue).SQLServer;
		
		if ($DatabaseUser)
		{
			$connectionObject.user = $DatabaseUser;
		}
		
		if ($DatabaseHost)
		{
			$connectionObject.host = $DatabaseHost;
		}
	}
	
	#############################################
	###  Only for 10.5                         ##
	#############################################
	
	# Start with 64-bit location
	
	$CommonPath = "$env:ProgramFiles\LabTech\LabTechCommon.dll";
	
	if (-NOT (Test-Path $CommonPath))
	{
		# try 32-bit location next.
		$CommonPath = "${env:ProgramFiles(x86)}\LabTech Client\LabTechCommon.dll";
		$exists = Test-Path $CommonPath;
	}
	
	# Check to see if we found DLL
	if ($exists -eq $false)
	{
		write-error "Failed to find LabTechCommon library."
		return $null;
	}
	
	try
	{
		[Reflection.Assembly]::LoadFile($CommonPath) | out-null
		write-log "Successfully loaded commonpath";
	}
	catch
	{
		# probably can't find file
		Write-Error -Message "Failed to load LabTechCommon" -Exception System.IO.FileNotFoundException;
		return $null;
	}
	
	# Get txt to decrypt
	if (Test-Path "HKLM:\Software\LabTech\Agent")
	{
		$txtToDecrypt = Get-ItemProperty HKLM:\Software\LabTech\Agent -Name MysqlPass | select -expand MySQLPass;
	}
	else
	{
		$txtToDecrypt = Get-ItemProperty HKLM:\Software\WOW6432Node\LabTech\Agent -Name MysqlPass | select -expand MySQLPass;
	}
	
	write-log "Text to decrypt: $txtToDecrypt"
	
	if (-not $txtToDecrypt)
	{
		Write-Error "Failed to locate mysqlPass key"
		return $null;
	}
	
	[array]$byteArray = @([byte]240, [byte]3, [byte]45, [byte]29, [byte]0, [byte]76, [byte]173, [byte]59);
	
	$lbtVector = [byte[]]$byteArray;
	$cryptoSvcProvider = New-Object System.Security.Cryptography.TripleDESCryptoServiceProvider;
	
	[byte[]]$InputBuffer = [System.Convert]::FromBase64String($txtToDecrypt);
	
	if ($InputBuffer.Length -lt 1)
	{
		write-error "Empty buffer. Cannot decrypt";
		return $null;
	}
	
	$hash = new-object LabTechCommon.clsLabTechHash;
	$hash.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($Phrase));
	$cryptoSvcProvider.Key = $hash.GetDigestBytes();
	$cryptoSvcProvider.IV = $lbtVector;
	
	$access = [System.Text.Encoding]::ASCII.GetString($cryptoSvcProvider.CreateDecryptor().TransformFinalBlock($InputBuffer, 0, $InputBuffer.Length));
	
	if ($access)
	{
		$connectionObject.pass = $access;
		return $connectionObject;
	}
	else
	{
		return $null;
	}
	
}

function Get-SQLResult
{
    param 
    (
	    [Parameter(Mandatory = $true, Position = 0)]
	    [string]$Query
	)

	$result = .\mysql.exe --user="root" --password="$rootpass" --database="LabTech" -e "$query" --batch --raw -N;
	return $result;
}

Function CheckRegKeyExists ($Dir,$KeyName) 
{

	try
    	{
        $CheckIfExists = Get-ItemProperty $Dir $KeyName -ErrorAction SilentlyContinue
        if ((!$CheckIfExists) -or ($CheckIfExists.Length -eq 0))
        {
            return $false
        }
        else
        {
            return $true
        }
    }
    catch
    {
    return $false
    }
	
}

function Download-MySQLExe
{
	
	try
	{
		$DownloadObj = new-object System.Net.WebClient;
		$DownloadObj.DownloadFile($DownloadURL, $MySQLZipPath);
	}
	catch
	{
		$Caughtexception = $_.Exception.Message;
	}
	
	if (!(Test-Path $MySQLZipPath))
	{
		write-log "[DOWNLOAD FAILED] :: Failed to download MySQL ZIP archive! If any exceptions, here they are: $Caughtexception";
		return $false;
	}
	
	# ok, the file exists. Let's ensure that it matches up with our hash.
	# mysql.zip hash
	$ExpectedHash = "40-FD-7B-E8-19-22-99-31-C6-64-D3-0C-46-C1-BF-F2";
	$fileMd5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
	$zipHash = [System.BitConverter]::ToString($fileMd5.ComputeHash([System.IO.File]::ReadAllBytes($MySQLZipPath)))
	
	if ($zipHash -ne $ExpectedHash)
	{
		# Integrity issue. Could be content filtering...
		write-log "[HASH MISMATCH] :: The mysql.zip file's md5 hash does not match the original."
		return $false;
	}
	else
	{
		return $true;
	}
	
	
}

Function Zip-Actions
{
       
       [CmdletBinding(DefaultParameterSetName = 'Zip')]
       param
       (
              [Parameter(ParameterSetName = 'Unzip')]
              [Parameter(ParameterSetName = 'Zip',
                              Mandatory = $true,
                              Position = 0)]
              [ValidateNotNull()]
              [string]$ZipPath,
              [Parameter(ParameterSetName = 'Unzip')]
              [Parameter(ParameterSetName = 'Zip',
                              Mandatory = $true,
                              Position = 1)]
              [ValidateNotNull()]
              [string]$FolderPath,
              [Parameter(ParameterSetName = 'Unzip',
                              Mandatory = $false,
                              Position = 2)]
              [ValidateNotNull()]
              [bool]$Unzip,
              [Parameter(ParameterSetName = 'Unzip',
                              Mandatory = $false,
                              Position = 3)]
              [ValidateNotNull()]
              [bool]$DeleteZip
       )
       
       write-log "Entering Zip-Actions Function."
       
       switch ($PsCmdlet.ParameterSetName)
       {
              'Zip' {
                     
                     If ([int]$psversiontable.psversion.Major -lt 3)
                     {
                           write-log "Step 1"
                           New-Item $ZipPath -ItemType file
                           $shellApplication = new-object -com shell.application
                           $zipPackage = $shellApplication.NameSpace($ZipPath)
                           $files = Get-ChildItem -Path $FolderPath -Recurse
                           write-log "Step 2"
                           foreach ($file in $files)
                           {
                                  $zipPackage.CopyHere($file.FullName)
                                  Start-sleep -milliseconds 500
                           }
                           
                           write-log "Exiting Zip-Actions Function."
                           break           
                     }
                     
                     Else
                     {
                           write-log "Step 3"
                           Add-Type -assembly "system.io.compression.filesystem"
                           $Compression = [System.IO.Compression.CompressionLevel]::Optimal
                           [io.compression.zipfile]::CreateFromDirectory($FolderPath, $ZipPath, $Compression, $True)
                           write-log "Exiting Zip-Actions Function."
                           break
                     }
              }
              
              'Unzip' {

			    $shellApplication = new-object -com shell.application
			    $zipPackage = $shellApplication.NameSpace($ZipPath)
			    $destinationFolder = $shellApplication.NameSpace($FolderPath)
			    $destinationFolder.CopyHere($zipPackage.Items(), 20)
                write-log "Exiting Unzip Section"
				
                        }
       }
       
}

Function Process-Results
{
       
       [cmdletbinding()]

       param
       (
            [Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
            [Array]$Updates
       )


        Foreach($Update in $Updates)
        {
       
            #Eliminate first 2 rows of the log
            ##################################
            If($Update -like '*LabTech Solution Center -*' -or $Update -like '*|1|2|3|4| Item Name*'){}


            #Parse Last Row of the Log
            ##################################
            ElseIf($Update -like '*items installed successfully*')
            {
                $Script:NumGoodUpdates = ([regex]::matches($Update, "(\d*)")).groups[1].value
                $Script:NumTotalUpdates = ([regex]::matches($Update, "(?:\d*\sof\s)(\d*)")).groups[1].value
            }

            #Parse All Other Rows
            ##################################
            ElseIf($Update -like '*|{*')
            {
                $Guid = ([regex]::matches($Update, "(?:\|{)(.*)(?:})")).groups[1].value
                $Col1 = ([regex]::matches($Update, "(?:\|{.+\}\s)(\w)")).groups[1].value
                $Col2 = ([regex]::matches($Update, "(?:\|{.+\}\s\w\s)(\w)")).groups[1].value
                $Col3 = ([regex]::matches($Update, "(?:\|{.+\}\s\w\s\w\s)(\w)")).groups[1].value
                $Col4 = ([regex]::matches($Update, "(?:\|{.+\}\s\w\s\w\s\w\s)(\w)")).groups[1].value
                $Name = ([regex]::matches($Update, "(?:\|{.+\}\s\w\s\w\s\w\s\w\s)(.*)")).groups[1].value

                $Script:ParsedUpdates += New-Object PSObject -Property @{
                Name             = $Name;
		        GUID             = $Guid;
		        Specified        = $Col1;
                Updated          = $Col2;
                WhyUpdateStatus  = $Col3;
                Result           = $Col4;
                }

            }

        }

        Return $Script:ParsedUpdates
     
       
}

Function Download-Patch
{
    Param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$DownloadURL,
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$SavePath 
    )
	try
	{
		$DownloadObj = new-object System.Net.WebClient;
		$DownloadObj.DownloadFile($DownloadURL, $SavePath);
	}
	catch
	{
            $Output = $_.exception | Format-List -force | Out-String
            Write-log "[*ERROR*] : $Output"
	}
}

Function Output-Exception
{
    $Output = $_.exception | Format-List -force | Out-String
    $result = $_.Exception.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($result)
    $UsefulData = $reader.ReadToEnd();

    Write-log "[*ERROR*] : `n$Output `n$Usefuldata "  
}

#endregion

#region Do Variable Declarations

$ErrorActionPreference="Continue"
$Newpassword = 'TestPass1'
$NewServerName = ""
$Paths = New-Object PSObject -Property @{
    LogFilePath = "$Env:windir\temp\script.txt"
    FailedUpdateLog = "$Env:windir\temp\failedupdates.txt"
    UpdateLog = "$Env:windir\temp\marketplaceupdates.txt"
    CommandFile = "$Env:windir\Temp\SCCommandfile.txt"
    PatchSavePath = "$Env:windir\Temp\CurrentPatch.exe"
    PatchResultsPath = "$Env:windir\Temp\LTPatchLog.txt"
}
[STRING]$KeyPhrase = 'Thank you for using LabTech.'
[Array]$ExtraItemsArray = @()
[String]$CustomTableName = "lt_patchinformation"
[String]$SQLDir = "C:\Program Files\MySQL\MySQL Server 5.5\bin"
[String]$PatchDownloadLink = "https://s3.amazonaws.com/labtech-msp/release/LabTechPatch_10.5.2.247.exe"

#endregion

Try
{
    Remove-Item $($Paths.LogFilePath) -Force -ErrorAction SilentlyContinue
    Remove-Item $($Paths.FailedUpdateLog) -Force -ErrorAction SilentlyContinue
    Remove-Item $($Paths.UpdateLog) -Force -ErrorAction SilentlyContinue
    Remove-Item $($Paths.CommandFile) -Force -ErrorAction SilentlyContinue
    Remove-Item $($Paths.PatchSavePath) -Force -ErrorAction SilentlyContinue
    Remove-Item $($Paths.PatchResultsPath) -Force -ErrorAction SilentlyContinue
}

Catch{}

Write-Log "***AMI PREP PROCESS BEGINS***"
Write-Log "#############################"

#region Do Windows Patching
######################################

Write-Log "***Windows Patching BEGINS***"
Write-Log "#############################"

$Source = ‘https://s3.amazonaws.com/ltpremium/modules/PSWindowsUpdate.zip’
$Destination = “$env:temp\PSWindowsUpdate.zip”

Write-Log "Downloading Windows Update PowerShell Module"

Invoke-WebRequest -Uri $Source -OutFile $Destination
Unblock-File $Destination

Write-Log "Unzipping Windows Update PowerShell Module"

Zip-Actions -ZipPath $Destination -FolderPath "$env:windir\System32\WindowsPowerShell\v1.0\Modules\" -Unzip $true -DeleteZip $true | Out-Null;

Write-Log "Importing Windows Update PowerShell Module"

Import-Module PSWindowsUpdate

Write-Log "Enabling the Windows Update Service"

Set-Service 'wuauserv' -StartupType Manual
Start-Service -Name 'wuauserv'

Write-Log "Installing Updates"

$PatchingProcess = Get-WUInstall -MicrosoftUpdate -IgnoreUserInput -AcceptAll -IgnoreReboot

If($PatchingProcess)
{
    Write-Log "($PatchingProcess | FT | Out-String)"
}

Else
{
    Write-Log "No Patches Needed."
}

Write-Log "Windows Patching Complete!"
#endregion

#region Do Latest LabTech Patch
#####################################

Write-Log "***LabTech Patching BEGINS***"
Write-Log "#############################"

#Get Root Pass
###########################
$rootpass = (get-itemproperty "HKLM:\SOFTWARE\Wow6432Node\LabTech\Setup").rootpassword

If($rootpass -eq $Null -or $rootpass -eq "")
{
    Write-log -Message "Unable to retrieve root password"
    Return "Unable to retrieve root password"
    exit;
}

#Kill the LTClient Process
###########################
IF(Get-process -Name 'LTClient' -ErrorAction SilentlyContinue)
{
    Stop-Process -Name 'LTClient' -Force
    Write-log -Message "The LTClient process has been killed!"
}

#Download the Patch
###########################
Set-Location "$Env:Windir\"
Download-Patch -DownloadURL $PatchDownloadLink -SavePath $($Paths.PatchSavePath)

IF(-not (Test-Path $($Paths.PatchSavePath)))
{
    Write-log "Failed to download the patch."
    Return "Failed to download the patch."
    exit;
}

#Run the Patch
###########################
$AllArgs = "/t 360 /p 360"
Start-Process -FilePath "$($Paths.PatchSavePath)" -ArgumentList $AllArgs -Wait -WindowStyle Hidden
$LogFileResults = Get-content -Path $($Paths.PatchResultsPath)

#Check For the new Table
###########################

set-location "$sqldir";

$TableQuery = @"
SELECT * 
FROM information_schema.tables
WHERE table_schema = 'LabTech' 
    AND table_name = `'$CustomTableName`'
LIMIT 1;
"@

$TableCheck = Get-SQLresult -query $TableQuery

If($TableCheck -eq $null)
{
    Write-log "Unable to find $CustomTableName in the database."
    [bool]$TableResult = $False
}

Else
{
    Write-log "Found $CustomTableName in the database."
    [bool]$TableResult = $True
}

If($LogFileResults -match "LabTech Server has been successfully updated" -and $TableResult -eq $True)
{
    Write-log "Patch was Successful"
}

Else
{
    Write-log "Patch Failed"
    Return "Patch Failed"
}

Write-Log "LabTech Patching Complete!"
#endregion

#region Do Solution Center Updates
#####################################

Write-Log "***Solution Patching BEGINS***"
Write-Log "#############################"

####This needs to be re-enabled once Tyrone has distributed the new marketplace.exe out there.#####
<#Write-Log "Downloading the newest version of the Marketplace EXE"

$Source = ‘http://labtech-msp.com/contrib/solutioncenter/105/LTMarketplace.exe’ #This needs to change for LT 11
$Destination = “C:\Program Files (x86)\LabTech Client\ltmarketplace.exe"
Invoke-WebRequest -Uri $Source -OutFile $Destination
#>

#Add all the custom items we want.
$ExtraItemsArray += New-Object PSObject -Property @{
    Name = 'Active Directory'
    Guid = '368b2769-4bd4-4408-92aa-08f1fe32f7d6'
}

$ExtraItemsArray += New-Object PSObject -Property @{
    Name = 'Deployment manager Dashboard'
    Guid = '5fb33973-c1ec-4cfb-b74b-aebea1c7032b'
}

$ExtraItemsArray += New-Object PSObject -Property @{
    Name = 'LT License Manager'
    Guid = 'b13fe3a1-0afd-4d58-a628-0b4632de886f'
}

$ExtraItemsArray += New-Object PSObject -Property @{
    Name = 'ScreenConnect'
    Guid = 'ccf607eb-1c6d-4b29-b68c-7e127477a7c1'
}

$ExtraItemsArray += New-Object PSObject -Property @{
    Name = 'Virtualization'
    Guid = '58d5204e-a51a-4f28-af0d-e563517c0cce'
}

$ExtraItemsArray += New-Object PSObject -Property @{
    Name = 'BlackListed Events'
    Guid = '08a92050-9af4-4f7c-954c-9969c2afd239'
}

$ExtraItemsArray += New-Object PSObject -Property @{
    Name = 'Database Management Pack'
    Guid = '23c9ddef-d410-4150-9239-6c7f18a8dba6'
}

$ExtraItemsArray += New-Object PSObject -Property @{
    Name = 'Messaging Management Pack'
    Guid = '91bf72c8-6718-45f8-8915-2782f69d701b'
}

$ExtraItemsArray += New-Object PSObject -Property @{
    Name = 'Web/Proxy Management Pack'
    Guid = '29486994-5a76-4658-bea8-2c1008078812'
}

[STRING]$ExtraItems = $ExtraItemsArray | % {"I:`{$($_.guid)`};"}
[ARRAY]$Script:ParsedUpdates = @()
[ARRAY]$FailedUpdates = @()


#Get the LT Share directory from the DB
######################################

set-location "$sqldir";

$LtShareQry = @"
SELECT `localltshare` FROM `config`
"@

$LTSharedir = get-sqlresult -query $Ltshareqry

If($LtshareDir -eq $null)
{
    Write-log "Unable to retrieve the Local LT Share path from the database."
    exit;
}

Else
{
    Write-log "Local LT Share : $LTShareDir"
}

#Set the proper file and folder permissions.
######################################
$ltsharedir = $ltsharedir.trimend("\")
attrib -r $Ltsharedir
icacls $ltsharedir\* /T /Q /C /RESET
Set-Location "C:\Program Files (x86)\LabTech Client"
attrib -R *.* /S
Set-Location  "C:\ProgramData\LabTech Client\Logs"
attrib -R *.* /S

#Run the solution center update process.
######################################

#We add our extra items to the commandfile that marketplace.exe will use.
Add-Content -Path $($Paths.CommandFile) -Value $ExtraItems

#Declare the arguments to use with Start-process
$AllArgs = "/update /fix /commandfile $($Paths.CommandFile)"

#Call marketplace.exe in a hidden window and wait for it to finish. We send all output to $UpdateLog
Start-Process -FilePath "${env:ProgramFiles(x86)}\LabTech Client\LTMarketplace.exe" -ArgumentList $AllArgs -PassThru -RedirectStandardOutput $($Paths.UpdateLog) -Wait -WindowStyle Hidden

#Grab the results form the log for parsing.
$updateResults = Get-content $($Paths.UpdateLog)

#Error Checking to validate the EXE ran correctly
######################################

If(!$updateResults)
{
    Write-log 'The marketplace update log was not generated.'
    exit;
}

Else
{
    If($updateResults | Where-object { $_ -match 'items installed successfully'})
    {
        Write-log "Update log contains valid results. Moving on."
    }

    Else
    {
        Write-log "Something is wrong with the update log."
        exit;
    }
}

#Parse the update results into an object
######################################

Process-Results $updateResults

#Find Failed Updates
######################################

$Script:ParsedUpdates | Where-Object {$_.Result -eq 'F'} | Select-Object -ExpandProperty Name | Add-Content -Path $($Paths.failedupdatelog) -ErrorAction SilentlyContinue

If ((Test-Path $($Paths.failedupdatelog)) -eq $False)
{
    Add-Content -Path $($Paths.LogFilePath) -Value "Complete Success"
}

ELSE
{
    Add-Content -Path $($Paths.LogFilePath) -Value "Check FailedUpdates Log"
}

Write-Log "Solution Patching Complete!"

#endregion

#region Do Rename Computer

Write-Log "Renaming Computer to $Newservername"
Rename-Computer -NewName $NewServerName -Force 

#endregion

#region Do Password Resets

Write-Log "Beginning Password Resets"

set-location "$sqldir";

$PasswordSetQuery = @"
UPDATE `users` 
SET `password`=(AES_ENCRYPT('$($NewPassword)',`userid`)) 
WHERE name = 'LTAdmin';
"@

Write-Log "Setting LTAdmin Password"

$PasswordSet = Get-SQLresult -query $PasswordSetQuery

Write-Log "Verifying LTAdmin Password"

$PasswordVerifyQuery = @"
SELECT (AES_DECRYPT(`password`,`userid`)) 
FROM `users` 
WHERE name = 'LTAdmin'
"@

$PasswordVerify = Get-SQLresult -query $PasswordVerifyQuery

If($PasswordVerify -eq $Newpassword)
{
    Write-log "Verification was successful. Password was changed to $Newpassword"
}

Else
{
    Write-log "Verification FAILED. Response from SQL was $Passwordverify"
    Return "Failed to change LTAdmin Password" 
}

Write-Log "Setting Local Server Account Password"

$ChangeLocalPass = net user LTAdmin $Newpassword


If($ChangeLocalPass -like "*completed successfully*")
{
    Write-log "Local User Account password was changed successfully."
}

Else
{
    Write-log "Verification FAILED. Response from NET USER was $Changelocalpass"
    Return "Failed to change Local Server Account Password" 
}
#endregion

Write-Log "Script Complete!"
Return "Success"

#Possible Return Errors
<#
LabTech Patching Errors
-------------------------
"Unable to retrieve root password"
"Failed to download the patch."
"Patch Failed"

Solution Center Updates Errors
-------------------------------
"Unable to retrieve the Local LT Share path from the database."
'The marketplace update log was not generated.'
"Something is wrong with the update log."

Password Reset Errors
---------------------
"Failed to change LTAdmin Password"
"Failed to change Local Server Account Password"
#>