<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2015 v4.2.89
	 Created on:   	9/2/2015 3:33 PM
	 Revision:      2.0.0
	 Created by:   	Marcus Bastian and Phillip Marshall
	 Organization: 	ConnectWise
	 Filename:     	Perform Installation of LabTech Server.ps1
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>

<#######################################################################################>
#Function Declarations

Function StartProcWithWait 
{

	$InstallApplication = $args[0];
    $InstallProc = New-Object System.Diagnostics.Process;
	$InstallProc.StartInfo = $InstallApplication;
	$InstallProc.Start();
	$InstallProc.WaitForExit();
	
}

Function Zip-Actions
{
       <#
       .SYNOPSIS
              A function to zip or unzip files.
       
       .DESCRIPTION
              This function has 3 possible uses.
              1) Zip a folder or files and save the zip to specified location.
              2) Unzip a zip file to a specified folder.
              3) Unzip a zip file and delete the original zip when complete.       
      
 
       .PARAMETER ZipPath
              The full path of the file to unzip or the full path of the zip file to be created.
       
       .PARAMETER FolderPath
              The path to the files to zip or the path to the directory to unzip the files to.
       
       .PARAMETER Unzip
              If $true the function will perform an unzip instead of a zip
       
       .PARAMETER DeleteZip
              If set to $True the zip file will be removed at then end of the unzip operation.
       
       .EXAMPLE
              PS C:\> Zip-Actions -ZipPath 'C:\Windows\Temp\ziptest.zip' -FolderPath 
              PS C:\> Zip-Actions -ZipPath 'C:\Windows\Temp\ziptest.zip' -FolderPath 'C:\Windows\Temp\ZipTest' -Unzip $true
              PS C:\> Zip-Actions -ZipPath 'C:\Windows\Temp\ziptest.zip' -FolderPath 'C:\Windows\Temp\ZipTest' -Unzip $true -DeleteZip $True

       
       .NOTES
              Additional information about the function.
#>
       
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
       
       Write-Output "Entering Zip-Actions Function." -Severity 1
       
       switch ($PsCmdlet.ParameterSetName)
       {
              'Zip' {
                     
                     If ([int]$psversiontable.psversion.Major -lt 3)
                     {
                           
                           New-Item $ZipPath -ItemType file
                           $shellApplication = new-object -com shell.application
                           $zipPackage = $shellApplication.NameSpace($ZipPath)
                           $files = Get-ChildItem -Path $FolderPath -Recurse
                           
                           foreach ($file in $files)
                           {
                                  $zipPackage.CopyHere($file.FullName)
                                  Start-sleep -milliseconds 500
                           }
                           
                           Write-Output "Exiting Zip-Actions Function." -Severity 1
                           break           
                     }
                     
                     Else
                     {
                           
                           Add-Type -assembly "system.io.compression.filesystem"
                           $Compression = [System.IO.Compression.CompressionLevel]::Optimal
                           [io.compression.zipfile]::CreateFromDirectory($FolderPath, $ZipPath, $Compression, $True)
                           Write-Output "Exiting Zip-Actions Function." -Severity 1
                           break
                     }
              }
              
              'Unzip' {
                     <#
                     Add-Type -assembly "system.io.compression.filesystem"
                     $Compression = [System.IO.Compression.CompressionLevel]::Optimal
                     [io.compression.zipfile]::ExtractToDirectory($ZipPath, $FolderPath)
                     
                     If ($DeleteZip) { Remove-item $ZipPath }
                     
                     Write-Output "Exiting Zip-Actions Function." -Severity 1
                                  break
                     #>
                     
                     $shellApplication = new-object -com shell.application 
                     $zipPackage = $shellApplication.NameSpace($ZipPath)
                     $destinationFolder = $shellApplication.NameSpace($FolderPath) 
                     $destinationFolder.CopyHere($zipPackage.Items(), 20)
       
              
                     
                     
                     
              }
       }
       
}

Function Process-Features
{

	
	param
		(
		[parameter(Mandatory = $true)]
		[String]$OS
	    )
        
        Switch ($OS)
        {
            2008R2
                  {
                        foreach($Feature in $2008R2Features)
                        {
                            Write-Output "Starting install of $($Feature)"
                            $Install = Add-WindowsFeature $Feature

                            If($install.success -ne $True)
                            {
                                $RolesThatFailedToInstall += $Feature
                                Write-output "ERROR: Feature $($Feature) failed to install." 
                            }

                            Else
                            {
                                Write-output "SUCCESS: Feature $($Feature) Installed." 
                            }
                        }	
                  }

            2012R2
                  {
                        foreach($Feature in $2012R2features)
                        {
                            Write-Output "Starting install of $($Feature)"
                            $Install = Add-WindowsFeature $Feature

                            If($install.success -ne $True)
                            {
                                $RolesThatFailedToInstall += $Feature
                                Write-output "ERROR: Feature $($Feature) failed to install." 
                            }

                            Else
                            {
                                Write-output "SUCCESS: Feature $($Feature) Installed." 
                            }
                        }		
                  }

            2012Standard  
                  {
                        foreach($Feature in $2012features)
                        {
                            Write-Output "Starting install of $($Feature)"
                            $Install = Add-WindowsFeature $Feature

                            If($install.success -ne $True)
                            {
                                $RolesThatFailedToInstall += $Feature
                                Write-output "ERROR: Feature $($Feature) failed to install." 
                            }

                            Else
                            {
                                Write-output "SUCCESS: Feature $($Feature) Installed." 
                            }
                        }	
                  }
            
        }

}

Function Write-Log
{

	
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
		1 { add-content -path $LogFilePath -value ($Date + "`t:`t" + $Note + $Message) }
		2 { add-content -path $LogFilePath -value ($Date + "`t:`t" + $Warning + $Message) }
		3 { add-content -path $LogFilePath -value ($Date + "`t:`t" + $Problem + $Message) }
		default { add-content -path $LogFilePath -value ($Date + "`t:`t" + $Message) }
	}
	
	
}

Function End-Script
{	
	param
		(
		[parameter(Mandatory = $true)]
		[String]$Result
	)
	
	$PsErrors = $Error
	Out-File -InputObject $Result -Filepath $ResultsPath
	Out-File -InputObject $PsErrors -FilePath $PSErrorsPath
	Write-Log ("********************************")
	Write-Log ("***** $($ScriptName) Ends *****")
	Write-Log ("********************************")
	exit;
}


<#######################################################################################>
#Variable Declarations

$ErrorActionPreference = 'silentlycontinue';
$LabTechCDKey = '@cdkey@';
$InstallErrorPath = "$env:windir\temp\ltinstall\InstallError.txt"
$WorkingDir = "$env:windir\temp\LTInstall"
$LogfilePath = "$WorkingDir\InstallLog.txt"
$AdminPasswordPath = "$WorkingDir\AdminPass.txt"
$InstallDownload = "http://ltpremium.s3.amazonaws.com/lt_installs/latest_release/installation_assemblies.zip"
$InstallSavePath = "$WorkingDir\LtInstall.zip"
$SqlYogDownload = "http://automationfiles.hostedrmm.com/third_party_apps/sqlyog_103/SQLyog-10.3.0-1Community.exe"
$SqlYogSavePath = "$WorkingDir\SQLyog-10.3.0-1Community.exe"
$ResultsPath = "$WorkingDir\LtInstallResults.txt"
$PSErrorsPath = "$workingDir\LTInstall_PSErrors.txt"
$RootUser = $args[0];
$RootPassword = $args[1];
$SQLServerAddress = $args[2];

# array that will be used to store any role names that fail to install.
$RolesThatFailedToInstall = @()

# List of features to be installed for Server 2012 R2
$2012R2features = @("AS-Web-Support","AS-WAS-Support","AS-WAS-Support","Web-WebServer","Web-Common-Http",
                    "Web-Static-Content","Web-Dir-Browsing","Web-Http-Errors","Web-Http-Redirect",
                    "Web-App-Dev","Web-Asp-Net","Web-Asp-Net45","Web-Net-Ext","Web-Net-Ext45",
                    "Web-ASP","Web-ISAPI-Ext","Web-ISAPI-Filter","Web-Health","Web-Http-Logging","Web-Performance",
                    "Web-Basic-Auth","Web-Windows-Auth","Web-Performance","Web-Stat-Compression","Web-Dyn-Compression",
                    "Web-Mgmt-Tools","Web-Mgmt-Console","Web-Scripting-Tools","Web-Mgmt-Service","Web-Mgmt-Compat",
                    "Web-Metabase","Web-WMI","Web-Lgcy-Scripting","Web-Lgcy-Mgmt-Console","NET-HTTP-Activation",
					"WAS", "WAS-Process-Model", "WAS-NET-Environment", "WAS-Config-APIs")

# List of features to be installed for Server 2012
$2012features = @("Application-Server","AS-Web-Support","AS-WAS-Support","Web-WebServer",
                    "Web-Common-Http","Web-Static-Content","Web-Dir-Browsing","Web-Http-Errors",
                    "Web-Http-Redirect","Web-App-Dev","Web-Asp-Net","Web-Asp-Net45",
                    "Web-Net-Ext","Web-Net-Ext45","Web-ASP","Web-ISAPI-Ext","Web-ISAPI-Filter",
                    "Web-Health","Web-Http-Logging","Web-Performance","Web-Basic-Auth",
                    "Web-Windows-Auth","Web-Stat-Compression","Web-Dyn-Compression",
                    "Web-Mgmt-Tools","Web-Mgmt-Console","Web-Mgmt-Service","Web-Mgmt-Compat","Web-Metabase",
                    "Web-WMI","Web-Lgcy-Scripting","Web-Lgcy-Mgmt-Console","NET-HTTP-Activation",
					"WAS", "WAS-Process-Model", "WAS-NET-Environment", "WAS-Config-APIs")

# List of features to be installed for Server 2008 R2

$2008R2Features = @("Application-Server","AS-Web-Support","AS-WAS-Support","AS-WAS-Support","Web-WebServer",
                    "Web-Common-Http","Web-Static-Content","Web-Dir-Browsing","Web-Http-Errors","Web-Http-Redirect",
                    "Web-App-Dev","Web-Asp-Net","Web-Net-Ext","Web-ASP","Web-ISAPI-Ext","Web-ISAPI-Filter",
                    "Web-Health","Web-Http-Logging","Web-Performance","Web-Basic-Auth","Web-Windows-Auth",
                    "Web-Performance","Web-Stat-Compression","Web-Dyn-Compression","Web-Mgmt-Tools",
                    "Web-Mgmt-Console","Web-Scripting-Tools","Web-Mgmt-Service","Web-Mgmt-Compat","Web-Metabase",
                    "Web-WMI","Web-Lgcy-Scripting","Web-Lgcy-Mgmt-Console","NET-Framework","NET-Framework-Core",
                    "NET-Win-CFAC","NET-HTTP-Activation","WAS","WAS-Process-Model","WAS-NET-Environment","WAS-Config-APIs")

<#######################################################################################>
#Clean up the directory structure and create a new directory.

<#
If(Test-Path $WorkingDir)
{
   	Write-Output "Failed to clean the working directory.";
    End-Script -Result 'Failure24'; 
}

#>
New-Item $WorkingDir -type Directory;

If (!(Test-Path $WorkingDir))
{
	Write-Output "Failed to create working directory.";
    End-Script -Result 'Failure01';
}


<#######################################################################################>
#Downloads the Install Assemblies

Write-Log "Beginning download of the assemblies."

try
{
	$DownloadObj = new-object System.Net.WebClient;
	$DownloadObj.DownloadFile($InstallDownload, $InstallSavePath);
}
catch
{
	$DownloadErr = $_.Exception.Message;
}
finally
{
	if(!(Test-Path $InstallSavePath))
	{
		Write-Output "Failed to download assemblies. Download exception was: $DownloadErr";
		Set-Content $WorkingDir\DownloadErrors.txt $DownloadErr;
    	End-Script -Result 'Failure02';
	}
}

<#######################################################################################>
#Extracts the Install Assemblies and Checks them

Write-Log "Beginning extraction of the assemblies."

Zip-Actions -ZipPath $InstallSavePath -FolderPath $WorkingDir -Unzip $true

$SizeOfFilesInFolder = ((get-childitem $WorkingDir | Measure-Object Length -sum) | `
							select @{name='Sum'; expression={[math]::ceiling($_.Sum / 1024 / 1024)}}).Sum;
	

Write-Output "Verifying that LTInstall folder has extracted files."							
							
IF($SizeOfFilesInFolder -gt 650)
{
 	Write-Output "All files successfully extracted."
}
ELSE
{
	Write-Output "Failed to extract all files. Contents of LTInstall directory were only: $($SizeOfFilesInFolder)";
    End-Script -Result 'Failure03';
}

<#######################################################################################>
#Adds Needed Firewall Rules for LabTech

Write-Output "Adding Firewall Rules."	

REG.exe ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}" /v "IsInstalled" /t REG_DWORD /d 0 /f
REG.exe ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}" /v "IsInstalled" /t REG_DWORD /d 0 /f

netsh.exe advfirewall firewall add rule name="LabTech Redirector Port" dir=in description="Required port for LabTech Software functionality" action=allow protocol=tcp localport=70-75
netsh.exe advfirewall firewall add rule name="LabTech Redirector Port" dir=in description="Required port for LabTech Software functionality" action=allow protocol=udp localport=70-75
netsh.exe advfirewall firewall add rule name="IIS Web Port" dir=in description="Required port for LabTech Software functionality" action=allow protocol=tcp localport=80
netsh.exe advfirewall firewall add rule name="IIS SSL Web Port" dir=in description="Required port for LabTech Software functionality" action=allow protocol=tcp localport=443
netsh.exe advfirewall firewall add rule name="LabTech Control Center Database Port" dir=in description="Required port for LabTech Software functionality" action=allow protocol=tcp localport=3306
netsh.exe advfirewall firewall add rule name="LabTech VNC Control Center Ports" dir=in description="Required port for LabTech Software functionality" action=allow protocol=tcp localport=40000-41000
netsh.exe advfirewall firewall add rule name="LabTech Mediator Port" dir=in description="Required port for LabTech Software functionality" action=allow protocol=tcp localport=8002
netsh.exe advfirewall firewall add rule name="LabTech Mediator Port" dir=in description="Required port for LabTech Software functionality" action=allow protocol=udp localport=8002

<#######################################################################################>
#Adds Needed Windows Features

$OS = (Get-WmiObject Win32_OperatingSystem).caption

Import-Module ServerManager;

if($OS -like '*2012 R2*')
{
    Process-Features -OS '2012R2'
}

IF($OS -like '*2012*')
{
    Process-Features -OS '2012Standard'
}

IF($OS -like '*2008*')
{
    Process-Features -OS '2008R2'
}


if ($RolesThatFailedToInstall.count -eq 0)
{
	write-output 'All required server roles were successfully installed!'
}

else
{
	$FailedRoles = $RolesThatFailedToInstall -join ", ";
    Write-Log "The following roles failed to install: $FailedRoles."
    out-file -FilePath "$Workingdir\FailedRoles.txt" -InputObject $FailedRoles
    End-Script -Result 'Failure04';
}


<#######################################################################################>
#Install Visual C++ 2005

	Write-Output "Beginning Installation of Visual C++ 2005 Redistributable ........."

$InstallVCRED = New-Object System.Diagnostics.ProcessStartInfo("$WorkingDir\vcredist_x86.exe", "/Q");
$installVcred.UseShellExecute = $False 
StartProcWithWait $InstallVCRED;

	if(!(gwmi -cl win32_product | where-object { $_.Name -like '*Microsoft Visual C++ 2005*' })) 
	{ 
	    Write-Log "FAILED to install Microsoft Visual C++ 2005 Redistributable ............";
	    End-Script -Result 'Failure05';
	}

	ELSE
	{
	    Write-Log "SUCCESSFULLY installed Microsoft Visual C++ 2005 Redistributable .......";
	}

<#######################################################################################>
#Install Visual C++ 2010 32 Bit

	Write-Output "Beginning Installation of Visual C++ 2010(32 Bit) Redistributable ........."

$InstallVCRED = New-Object System.Diagnostics.ProcessStartInfo("$WorkingDir\vcredist_x86_2010.exe", "/Q");
$installVcred.UseShellExecute = $False 
StartProcWithWait $InstallVCRED;

	if(!(gwmi -cl win32_product | where-object { $_.Name -like '*Microsoft Visual C++ 2010  x86 Redistributable*' })) 
	{ 
	    Write-Log "FAILED to install Microsoft Visual C++ 2010(32 Bit) Redistributable ............";
	    End-Script -Result 'Failure05';
	}

	ELSE
	{
	    Write-Log "SUCCESSFULLY installed Microsoft Visual C++ 2010(32 Bit) Redistributable .......";
	}

<#######################################################################################>
#Install Visual C++ 2010 64 Bit

	Write-Output "Beginning Installation of Visual C++ 2010(64-Bit) Redistributable ........."

$InstallVCRED = New-Object System.Diagnostics.ProcessStartInfo("$WorkingDir\vcredist_x64.exe", "/Q");
$installVcred.UseShellExecute = $False 
StartProcWithWait $InstallVCRED;

	if(!(gwmi -cl win32_product | where-object { $_.Name -like '*Microsoft Visual C++ 2010  x64 Redistributable*' })) 
	{ 
	    Write-Log "FAILED to install Microsoft Visual C++ 2010(64 Bit) Redistributable ............";
	    End-Script -Result 'Failure06';
	}

	ELSE
	{
	    Write-Log "SUCCESSFULLY installed Microsoft Visual C++ 2010(64 Bit) Redistributable .......";
	}

<#######################################################################################>
#Install Crystal Reports

	Write-Output "Beginning Installation of Crystal Reports 2008 Runtime SP2 ........"

$InstallCrystal = New-Object System.Diagnostics.ProcessStartInfo("$ENV:windir\system32\msiexec.exe", "/i $WorkingDir\CRRuntime_12_2_mlb.msi /qn /norestart");
$InstallCrystal.UseShellExecute = $False 
StartProcWithWait $InstallCrystal;

	if(!(gwmi -cl win32_product | where-object { $_.Name -like '*Crystal Reports 2008 Runtime SP2*' })) 
	{ 
	    Write-Log "FAILED to install Crystal Reports 2008 Runtime SP2 ................";
	    End-Script -Result 'Failure07';
	}
	ELSE
	{
	    Write-output "SUCCESSFULLY installed Crystal Reports 2008 Runtime SP2 ...........";
	}

<#######################################################################################>
#Install .NET 4.0 Extended

if($OS -notlike '*2012*') # 2008 R2 ONLY. 4.0 comes with 4.5 server role on Server 2012+
{
	Write-Output "Beginning installation of .NET Framework 4.0 extended .............";

$InstallDotNet4 = New-Object System.Diagnostics.ProcessStartInfo("$WorkingDir\dotnetfx40_full_x86_x64.exe", "/q /norestart");
$InstallDotNet4.UseShellExecute = $False 
StartProcWithWait $InstallDotNet4;

	if(!(gwmi -cl win32_product | where-object { $_.Name -like '*Microsoft .NET Framework 4*' })) 
	{ 
	    Write-Log "FAILED to install .NET Framework 4.0 ..............................";
	    End-Script -Result 'Failure08';
	}

	ELSE
	{
	    Write-output "SUCCESSFULLY installed .NET Framework 4.0 .........................";
	}
}

<#######################################################################################>
#Install MySQL ODBC
	
Write-Log "Beginning installation of MySQL ODBC Connector(32bit) 5.3.4 ...............";

$InstallMySQLODBC32Bit = New-Object System.Diagnostics.ProcessStartInfo("$ENV:windir\system32\msiexec.exe ", "/i $WorkingDir\mysql-connector-odbc-5.3.4-win32.msi /qn /norestart");
$InstallMySQLODBC32Bit.UseShellExecute = $False 
StartProcWithWait $InstallMySQLODBC32Bit;

If(!(GP HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.Displayname -like '*MySQL Connector/ODBC 5.3*'}))
{ 
    Write-Log "FAILED to install MySQL ODBC Connector ............................";
    End-Script -Result 'Failure09';
}

Write-Log "Beginning installation of MySQL ODBC Connector(64bit) 5.3.4 ...............";

$InstallMySQLODBC64Bit = New-Object System.Diagnostics.ProcessStartInfo("$ENV:windir\system32\msiexec.exe ", "/i $WorkingDir\mysql-odbc-5.3.4-winx64.msi /qn /norestart");
$InstallMySQLODBC64Bit.UseShellExecute = $False 
StartProcWithWait $InstallMySQLODBC64Bit;

If(!(gp HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.Displayname -like '*MySQL Connector/ODBC 5.3*'}))
{ 
    Write-Log "FAILED to install MySQL ODBC Connector ............................";
    End-Script -Result 'Failure10';
}

Write-Log "Beginning installation of MySQL NET Connector(32bit) 6.9.7 ...............";

$InstallMySQLODBCNET = New-Object System.Diagnostics.ProcessStartInfo("$ENV:windir\system32\msiexec.exe ", "/i $WorkingDir\mysql-net-6.9.7.msi /qn /norestart");
$InstallMySQLODBCNET.UseShellExecute = $False 
StartProcWithWait $InstallMySQLODBCNET;

If(!(GP HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.Displayname -like '*MySQL Connector Net 6.9.*'}))
{ 
    Write-Log "FAILED to install MySQL Connector Net 6.9 ............................";
    End-Script -Result 'Failure11';
}

ELSE
{
    Write-Log "SUCCESSFULLY installed MySQL Connector Net 6.9  .......................";
}

<#######################################################################################>
#Install LabTech Server

Write-Output "Beginning installation of LabTech Server ..........................";

$InstallLabTech = New-Object System.Diagnostics.ProcessStartInfo("$ENV:windir\system32\msiexec.exe", "/i `"$WorkingDir\InstallServer.msi`" /qn /norestart /L*V `"$WorkingDir\LT_Install_Log.txt`" CDKEYTEXT=`"$($LabTechCDKey)`" MYSQLSERVERADDRESS=`"$SQLServerAddress`" MYSQLROOTUSER=`"$RootUser`" MySqlRootPassword=`"$RootPassword`" LABTECHIGNITE=1 LTADMINWEBACCESS=1");
$InstallLabTech.UseShellExecute = $False 
StartProcWithWait $InstallLabTech;

<#######################################################################################>
#Check for Errors

$InstallLog = Get-Content -Path "$WorkingDir\LT_Install_Log.txt" | out-string

if(!(gwmi -cl win32_product | where-object { $_.Name -like '*LabTech® Software*' })) 
	{ 
        If (Test-Path -Path "$WorkingDir\LT_Install_Log.txt")
        {
            $ErrorRegex = "(?:INSTALLATION ERROR BEGIN[\s][\s\S]+[\d\d:]+\d+.[\d]+:\s)([\s\S]+)(?:Action\sLog\s)"
            $InstallError = ([regex]::matches($InstallLog, $ErrorRegex)).groups[1].value
            If($InstallError)
            {
                Out-File -FilePath $InstallErrorPath -InputObject $InstallError
                Write-Log "FAILED to install LabTech Server ..................................";
                End-Script -Result 'Failure12';
            }
            Write-Log "No Install Error Could Be Parsed .................................."
            Write-Log "FAILED to install LabTech Server ..................................";
	        End-Script -Result 'Failure13';
        }

        Else
        {
            Write-Log "No Install Log was found .................................."
            Write-Log "FAILED to install LabTech Server ..................................";
	        End-Script -Result 'Failure14';
        }
	}
	ELSE
	{
	    Write-Log "SUCCESSFULLY installed LabTech Server .............................";
        Write-Log "Retrieving LTAdmin Password from the Install Log."
        $PasswordRegex = "(?:LTADMINPASSWORD\s=\s)(.+)"
        $AdminPassword = ([regex]::matches($InstallLog, $PasswordRegex)).groups[1].value
        If($AdminPassword)
        {
            Write-Log "Admin Password successfully retrieved. We are NOT logging it here for security reasons. Saving it to a file which will be deleted instead."
            Out-file -FilePath $AdminPasswordPath -InputObject $AdminPassword
        }

        Else
        {
            Write-Log "Unable to retrieve the admin password from the installation log. Logging this problem to a file but continuing with the installation process."
            Out-file -FilePath $AdminPasswordPath -InputObject "NOPASS"
        }
	}

###############################################################################
# Create the Robots.txt file 
# This prevent crawlers from picking up this LabTech server's web components.
###############################################################################

Set-Content "${env:SystemDrive}\inetpub\wwwroot\robots.txt" "User-agent: *`r`nDisallow: /";

set-ItemProperty -PATH "HKLM:\SYSTEM\CurrentControlSet\services\Tcpip\Parameters" -NAME EnableTCPChimney -Value 0
set-ItemProperty -PATH "HKLM:\SYSTEM\CurrentControlSet\services\Tcpip\Parameters" -NAME MaxUserPort -Value 65534
set-ItemProperty -PATH "HKLM:\SYSTEM\CurrentControlSet\services\Tcpip\Parameters" -NAME TcpTimedWaitDelay -Value 30
set-ItemProperty -PATH "HKLM:\SYSTEM\CurrentControlSet\services\Tcpip\Parameters" -NAME ReservedPorts -Value '3306-3306,8000-9024,40000-40100'


<#######################################################################################>
#Install SQLYog

$DownloadObj = new-object System.Net.WebClient;
$DownloadObj.DownloadFile($SqlYogDownload, $SqlYogSavePath);

$InstallYog = New-Object System.Diagnostics.ProcessStartInfo("$SqlYogSavePath","/S");
$InstallYog.UseShellExecute = $False 
StartProcWithWait $InstallYog;

End-Script -Result "LabTech Installation Process Completed!";