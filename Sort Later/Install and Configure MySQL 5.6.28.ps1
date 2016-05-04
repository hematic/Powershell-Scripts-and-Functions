<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2015 v4.2.99
	 Created on:   	1/27/2016 9:50 AM
	 Created by:   	Marcus Bastian
	 Filename:     	Install and Configure MySQL 5.6.28.ps1
	===========================================================================
	.DESCRIPTION

		Just a quick PowerShell script that:

		+ Downloads and installs MySQL 5.6, which runs under the Network Service account
		+ Optimizes the MY.INI so that the LabTech installation proceeds without errors
#>

Function Generate-RandomPassword ($length = 14)
{
    $digits = 48..57
    $letters = 65..90 + 97..122

    # Thanks to
    # https://blogs.technet.com/b/heyscriptingguy/archive/2012/01/07/use-pow
    $password = get-random -count $length `
        -input ($punc + $digits + $letters) |
            % -begin { $aa = $null } `
            -process {$aa += [char]$_} `
            -end {$aa}

    return $password
}

function Download-FileFromURL ($url, $saveAs)
{
	$webclient = New-Object System.net.WebClient;
	$webclient.DownloadFile($url, $saveAs);
	
	if (Test-Path $saveAs)
	{
		return $true;
	}
	else
	{
		return $false;
	}
	
	
}

Function Start-ProcessWithWait 
{

	$InstallApplication = $args[0];
    $InstallProc = New-Object System.Diagnostics.Process;
	$InstallProc.StartInfo = $InstallApplication;
	$InstallProc.Start();
	$InstallProc.WaitForExit();
	
}



#####################################################################
# 			CHECK/INSTALL .NET FRAMEWORK 4.0    - 2008 R2 Only
#####################################################################

$Url = "http://automationfiles.hostedrmm.com/third_party_apps/.NET/dotnetfx40_full_x86_x64.exe"
$SaveAs = "$env:windir\temp\dotnetfx40_full_x86_x64.exe"
$OS = (Get-WmiObject Win32_OperatingSystem).caption;

if($OS -like '*2008*') # 2008 R2 ONLY
{
		Download-FileFromURL -url $url -saveAs $saveAs 
	Write-Output "Beginning installation of .NET Framework 4.0 extended .............";

	$InstallDotNet4 = New-Object System.Diagnostics.ProcessStartInfo("$saveAs", "/q /norestart");
	Start-ProcessWithWait $InstallDotNet4;

	if(!(gwmi -cl win32_product | where-object { $_.Name -like '*Microsoft .NET Framework 4.0*' })) 
	{ 
		Write-Output "FAILED to install .NET Framework 4.0 ..............................";
		return;
	}
	ELSE
	{
		Write-output "SUCCESSFULLY installed .NET Framework 4.0 .........................";
	}
}

#####################################################################
# 			EXECUTE MYSQL 5.6 CONFIGURATION TOOLS MSI SILENTLY
#####################################################################
# Note:  	 This MSI doesn't actually create the MySQL instance
#####################################################################

# Execute the MySQL installer MSI... This simply installs the MySQLInstallerConsole into program files x86
$Url = "http://automationfiles.hostedrmm.com/third_party_apps/mysql_x64/mysql-installer-community-5.6.28.0.msi"
$MSIPath = "$env:windir\temp\mysql-installer-community-5.6.28.0.msi"

Download-FileFromURL -url $url -saveAs $MSIPath

if (Test-Path $MSIPath)
{
	Write-Output "Successfully downloaded MySQL community installer for 5.6.28."
}
else
{
	Write-Error "Failed to download the MySQL Community installer for 5.6.28. Please make sure MSIs are not blocked by a firewall."
	return;
}

# Download was successful. Launch the MSIExec process.
$InstallMySQL = New-Object System.Diagnostics.ProcessStartInfo("$ENV:windir\system32\msiexec.exe", "/i `"$MSIPath`" /qn");
$InstallMySQL.UseShellExecute = $False;

Start-ProcessWithWait $InstallMySQL | out-null;

#####################################################################
# 			CREATE MYSQL INSTANCE USING INSTANCE CONFIG EXE
#####################################################################

$Location = Read-Host "Please enter the path to the directory you would like your database data files to be stored. It is recommended that you not store these files on your OS drive ($($env:SystemDrive))"

$IsValidDirectory = Test-Path $Location;



# Create MySQL instance "LabMySQL"
$NewPassword = Generate-RandomPassword;

Start-Process "${env:ProgramFiles(x86)}\MySQL\MySQL Installer for Windows\MySQLInstallerConsole.exe" `
			  -ArgumentList "community",
			  "install",
			  "server;5.6.28;x64:*:type=config;openfirewall=true;servicename=LabMySQL;binlog=false;serverid=3306;enable_tcpip=true;port=3306;rootpasswd=$NewPassword;installdir=`"C:\Program Files\MySQL\MySQL Server 5.6`";datadir=`"C:\ProgramData\MySQL\MySQL Server 5.6`"",
			  "-silent" -Wait -NoNewWindow;



# Verify that the service was created successfully
IF (Get-Service LabMySQL)
{
	Write-Host "Successfully installed MySQL 5.6 as `"LabMySQL`"!" -ForegroundColor Green;
	Write-Host "Pwd:   $NewPassword";
}
else
{
	Write-Error "Failed to install MySQL 5.6 as `"LabMySQL`"!";
	Pause;
}

