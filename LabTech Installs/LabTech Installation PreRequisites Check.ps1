function Test-Key([string]$path, [string]$key)
{
    if(!(Test-Path $path)) { return $false }
    if ((Get-ItemProperty $path).$key -eq $null) { return $false }
    return $true 
}

#$ErrorActionPreference
[INT]$MissingPrereqs = 0
[Array]$MissingPrereqsDetails = @()
[String]$MySQLPass = '@MySQLPass@'
$ResultsPath = "$Env:Windir\temp\LTPrereqsResults.txt"
Remove-Item $ResultsPath -force | Out-Null

#Check for .NET 3.5
#########################################################################

$Net35Path = "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v3.5\"
$Net35Key = "Install"
$Net35Result = Test-key -path $Net35Path -key $Net35Key

If($Net35Result -ne $True)
{
    $MissingPrereqsDetails += ".NET 3.5 is not installed!"
    $MissingPrereqs++
}

#Check for .NET 4.5.2
#########################################################################
$Net452Path = "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full\"
$Net452Key = "Release"
$Net452Result = Get-ItemProperty "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full\"

If([INT]$Net452Result.release -lt 379893)
{
    $MissingPrereqsDetails += ".NET 4.5.2 is not installed!"
    $MissingPrereqs++
}

#Check for MySQL Compatible Database
#########################################################################

$LabMySQLService = Get-WmiObject win32_service | ?{$_.PathName -like '*MySQLD*'} | select Name, DisplayName, State, PathName
$MySQLDir = (($LabMySQLService.PathName -split '" ')[0] -replace '"','') -replace 'mysqld.exe',''

If(!$LabMySQLService)
{
    $MissingPrereqsDetails += "No Service named LabMySQL Exists!"
    $MissingPrereqs++  
}

set-location $MySQLDir;

$MySQLVersion = .\mysql.exe --user=root --password=$MySQLPass -e "
SELECT @@version;
;" --batch -N

If(!$MySQLVersion)
{
    Add-Content -Path $ResultsPath -Value "Bad MySQL Credentials!"
    exit
}

[Array]$TempMySQLVersion = $MySQLVersion -split '-'
[Int]$TempMySQLVersion = [string]$TempMySQLVersion[0].replace(".","")

If($MySQLVersion -like '*MariaDB*')
{
    If([Int]$TempMySQLVersion -lt 10019)
    {
        $MissingPrereqsDetails += "Unsupported version of MariaDB Detected!"
        $MissingPrereqs++
    }
}

Else
{
    
    IF([Int]$TempMySQLVersion -lt 5600)
    {
        $MissingPrereqsDetails += "Unsupported version of MySQL Detected!"
        $MissingPrereqs++ 
    }
    
    IF([Int]$TempMySQLVersion -gt 5700)
    {
        $MissingPrereqsDetails += "Unsupported version of MySQL Detected!"
        $MissingPrereqs++ 
    }
}

#Process Results
#########################################################################

If ([INT]$MissingPrereqs -eq 0)
{
    Add-Content -Path $ResultsPath -Value "No missing prerequisites"
}

Else
{
    Foreach($Item in $MissingPrereqsDetails)
    {
        Add-Content -Path $ResultsPath -Value "$Item `n"
    }
}