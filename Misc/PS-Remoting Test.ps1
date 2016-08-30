$Credential = Get-credential
$ListofComputers = Get-ADComputer -Filter { OperatingSystem -like "*Server*" } -Property name, operatingsystem
$Filtered = $ListofComputers | where-object {$_.operatingsystem -notlike '*2003*' -and $_.operatingsystem -notlike '*2000*' -and $_.name -notlike '*TCRA*' -and $_.name -notlike '*RAAP*' -and $_.name -notlike '*DC*'}

$NonPingableServers = @()
$PingableServers = @()

$Filtered | Start-RSjob -Name { "$($_.name)" } -ScriptBlock {
	Param ($Filtered)
	
	invoke-command -computername $Filtered.name -credential $Using:Credential -ScriptBlock {
        
        Return hostname
	}
}


$Failures = Get-RSJob | where-object {$_.HasErrors -eq $True} | select -ExpandProperty name

$FilteredFailures = $Failures | where-object {$_ -notlike '*TCRA*' -and $_ -notlike '*RAAP*' -and $_ -notlike '*DC*'}
$FailedOSList = @()

$FailedServerObj = @()

Foreach($Failure in $FilteredFailures) {

    $Server = Get-ADComputer -Filter { name -eq $Failure } -Property name, operatingsystem

    If(!$Server)
    {
        Write-output "Couldn't find $Failure in AD"
    }
    $FailedServerObj += $Server

 }

 $FinalFilterList = $FailedServerObj | where-object {$_.operatingsystem -notlike '*ubuntu*' -and $_.operatingsystem -notlike '*Windows 7*' -and $_.operatingsystem -notlike '*Windows 10*' -and $_.operatingsystem -notlike '*Windows XP*' -and $_.operatingsystem -notlike '*Windows server 2003*' -and $_.name -notlike '*APPR*'} | select -Property name, operatingsystem

 Foreach($Server in  $FinalFilterList)
{
    If(Test-Connection -ComputerName $Server.name -Count 1 -Quiet)
    {
        $PingableServers += $Server
    }

    Else
    {
        $NonPingableServers += $Server
    }

}