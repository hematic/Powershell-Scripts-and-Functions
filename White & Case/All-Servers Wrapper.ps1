$Servers = Get-ADComputer -LDAPFilter “(&(objectcategory=computer)(OperatingSystem=*server*))”

If(!$Servers)
{
    Write-Log "Gathering list of servers from Active Directory failed."
    exit  
}







Foreach($Server in $Servers)
{
    $ServerResult = Invoke-Command -ComputerName $($Server.name) -FilePath "C:\Users\marshph\Documents\Hematic_Github\Powershell-Scripts-and-Functions\White & Case" -Credential
}


