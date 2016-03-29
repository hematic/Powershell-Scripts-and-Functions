function Get-GeoIP 
{
    param
	(
    	[Parameter(Mandatory=$true)]
    	[System.Net.IPAddress]$ip
    )

    $webclient = New-Object System.Net.webclient
    $providerRoot = "http://freegeoip.net/xml/"
    [xml]$geoData = $webclient.downloadstring($providerRoot+$ip)
    Write-Output $geoData.response
}