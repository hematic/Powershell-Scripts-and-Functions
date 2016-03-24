function Execute-SOAPRequest
{
	Param
	(
	    [Xml]    $SOAPRequest, 
        [String] $URL 
	) 
 
        $soapWebRequest = [System.Net.WebRequest]::Create($URL) 
        $soapWebRequest.Headers.Add("SOAPAction","`"http//localhost/CommunityConnector/GetLatestMarketplaceVersionByLabTechVersion`"")

        $soapWebRequest.ContentType = "text/xml; charset=utf-8" 
        $soapWebRequest.Accept      = "text/xml" 
        $soapWebRequest.Method      = "POST" 
        
        $requestStream = $soapWebRequest.GetRequestStream() 
        $SOAPRequest.Save($requestStream) 
        $requestStream.Close() 
        
        $resp = $soapWebRequest.GetResponse() 
        $responseStream = $resp.GetResponseStream() 
        $soapReader = [System.IO.StreamReader]($responseStream) 
        $ReturnXml = [Xml] $soapReader.ReadToEnd() 
        $responseStream.Close() 
        
        return $ReturnXml 
}

$url = 'http://mp2014.hostedrmm.com/CommunityConnector/CommunityConnector.asmx'
$soap = [xml]@'
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" 
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
    xmlns:xsd="http://www.w3.org/2001/XMLSchema">
<soap:Body>
    <GetLatestMarketplaceVersionByLabTechVersion 
        xmlns="http//localhost/CommunityConnector/">
        <labTechVersion>105226</labTechVersion>
    </GetLatestMarketplaceVersionByLabTechVersion>
</soap:Body>
</soap:Envelope>
'@
$ret = Execute-SOAPRequest $soap $url

$major = $ret.GetElementsByTagName("Major").'#text'
$minor = $ret.GetElementsByTagName("Minor").'#text'
$build = $ret.GetElementsByTagName("Build").'#text'
$revision = $ret.GetElementsByTagName("Revision").'#text'
$ReturnURL = $ret.GetElementsByTagName("URL").'#text'

$ltcheck= test-path "C:\Program Files (x86)\LabTech Client\ltmarketplace.exe"

if($ltcheck -ne 'True') 
{
    Return $ReturnURL
}

else
{
    $FullBuild = $major + '.' + $minor + '.' + $build + '.' + $revision
    $ExeBuild = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("C:\Program Files (x86)\LabTech Client\ltmarketplace.exe").ProductVersion


    If ($FullBuild -eq $ExeBuild -or $ExeBuild -gt $fullbuild)
    {
        Return '1'
    }

    Else
    {
        Return $ReturnURL 
    }
}