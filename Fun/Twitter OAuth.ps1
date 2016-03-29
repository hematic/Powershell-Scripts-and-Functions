##########################################################################
#Function Declarations

function Get-OAuth {
     <#
          .SYNOPSIS
           This function creates the authorization string needed to send a POST or GET message to the Twitter API

          .PARAMETER AuthorizationParams
           This hashtable should the following key value pairs
           HttpEndPoint - the twitter resource url [Can be found here: https://dev.twitter.com/rest/public]
           RESTVerb - Either 'GET' or 'POST' depending on the action
           Params - A hashtable containing the rest parameters (key value pairs) associated that method
           OAuthSettings - A hashtable that must contain only the following keys and their values (Generate here: https://dev.twitter.com/oauth)
                       ApiKey 
                       ApiSecret 
		               AccessToken
	                   AccessTokenSecret
          .LINK
           This function evolved from code found in Adam Betram's Get-OAuthAuthorization function in his MyTwitter module.
           The MyTwitter module can be found here: https://gallery.technet.microsoft.com/scriptcenter/Tweet-and-send-Twitter-DMs-8c2d6f0a
           Adam Betram's blogpost here: http://www.adamtheautomator.com/twitter-powershell/ provides a detailed explanation
           about how to generate an access token needed to create the authorization string 

          .EXAMPLE
            $OAuth = @{'ApiKey' = 'yourapikey'; 'ApiSecret' = 'yourapisecretkey';'AccessToken' = 'yourapiaccesstoken';'AccessTokenSecret' = 'yourapitokensecret'}	
            $Parameters = @{'q'='rumi'}
            $AuthParams = @{}
            $AuthParams.Add('HttpEndPoint', 'https://api.twitter.com/1.1/search/tweets.json')
            $AuthParams.Add('RESTVerb', 'GET')
            $AuthParams.Add('Params', $Parameters)
            $AuthParams.Add('OAuthSettings', $OAuth)
            $AuthorizationString = Get-OAuth -AuthorizationParams $AuthParams

          
     #>
    [OutputType('System.Management.Automation.PSCustomObject')]
	 Param($AuthorizationParams)
     process{
     try {

    	    ## Generate a random 32-byte string. I'm using the current time (in seconds) and appending 5 chars to the end to get to 32 bytes
	        ## Base64 allows for an '=' but Twitter does not.  If this is found, replace it with some alphanumeric character
	        $OauthNonce = [System.Convert]::ToBase64String(([System.Text.Encoding]::ASCII.GetBytes("$([System.DateTime]::Now.Ticks.ToString())12345"))).Replace('=', 'g')
    	    ## Find the total seconds since 1/1/1970 (epoch time)
		    $EpochTimeNow = [System.DateTime]::UtcNow - [System.DateTime]::ParseExact("01/01/1970", "dd/MM/yyyy", $null)
		    $OauthTimestamp = [System.Convert]::ToInt64($EpochTimeNow.TotalSeconds).ToString();
        	## Build the signature
			$SignatureBase = "$([System.Uri]::EscapeDataString($AuthorizationParams.HttpEndPoint))&"
			$SignatureParams = @{
				'oauth_consumer_key' = $AuthorizationParams.OAuthSettings.ApiKey;
				'oauth_nonce' = $OauthNonce;
				'oauth_signature_method' = 'HMAC-SHA1';
				'oauth_timestamp' = $OauthTimestamp;
				'oauth_token' = $AuthorizationParams.OAuthSettings.AccessToken;
				'oauth_version' = '1.0';
			}
	        $AuthorizationParams.Params.Keys | % { $SignatureParams.Add($_ , [System.Net.WebUtility]::UrlEncode($AuthorizationParams.Params.Item($_)).Replace('+','%20'))}
        
		 
			## Create a string called $SignatureBase that joins all URL encoded 'Key=Value' elements with a &
			## Remove the URL encoded & at the end and prepend the necessary 'POST&' verb to the front
			$SignatureParams.GetEnumerator() | sort name | foreach { $SignatureBase += [System.Uri]::EscapeDataString("$($_.Key)=$($_.Value)&") }

            $SignatureBase = $SignatureBase.Substring(0,$SignatureBase.Length-1)
            $SignatureBase = $SignatureBase.Substring(0,$SignatureBase.Length-1)
            $SignatureBase = $SignatureBase.Substring(0,$SignatureBase.Length-1)
			$SignatureBase = $AuthorizationParams.RESTVerb+'&' + $SignatureBase
			
			## Create the hashed string from the base signature
			$SignatureKey = [System.Uri]::EscapeDataString($AuthorizationParams.OAuthSettings.ApiSecret) + "&" + [System.Uri]::EscapeDataString($AuthorizationParams.OAuthSettings.AccessTokenSecret);
			
			$hmacsha1 = new-object System.Security.Cryptography.HMACSHA1;
			$hmacsha1.Key = [System.Text.Encoding]::ASCII.GetBytes($SignatureKey);
			$OauthSignature = [System.Convert]::ToBase64String($hmacsha1.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($SignatureBase)));
			
			## Build the authorization headers using most of the signature headers elements.  This is joining all of the 'Key=Value' elements again
			## and only URL encoding the Values this time while including non-URL encoded double quotes around each value
			$AuthorizationParams = $SignatureParams
			$AuthorizationParams.Add('oauth_signature', $OauthSignature)
		
			
			$AuthorizationString = 'OAuth '
			$AuthorizationParams.GetEnumerator() | sort name | foreach { $AuthorizationString += $_.Key + '="' + [System.Uri]::EscapeDataString($_.Value) + '", ' }
			$AuthorizationString = $AuthorizationString.TrimEnd(', ')
            Write-Verbose "Using authorization string '$AuthorizationString'"			
			$AuthorizationString

        }
        catch {
			Write-Error $_.Exception.Message
		}

     }

}

function Invoke-TwitterRestMethod{
<#
          .SYNOPSIS
           This function sends a POST or GET message to the Twitter API and returns the JSON response. 

          .PARAMETER ResourceURL
           The desired twitter resource url [REST APIs can be found here: https://dev.twitter.com/rest/public]
           
          .PARAMETER RestVerb
           Either 'GET' or 'POST' depending on the resource URL

           .PARAMETER  Parameters
           A hashtable containing the rest parameters (key value pairs) associated that resource url. Pass empty hash if no paramters needed.

           .PARAMETER OAuthSettings 
           A hashtable that must contain only the following keys and their values (Generate here: https://dev.twitter.com/oauth)
                       ApiKey 
                       ApiSecret 
		               AccessToken
	                   AccessTokenSecret

           .EXAMPLE
            $OAuth = @{'ApiKey' = 'yourapikey'; 'ApiSecret' = 'yourapisecretkey';'AccessToken' = 'yourapiaccesstoken';'AccessTokenSecret' = 'yourapitokensecret'}
            Invoke-TwitterRestMethod -ResourceURL 'https://api.twitter.com/1.1/statuses/mentions_timeline.json' -RestVerb 'GET' -Parameters @{} -OAuthSettings $OAuth 
            Invoke-TwitterRestMethod -ResourceURL 'https://api.twitter.com/1.1/statuses/user_timeline.json' -RestVerb 'GET' -Parameters @{'count' = '1'} -OAuthSettings $OAuth 
            Invoke-TwitterRestMethod -ResourceURL 'https://api.twitter.com/1.1/statuses/home_timeline.json' -RestVerb 'GET' -Parameters @{'count' = '1'} -OAuthSettings $OAuth 
            Invoke-TwitterRestMethod -ResourceURL 'https://api.twitter.com/1.1/statuses/retweets_of_me.json' -RestVerb 'GET' -Parameters @{} -OAuthSettings $OAuth 
            Invoke-TwitterRestMethod -ResourceURL 'https://api.twitter.com/1.1/search/tweets.json' -RestVerb 'GET' -Parameters @{'q'='powershell';'count' = '1'}} -OAuthSettings $OAuth
            Invoke-TwitterRestMethod -ResourceURL 'https://api.twitter.com/1.1/account/settings.json' -RestVerb 'POST' -Parameters @{'lang'='tr'} -OAuthSettings $OAuth 
            Invoke-TwitterRestMethod -ResourceURL 'https://api.twitter.com/1.1/statuses/retweets/509457288717819904.json' -RestVerb 'GET' -Parameters @{} -OAuthSettings $OAuth
            Invoke-TwitterRestMethod -ResourceURL 'https://api.twitter.com/1.1/statuses/show.json' -RestVerb 'GET' -Parameters @{'id'='123'} -OAuthSettings $OAuth 
            Invoke-TwitterRestMethod -ResourceURL 'https://api.twitter.com/1.1/statuses/destroy/240854986559455234.json' -RestVerb 'GET' -Parameters @{} -OAuthSettings $OAuth 
            Invoke-TwitterRestMethod -ResourceURL 'https://api.twitter.com/1.1/statuses/update.json' -RestVerb 'POST' -Parameters @{'status'='@FollowBot'} -OAuthSettings $OAuth 
            Invoke-TwitterRestMethod -ResourceURL 'https://api.twitter.com/1.1/direct_messages.json' -RestVerb 'GET' -Parameters @{} -OAuthSettings $OAuth 
            Invoke-TwitterRestMethod -ResourceURL 'https://api.twitter.com/1.1/direct_messages/destroy.json' -RestVerb 'POST' -Parameters @{'id' = '559298305029844992'} -OAuthSettings $OAuth 
            Invoke-TwitterRestMethod -ResourceURL 'https://api.twitter.com/1.1/direct_messages/new.json' -RestVerb 'POST' -Parameters @{'text' = 'hello, there'; 'screen_name' = 'ruminaterumi' } -OAuthSettings $OAuth 
            $mediaId = Invoke-TwitterMEdiaUpload -MediaFilePath 'C:\Books\pic.png' -ResourceURL 'https://upload.twitter.com/1.1/media/upload.json' -OAuthSettings $OAuth 
            Invoke-TwitterRestMethod -ResourceURL 'https://api.twitter.com/1.1/statuses/update.json' -RestVerb 'POST' -Parameters @{'status'='FollowBot'; 'media_ids' = $mediaId } -OAuthSettings $OAuth 

     #>
         [CmdletBinding()]
	     [OutputType('System.Management.Automation.PSCustomObject')]
         Param(
                [Parameter(Mandatory)]
                [string]$ResourceURL,
                [Parameter(Mandatory)]
                [string]$RestVerb,
                [Parameter(Mandatory)]
                $Parameters,
                [Parameter(Mandatory)]
                $OAuthSettings

                )

          process{
              try{

                    $AuthParams = @{}
                    $AuthParams.Add('HttpEndPoint', $ResourceURL)
                    $AuthParams.Add('RESTVerb', $RestVerb)
                    $AuthParams.Add('Params', $Parameters)
                    $AuthParams.Add('OAuthSettings', $OAuthSettings)
                    $AuthorizationString = Get-OAuth -AuthorizationParams $AuthParams                 
                    $HTTPEndpoint= $ResourceURL
                    if($Parameters.Count -gt 0)
                    {
                        $HTTPEndpoint = $HTTPEndpoint + '?'
                        $Parameters.Keys | % { $HTTPEndpoint = $HTTPEndpoint + $_  +'='+ [System.Net.WebUtility]::UrlEncode($Parameters.Item($_)).Replace('+','%20') + '&'}
                        $HTTPEndpoint = $HTTPEndpoint.Substring(0,$HTTPEndpoint.Length-1)
  
                    }
                    Invoke-RestMethod -URI $HTTPEndpoint -Method $RestVerb -Headers @{ 'Authorization' = $AuthorizationString } -ContentType "application/x-www-form-urlencoded" 
                  }
                  catch{
                    Write-Error $_.Exception.Message
                  }
            }
}

function Invoke-ReadFromTwitterStream{
<#
          .SYNOPSIS
           This function can be used to download info from the Twitter Streaming APIs and record the json ouptut in a text file. 

          .PARAMETER ResourceURL
           The desired twitter resource url [Streaming APIs can be found here: https://dev.twitter.com/streaming/overview]
           
          .PARAMETER RestVerb
           Either 'GET' or 'POST' depending on the resource URL

           .PARAMETER  Parameters
           A hashtable containing the rest parameters (key value pairs) associated that resource url. Pass empty hash if no paramters needed.

           .PARAMETER OAuthSettings 
           A hashtable that must contain only the following keys and their values (Generate here: https://dev.twitter.com/oauth)
                       ApiKey 
                       ApiSecret 
		               AccessToken
	                   AccessTokenSecret

           .PARAMETER  MinsToCollectStream
           The number of minutes you want to attempt to stream content. Use -1 to run infinte loop. 

           .PARAMETER  OutFilePath
           The location of the out file text. Will create file if dne yet.

           .EXAMPLE 
            $OAuth = @{'ApiKey' = 'yourapikey'; 'ApiSecret' = 'yourapisecretkey';'AccessToken' = 'yourapiaccesstoken';'AccessTokenSecret' = 'yourapitokensecret'}
            Invoke-ReadFromTwitterStream -OAuthSettings $o -OutFilePath 'C:\books\foo.txt' -ResourceURL 'https://stream.twitter.com/1.1/statuses/filter.json' -RestVerb 'POST' -Parameters @{'track' = 'foo'} -MinsToCollectStream 1

           .LINK
           This function evolved from the following blog posts http://thoai-nguyen.blogspot.com.tr/2012/03/consume-twitter-stream-oauth-net.html, https://code.google.com/p/pstwitterstream/
#>
           [CmdletBinding()]
           Param(
                [Parameter(Mandatory)]
                $OAuthSettings,
                [Parameter(Mandatory)] 
                [String] $OutFilePath,
                [Parameter(Mandatory)] 
                [string]$ResourceURL,
                [Parameter(Mandatory)] 
                [string]$RestVerb,
                [Parameter(Mandatory)] 
                $Parameters,
                [Parameter(Mandatory)] 
                $MinsToCollectStream
                )

                process{
                $Ti = Get-Date  
                while($true)
                {
                  $NewD = Get-Date
                  if(($MinsToCollectStream -ne -1) -and (($NewD-$Ti).Minutes -gt $MinsToCollectStream))
                  { return "Finished"}
     
                  try
                  {
                    $AuthParams = @{}
                    $AuthParams.Add('HttpEndPoint', $ResourceURL)
                    $AuthParams.Add('RESTVerb', $RestVerb)
                    $AuthParams.Add('Params', $Parameters)
                    $AuthParams.Add('OAuthSettings', $OAuthSettings)
                    $AuthorizationString = Get-OAuth -AuthorizationParams $AuthParams

                    [System.Net.HttpWebRequest]$Request = [System.Net.WebRequest]::Create($ResourceURL)
                    $Request.Timeout = [System.Threading.Timeout]::Infinite
                    $Request.Method = $RestVerb
                    $Request.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip, [System.Net.DecompressionMethods]::Deflate 
                    $Request.Headers.Add('Authorization', $AuthorizationString)
                    $Request.Headers.Add('Accept-Encoding', 'deflate,gzip')
                    $filter = $Null
                    if($Parameters.Count -gt 0)
                    {
                        $Parameters.Keys | % { $filter = $filter + $_  +'='+ [System.Net.WebUtility]::UrlEncode($Parameters.Item($_)).Replace('+','%20') + '&'}
                        $filter = $filter.Substring(0, $filter.Length-1)
                        $POSTData = [System.Text.Encoding]::UTF8.GetBytes($filter)
                        $Request.ContentType = "application/x-www-form-urlencoded"
                        $Request.ContentLength = $POSTData.Length
                        $RequestStream = $Request.GetRequestStream()
                        $RequestStream.Write($POSTData, 0, $POSTData.Length)
                        $RequestStream.Close()
                    }
                 
                    $Response =  [System.Net.HttpWebResponse]$Request.GetResponse()
                    [System.IO.StreamReader]$ResponseStream = $Response.GetResponseStream()
                    
                    while ($true) 
                    {
                            $NewDt = Get-Date
                            if(($MinsToCollectStream -ne -1) -and (($NewDt-$Ti).Minutes -gt $MinsToCollectStream))
                            { return "Finished"}

                            $Line = $ResponseStream.ReadLine()
                            if($Line -eq '') 
                            { continue }
                            Add-Content $OutFilePath $Line
                            $PowerShellRepresentation = $Line | ConvertFrom-Json
                            $PowerShellRepresentation
                            If ($ResponseStream.EndOfStream) { Throw "Stream closed." }                  
                    }
                 }
                 catch{
                    Write-Error $_.Exception.Message
                }
                }
              }
}

Function Get-TwitterStream{
##
#.SYNOPSIS
# Get the Twitter time line of a user or a set of users.
#.DESCRIPTION
# Get the Twitter time line of a user or a set of users using the REST API.
# 
# The api allows us to retrieve the timelines of users when queried:
# http://api.twitter.com/1/statuses/user_timeline.xml?screen_name=UserName
# 
# You can retrieve multiple pages but, for this script we only retrieve top
# twenty tweets in a single page.
# 
# Since, the function expects a string argument when passing multiple parameters
# you need to pass them as: "handle#1, handle#2, ...handle#n"
# 
#.PARAMETER TwitterHandle
# A user's "Twitter handle" is the username they have selected.
# 
#.EXAMPLE
# Get the Twitter timeline of a single user:
# Get-TwitterStream -TwitterHandle SqlChow
# 
#.EXAMPLE
# When getting the Twitter timeline of multiple users 
# Get-TwitterStream -TwitterHandle "SqlChow, BrentO, PaulRandal"
# 
#.NOTES
# ###########################################################################
# #
# # NAME: Get-TwitterStream
# #
# # AUTHOR: @SqlChow
# #
# # COMMENT: Get-TwitterStream -TwitterHandle "SqlChow, BrentO,realpreityzinta"
# #
# # COPYRIGHT: © 2011 SqlChow. Code provided as is. User discretion advised.
# #
# # VERSION HISTORY:
# # 1.0 12/25/2011 - Initial release
# # 1.1 12/27/2011 - Completed todo list
# ########################################################################### 

param(
	[parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[String]
	$TwitterHandle
	) 

	Begin
	{
	$script:flip = $true
	$origColor = $Host.ui.rawui.ForegroundColor 

	Function Color-Tweet ($text){
	if($flip){$Host.UI.RawUI.ForegroundColor="cyan"; "{0,4}" -f "$text `n";$script:flip=$false}
	else{$Host.ui.rawui.foregroundcolor = "yellow";"$text `n";$script:flip=$true}
	} 

	$colorTweet = @{
	label = "Tweets:" ;
	Expression = { Color-Tweet $_.text };
	Width = 47; Alignment="Left"
	} 

	} 

	Process
	{
	try
	{
	$objWebClient = New-Object System.Net.WebClient 

	$arryOfHandles = $TwitterHandle.Split(",")
	foreach($handle in $arryOfHandles)
	{
	Write-Host "`@"$handle.Trim()" tweets:" -Fore Green -Back Black
	$strTwitURL = "http://api.twitter.com/1/statuses/user_timeline.xml?screen_name="+ $handle.Trim() + "&count=20&page=1"
	Write-Debug $strTwitUrl
	$strContentString = $objWebClient.DownloadString($strTwitURL)
	$xmlContents = [xml]$strContentString 

	#$xmlContents.statuses.ChildNodes|ForEach-Object{ if ($flip) { $flip=$false;$Host.ui.rawui.ForegroundColor = "cyan"; "{0,4}" -f $_.text;"`n";} else { $flip=$true;$Host.ui.rawui.foregroundcolor = "yellow";"{0,4}" -f $_.text;"`n";} }
	#$xmlContents.statuses.ChildNodes|Add-Member NoteProperty Tweeter $handle -PassThru|Format-Table $colorTweet, Tweeter -Wrap
	$xmlContents.statuses.ChildNodes| Format-Table $colorTweet -Wrap -HideTableHeaders
	$Host.ui.rawui.ForegroundColor = $origColor
	}
	}
	catch
	{
	Write-Error "Unable to get stream data"
	}
	finally
	{
	$objWebClient.Dispose()
	$Host.ui.rawui.ForegroundColor = $origColor
	}
	}
}

##########################################################################
#Variable Declarations

$ErrorActionPreference = 'silentlycontinue'
[String]$GmailUsername = ''
$GmailPassword = '' | ConvertTo-SecureString -asPlainText -Force
$Creds = New-Object System.Management.Automation.PSCredential($GmailUsername,$GmailPassword)
[String]$oauth_consumer_key = "";
[String]$oauth_consumer_secret = "";
[String]$oauth_token = "";
[String]$oauth_token_secret = "";

$OAuth = @{ 'ApiKey' = $oauth_consumer_key; 
            'ApiSecret' = $oauth_consumer_secret; 
            'AccessToken' = $oauth_token; 
            'AccessTokenSecret' = $oauth_token_secret;}

[String]$TwitterUsr = "";
[String]$AttachmentPath = "$($env:windir)\temp\TweetedPic.jpg"
[String]$LastTweetPath = "$($env:windir)\temp\LastTweet.txt"

##########################################################################
# Fetches the tweets from our twitter account

$TwitterParams = @{

ResourceURL = 'https://api.twitter.com/1.1/statuses/user_timeline.json'; 
RestVerb = 'GET'; 
Parameters =  @{ 'screen_name' = $TwitterUsr; 'count' = 20; } 
OAuthSettings =  $OAuth;
}
$Tweets = Invoke-TwitterRestMethod @TwitterParams

##########################################################################
# Fetches the last tweet we scanned

[Long]$LastTweet = Get-Content $LastTweetPath
If([Long]$LastTweet -eq ''){$LastTweet = 0} 

##########################################################################
# Exracts the images and emails them.

Foreach($Tweet in $Tweets)
{
    If (($tweet.extended_entities.media.media_url) -and ($Tweet.id -gt $LastTweet))
    {
        Write-Output @"

        ##########################################################################
        Tweet ID:   $($Tweet.id)
        Tweet Text: $($Tweet.text)
        Result:     We've Got Pictures!
        ##########################################################################

"@

        Foreach ($Link in $tweet.extended_entities.media.media_url)
        {
            $GmailParams = @{
            BodyAsHtml = $true
            SmtpServer = 'smtp.gmail.com'
            Port = 587
            UseSsl = $true
            Credential  = $Creds
            From = 'Art <artlabtecher@gmail.com>'
            To = 'Phillip <pmarshall@labtechsoftware.com>'
            Subject = "New Pictures from Art's Twitter"
            Body = @"
            Art posted a new picture! Check it out!
            <BR>
            <img src="$Link">

"@
        }
            Send-MailMessage @GmailParams
        }
    }

    Else
    {
        Write-Output @"

        ##########################################################################
        Tweet ID:   $($Tweet.id)
        Tweet Text: $($Tweet.text)
        Result:     Not Using this Tweet
        ##########################################################################

"@

    }
}

##########################################################################
# Calculates highest parsed Tweet Id and stores it in our file.

<# In this case measure object, while simpler, does NOT work. Measure object 
   returns a type of variable called a "Double" (More info here):
   https://en.wikipedia.org/wiki/Double-precision_floating-point_format

   This rounds the numbers to where they are slightly off.

   Example: 623577461557755906 in an Int64 or a [LONG] type. If you convert 
   it to a double it becomes 6.23577461557756E+17. if you convert that
   back to a long it is 623577461557755904. This obviously is not a good 
   thing so we couldnt use measure object here.
#>

#$MaxTweetID = ($Tweets.id | measure -Maximum).Maximum

[Long]$MaxTweetID = 0

ForEach($Number in $Tweets.Id)
{
    If($Number -gt $MaxTweetID)
    {
        [Long]$MaxTweetID = $Number;
    }
}


Set-Content -Path $LastTweetPath -Value $MaxTweetID