# (Not an entry just a useful script to build upon)
# Originally From https://gallery.technet.microsoft.com/Send-Tweets-via-a-72b97964
# Modified to stop it being a workflow (I couldn't get it to work as a workflow)
# And made it validate the four different twitter passwords you need (seriously twitter... chill out)

function send-text($message) {    
    [Reflection.Assembly]::LoadWithPartialName("System.Security")  
    [Reflection.Assembly]::LoadWithPartialName("System.Net")

<#
    After visiting https://apps.twitter.com/app/new you can generate these four things...
    Perhaps set them in your profile
    
    $env:twitter_oauth_consumer_key = "YOUR_API_KEY"
    $env:twitter_oauth_consumer_secret = "YOUR_API_SECRET"
    $env:twitter_oauth_token = "YOUR_ACCESS_TOKEN"
    $env:twitter_oauth_token_secret = "YOUR_ACCESS_TOKEN_SECRET"
#>    

    If(!$env:twitter_oauth_consumer_key)
            {
                Throw "You need to register and get an oauth_consumer_key from twitter and save it as environment variable `$env:twitter_oauth_consumer_key = `"YOUR_API_KEY`" `nFollow this link and get the API Key - https://apps.twitter.com/app/new `n`n "
            }
    If(!$env:twitter_oauth_consumer_secret)
            {
                Throw "You need to register and get an oauth_consumer_secret from twitter and save it as environment variable `$env:twitter_oauth_consumer_secret = `"YOUR_API_SECRET`" `nFollow this link and get the API secret - https://apps.twitter.com/app/new `n`n "
            }
    If(!$env:twitter_oauth_token)
            {
                Throw "You need to register and get an oauth_token from twitter and save it as environment variable `$env:twitter_oauth_token = `"YOUR_ACCESS_TOKEN`" `nFollow this link and get the access token - https://apps.twitter.com/app/new `n`n "
            }
    If(!$env:twitter_oauth_token_secret)
            {
                Throw "You need to register and get an oauth_token_secret from twitter and save it as environment variable `$env:twitter_oauth_token_secret = `"YOUR_ACCESS_TOKEN_SECRET`" `nFollow this link and get the access token secret - https://apps.twitter.com/app/new `n`n "
            }

    $status = [System.Uri]::EscapeDataString($Using:Message);  
    
    $oauth_consumer_key = $env:twitter_oauth_consumer_key; #"<YourAPIKey>";  
    $oauth_consumer_secret = $env:twitter_oauth_consumer_secret; #"<YourAPISecret>";  
    $oauth_token = $env:twitter_oauth_token; #"<YourAccessToken>";  
    $oauth_token_secret = $env:twitter_oauth_token_secret; #"<YourAccessTokenSecret>";  
    
    $oauth_nonce = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes([System.DateTime]::Now.Ticks.ToString()));  
    $ts = [System.DateTime]::UtcNow - [System.DateTime]::ParseExact("01/01/1970", "dd/MM/yyyy", $null).ToUniversalTime();  
    $oauth_timestamp = [System.Convert]::ToInt64($ts.TotalSeconds).ToString();  

    $signature = "POST&";  
    $signature += [System.Uri]::EscapeDataString("https://api.twitter.com/1.1/statuses/update.json") + "&";  
    $signature += [System.Uri]::EscapeDataString("oauth_consumer_key=" + $oauth_consumer_key + "&");  
    $signature += [System.Uri]::EscapeDataString("oauth_nonce=" + $oauth_nonce + "&");   
    $signature += [System.Uri]::EscapeDataString("oauth_signature_method=HMAC-SHA1&");  
    $signature += [System.Uri]::EscapeDataString("oauth_timestamp=" + $oauth_timestamp + "&");  
    $signature += [System.Uri]::EscapeDataString("oauth_token=" + $oauth_token + "&");  
    $signature += [System.Uri]::EscapeDataString("oauth_version=1.0&");  
    $signature += [System.Uri]::EscapeDataString("status=" + $status);  

    $signature_key = [System.Uri]::EscapeDataString($oauth_consumer_secret) + "&" + [System.Uri]::EscapeDataString($oauth_token_secret);  

    $hmacsha1 = new-object System.Security.Cryptography.HMACSHA1;  
    $hmacsha1.Key = [System.Text.Encoding]::ASCII.GetBytes($signature_key);  
    $oauth_signature = [System.Convert]::ToBase64String($hmacsha1.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($signature)));  

    $oauth_authorization = 'OAuth ';  
    $oauth_authorization += 'oauth_consumer_key="' + [System.Uri]::EscapeDataString($oauth_consumer_key) + '",';  
    $oauth_authorization += 'oauth_nonce="' + [System.Uri]::EscapeDataString($oauth_nonce) + '",';  
    $oauth_authorization += 'oauth_signature="' + [System.Uri]::EscapeDataString($oauth_signature) + '",';  
    $oauth_authorization += 'oauth_signature_method="HMAC-SHA1",'  
    $oauth_authorization += 'oauth_timestamp="' + [System.Uri]::EscapeDataString($oauth_timestamp) + '",'  
    $oauth_authorization += 'oauth_token="' + [System.Uri]::EscapeDataString($oauth_token) + '",';  
    $oauth_authorization += 'oauth_version="1.0"';  

    $post_body = [System.Text.Encoding]::ASCII.GetBytes("status=" + $status);   
    [System.Net.HttpWebRequest] $request = [System.Net.WebRequest]::Create("https://api.twitter.com/1.1/statuses/update.json");  
    $request.Method = "POST";  
    $request.Headers.Add("Authorization", $oauth_authorization);  
    $request.ContentType = "application/x-www-form-urlencoded";  
    $body = $request.GetRequestStream();  
    $body.write($post_body, 0, $post_body.length);  
    $body.flush();  
    $body.close();  
    $response = $request.GetResponse();
}
