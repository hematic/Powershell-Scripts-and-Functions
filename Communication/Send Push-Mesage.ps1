Function Send-PushMessage {
[CmdletBinding(DefaultParameterSetName='Message')]

#Usage Examples:
#Send an Address: Send-PushMessage -Type Address -PlaceName "Bob's House" -PlaceAddress "5555 Ashburn Lake Dr Tampa Fl 33610"
#Send a Message:  Send-PushMessage -Type Message -Title "This is a test" -msg "The message goes here!"
#Send a Link:     Send-PushMessage -Type Link -Title "This Is a Link" -msg "Link Text Goes here!" -url "www.pushbullet.com"
#Send a File:     Send-PushMessage -Type File -Filename "Webroot.zip" -FileType "Anything" -url "http://download.webroot.com/Webroot-Deploy-Solution.zip"
#Send a list:     **This Function is not yet working** Send-PushMessage -Type List -title "My List" -items "Item 1, Item 2, Item 3"
#Upload a file:   **This Function is not yet working**
 
param(
        [Parameter(Mandatory=$false,ParameterSetName="File")]$FileName,
        [Parameter(Mandatory=$true, ParameterSetName="File")]$FileType,
        [Parameter(Mandatory=$true, ParameterSetName="File")]

        [Parameter(Mandatory=$true, ParameterSetName="Link")]$url,
         
        [Parameter(Mandatory=$false,ParameterSetName="Address")]$PlaceName,
        [Parameter(Mandatory=$true, ParameterSetName="Address")]$PlaceAddress,
 
        [Parameter(Mandatory=$false)]

        [ValidateSet("Address","Message", "File", "List","Link")]

        [Alias("Content")] 

        $Type,
         
        [switch]$UploadFile,
        [string[]]$items,

        $title="PushBullet Message",
        $msg)

begin{
        $api = "xxxxxxxxxxxx" #Hard set the API key here.
        $PushURL = "https://api.pushbullet.com/v2/pushes"
        $devices = "https://api.pushbullet.com/v2/devices"
        $uploadRequestURL   = "https://api.pushbullet.com/v2/upload-request"
        $uploads = "https://s3.amazonaws.com/pushbullet-uploads"
 
        $cred = New-Object System.Management.Automation.PSCredential ($api,(ConvertTo-SecureString $api -AsPlainText -Force))
 
        if (($PlaceName) -or ($PlaceAddress)){$type = "address"}
     }
 
process{
 
            switch($Type)
                                                                                                                                                                                                                                                                                                                                                                                    {

            'Address'
            {
                $body = @{
                type = "address"
                title = $Placename
                address = $PlaceAddress
                         }
            }

            'Message'
            {
                $body = @{
                type = "note"
                title = $title
                body = $msg
                         }
            }

            'List'   
            {
                $body = @{
                type = "list"
                title = $title
                items = $items
                         } 
                "body preview"
                $body
            }

            'Link'   
            {
                $body = @{
                type = "link"
                title = $title
                body = $msg
                url = $url
                         }
            }

            'File'   
            {
                If ($UploadFile) 
                {  
                    $UploadRequest = @{
                    file_name = $FileName
                    fileType  = $FileType
                                      }
         
                #Ref: Pushing files https://docs.pushbullet.com/v2/pushes/#pushing-files
                # "Once the file has been uploaded, set the file_name, file_url, and file_type returned in the response to the upload request as the parameters for a new push with type=file."
                #Create Upload request first

                $attempt = Invoke-WebRequest -Uri $uploadRequestURL -Credential $cred -Method Post -Body $UploadRequest -ErrorAction SilentlyContinue
                If ($attempt.StatusCode -eq "200")
                {
                    Write-Verbose "Upload Request OK"
                }
                else 
                {
                    Write-Warning "error encountered, check `$Uploadattempt for more info"
                    $global:Uploadattempt = $attempt
                }
 
                #Have to include the data field from the full response in order to begin an upload
                $UploadApproval = $attempt.Content | ConvertFrom-Json | select -ExpandProperty data 
         
                #Have to append the file data to the Upload request        
                $UploadApproval | Add-Member -Name "file" -MemberType NoteProperty -Value ([System.IO.File]::ReadAllBytes((get-item C:\TEMP\upload.txt).FullName))
 
                #Upload the file and get back the url
                #$UploadAttempt = 
                #Invoke-WebRequest -Uri $uploads -Credential $cred -Method Post -Body $UploadApproval -ErrorAction SilentlyContinue
                #Doesn't work...maybe invoke-restMethod is the way to go?
             
                Invoke-WebRequest -Uri $uploads -Method Post -Body $UploadApproval -ErrorAction SilentlyContinue
                #End Of Upload File scriptblock
                }
            Else {
                #If we don't need to upload the file
                 
                $body = @{
                    type = "file"
                    file_name = $fileName
                    file_type = $filetype
                    file_url = $url
                    body = $msg
                    } 
                 
            }
            $global:UploadApproval = $UploadApproval
            BREAK
            #End of File switch
 
            }
        }
 
            write-debug "Test-value of `$body before it gets passed to Invoke-WebRequest"
 
            $Sendattempt = Invoke-WebRequest -Uri $PushURL -Credential $cred -Method Post -Body $body -ErrorAction SilentlyContinue
 
            If ($Sendattempt.StatusCode -eq "200")
            {
                Write-Verbose "OK"
            }
            else 
            {
                Write-Warning "error encountered, check `$attempt for more info"
                $global:Sendattempt = $Sendattempt  
            }
        }
 
end{$global:Sendattempt = $Sendattempt}
 
 
}