Function Send-Email{

Param (
        <#
        .PARAMETER EmailFrom = The senders email address.
        #>
                [Parameter(Mandatory=$True,Position=0)]
                [String]$EmailFrom,
        <# 
        .PARAMETER Emailto = The Recipients email address.
        #>
                [Parameter(Mandatory=$True,Position=1)]
                [String]$EmailTo,
                
        <# 
        .PARAMETER Subject = The Subject of the email.
        #>
        
                [Parameter(Mandatory=$True,Position=2)]
                [String]$Subject,
        
        <# 
        .PARAMETER Body = The body of the email.
        #>
        
                [Parameter(Mandatory=$True,Position=3)]
                [String]$Body,
				
		<# 
        .PARAMETER SMTPServer = The address of the SMTP server to use. EX : smtp.gmail.com
        #>
        
                [Parameter(Mandatory=$True,Position=4)]
                [String]$SMTPServer,
				
		<# 
        .PARAMETER SMTPPort = The SMTP port to use.
        #>
        
                [Parameter(Mandatory=$True,Position=5)]
                [String]$SMTPPort,
				
				 <# 
        .PARAMETER Username = The username to send.
        #>
        
                [Parameter(Mandatory=$True,Position=6)]
                [String]$Username,
				
				 <# password to send.
        #>
        
                [Parameter(Mandatory=$True,Position=7)]
                [String]$Password
              )
 
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, $smtpport) 
$SMTPClient.EnableSsl = $true 
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($username, $password); 
$SMTPClient.Send($EmailFrom, $EmailTo, $Subject, $Body)
}