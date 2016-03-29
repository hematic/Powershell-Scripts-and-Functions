Function Get-ExpiringCerts
{

    Param (
        <#
        .PARAMETER FutureDays = How man days into the future this script looks.
        #>
                [Parameter(Mandatory=$True,Position=0)]
                [String]$FutureDays,

        <#
        .PARAMETER PastDays = Modify the variable below to control how far into the past we look for expired certificates. Use a negative number for the past and 0 for now
        #>
                [Parameter(Mandatory=$True,Position=1)]
                [String]$PastDays
      )

            # Set up a variable with a datetime object representing right now
            $now = Get-Date
            # Calculate a new datetime object that represents the past
            $Past = $now.AddDays($PastDays)
            # Calculate a new Datetime object that represents the future
            $Future = $now.AddDays($FutureDays)
            # Create an array of all the certificates on the local system
            $certs = Get-ExchangeCertificate
            # Filter the filterd list down to those whose expiration date falls within the desired range
            $expiringcerts = $certs | Where-Object {$_.notafter -ge $Past -and $_.notafter -le $future}


            #The line below simply presents the filtered list.  You can alter this as you see fit
            #$expiringcerts | sort-object notafter | Format-Table subject,friendlyname,notafter -AutoSize

            If ($expiringcerts) { 
                $msgTo = ""
                $msgFrom = ""
                $msgServer = ""
                $msgSubject = "***ALERT*** Expiring Exchange Certificates"
                $msgBody1 = "The following certificates are set to expire:`n"
                [string]$msgBody2 = foreach ($ecert in $expiringcerts) {$ecert.FriendlyName + "`t" + $(($ecert.NotAfter).ToString("MM/dd/yyyy")) + " `n"}
                $msgBody = $msgBody1 + $msgBody2

            Send-MailMessage -SMTPServer $msgServer -From $msgFrom -To $msgTo -Subject $msgSubject -Body $msgBody2
}
}