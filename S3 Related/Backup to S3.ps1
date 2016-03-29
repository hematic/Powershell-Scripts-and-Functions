Function Write-Log
{
	<#
	.SYNOPSIS
		A function to write ouput messages to a logfile.
	
	.DESCRIPTION
		This function is designed to send timestamped messages to a logfile of your choosing.
		Use it to replace something like write-host for a more long term log.
	
	.PARAMETER StrMessage
		The message being written to the log file.
	
	.EXAMPLE
		PS C:\> Write-Log -StrMessage 'This is the message being written out to the log.' 
	
	.NOTES
		N/A
#>
	
	Param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[String]$Message
	)

    add-content -path $PSLog -value $Message
}

####################################################################
#Variable Declarations

    #S3Copy Vars
    [String]$S3BackupLog = "$Env:windir\temp\S3ScriptLog.txt"
    [String]$PSLog = "$Env:windir\temp\PSLog.txt"
    [String]$S3CopyLink = "http://s3.amazonaws.com/ltpremium/tools/s3copy/s3copy.exe"
    [String]$S3CopyDownloadPath = "$Env:windir\temp\S3copy.exe"
    [String]$Bucketname = 'partner_backups'

    #Supplied By LabTech
    [String]$LTBackupBaseLocation = "@LTPath@\Backup\LabTech.1.zip"
    [String]$LTBackupTempLocation = "$Env:windir\temp\"
    [String]$LTBackupFileName = "@BackupFileName@" -replace " ","_"
    [String]$AccessKey = "@AccessKey@"
    [String]$SecretKey = "@SecretKey@"

####################################################################
#Download S3Copy.exe

Write-Log "Starting Download Job for S3Copy.exe"

#Declare the Job
$S3CopyDownloadJob = Start-BitsTransfer `
    -Source $S3CopyLink `
    -Destination $S3CopyDownloadPath `
    -Asynchronous

#Make sure the transfer has completed. 
while (($S3CopyDownloadJob.JobState -eq "Transferring") -or ($S3CopyDownloadJob.JobState -eq "Connecting"))
{ 
    sleep 3;
}

Write-Log "Download Job finished. Checking status."

#Check the staus of the jobs completion. 
Switch($S3CopyDownloadJob.JobState)
{
       "Transferred" {Complete-BitsTransfer -BitsJob $S3CopyDownloadJob; Write-Log "S3Copy.exe has successfully downloaded.";}
       "Error"       {Write-Log "Failed to download S3Copy.exe. Error was : $($S3CopyDownloadJob.ErrorDescription)"; exit; }
       default       {Write-Log "Something went wrong with the download of S3Copy..."}
}


####################################################################
#Copy backup file to a temp directory

$copyResults = Copy-Item `
    -Path $LTBackupBaseLocation `
    -Destination "$LTBackupTempLocation$LTBackupFileName" `
    -Force

If((Test-Path "$LTBackupTempLocation$LTBackupFileName") -eq $False)
{
    Write-Log "Copy of the DB to the temp location failed."
    exit;
}

####################################################################
#Upload the Backup Using S3Copy

$AllArgs = "$LTBackupTempLocation $BucketName $LTBackupFileName $AccessKey $SecretKey"

Write-Log "Starting the upload now."

Start-Process `
    -FilePath $S3CopyDownloadPath `
    -ArgumentList $AllArgs `
    -RedirectStandardOutput $S3BackupLog `
    -PassThru `
    -Wait `
    -NoNewWindow
