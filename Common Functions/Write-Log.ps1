Function Write-Log
{
	<#
	.SYNOPSIS
		A function to write ouput messages to a logfile and the console.
	
	.DESCRIPTION
		This function is designed to send messages to a logfile of your choosing.
		Use it to replace something like write-host for a more long term log.
	
	.PARAMETER Message
		The message being written to the log file.
	
	.EXAMPLE
		PS C:\> Write-Log -StrMessage 'This message being written has no severity.'
	
	.NOTES
		N/A
#>
	Param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[String]$Message
	)
    
	add-content -path $($Paths.logfilepath) -value ($Message)
    Write-Output $Message	
}