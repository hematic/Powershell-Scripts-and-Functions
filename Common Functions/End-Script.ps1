Function End-Script
{
	
	<#
	.SYNOPSIS
		A function to wrap up the end of the script.
	
	.DESCRIPTION
		Function has multiple tasks:
		1) Out-files a list of user created variables and their values.
			a) Assuming you are also using my Get-UserVariables function.
		2) Out-files the contents of $Error to a log for later.
		3) Out-files $Result
		4) Terminates the Script.
		5) Note that the paths to these out-files will need to be set in the main script somewhere.
		6) $Scriptname needs to be set as a variable.
	
	.PARAMETER $Result
		The result string to outfile.
	
	.EXAMPLE
		PS C:\> End-Script -Result $Result
	
	.NOTES
		N/A
	#>
	
	param
		(
		[parameter(Mandatory = $true)]
		[String]$Result
		)
	
	$Mystuff = Get-UserVariables
	Out-File -InputObject $MyStuff -FilePath $OutVarPath
	Out-File -InputObject $Error -FilePath $ErrorPath
	Out-File -InputObject $Result -Filepath $ResultsPath
	Write-Log ("********************************")
	Write-Log ("*****  $($ScriptName) Ends *****")
	Write-Log ("********************************")
	exit;
	
}