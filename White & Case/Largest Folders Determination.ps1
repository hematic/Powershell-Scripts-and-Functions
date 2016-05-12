﻿<#
	.SYNOPSIS
		A script to recurse through every folder on every logical drive and find folders
        that are over a certain size that have not been accessed since a certain time.
	
	.Notes
		This script needs two key variables set:

        1) "SizeThreshold" which is an [INT] and should be in bytes. For example if you want
           a 1 gigabyte threshold use 1073741824.

        2) "AgeThreshold" which is a [DateTime] where you specify the days to go back from the current
           date. For example if today is 5/12/16 and you want to look for folders not accessed in the 
           last 60 days "AgeThreshold" should be set to [DateTime]$AgeThreshold = (get-date).AddDays(-60).
	
	.PARAMETERS
		This script expects no parameters.
    
    .RETURNS
        This script has 3 possible return values:
        
        1) "No Bad Folders"
        2) "No Drives Detected"
        3) $Suspectfolders which will be an [Array] of objects.
	
	.VERSION
        1.0 - Initial Revision.
	
	.Author
		Phillip Marshall
        Sr Datacenter Automation Engineer
        White & Case
#>

<#######################################################################################>
#Function Declarations
Function Write-Log
{
	<#
	.SYNOPSIS
		A function to write ouput messages to a logfile.
	
	.DESCRIPTION
		This function is designed to send timestamped messages to a logfile of your choosing.
		Use it to replace something like write-host for a more long term log.
	
	.PARAMETER Message
		The message being written to the log file.
	
	.EXAMPLE
		PS C:\> Write-Log -Message 'This is the message being written out to the log.' 
	
	.NOTES
		N/A
#>
	
	Param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[String]$Message
	)

	add-content -path $LogFilePath -value ($Message)
    Write-Output $Message
}

Function Process-Drive
{
<#
	.SYNOPSIS
		A function to recurse through all folders on a drive.
	
	.DESCRIPTION
		This function is designed to take a drive letter and recursivley loop through
        all folders on that drive and build an object for each one. The object contains:
        1) The size of the folder in bytes.
        2) The number of files in that folder.
        3) The full path of the folder.
        4) The lastaccesstime of the folder.

        Once the object is built for each folder an array is returned to the main script
        that contains all the objects.
	
	.PARAMETER DriveID
		The drive letter to be scanned for folders. eg; C: or d:
	
	.EXAMPLE
		PS C:\> Process-Drive -DriveID $DrivePath
	
	.NOTES
		N/A
#>
	Param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[String]$DriveID
	)
    
    [Array]$TempFolderArray = @()
    $Items = Get-ChildItem $DriveID -Attributes directory -Recurse -Force -ea SilentlyContinue

    If(!$Items)
    {
        Write-Log "No folders found on Drive $($Drive.DeviceID) for Server $(hostname)"
        Return "No Folders"
    }

    Write-Log "Beginning to process $($Items.count) on drive $Root"

    Foreach($Item in $Items)
    {
        $FolderStats = Get-ChildItem -Path $($Item.fullname) -ErrorAction SilentlyContinue

        $FolderData = New-Object PSObject -Property @{
            FolderSize =  $Folderstats | measure-object -property length -sum -ErrorAction SilentlyContinue | Select -ExpandProperty sum -ErrorAction SilentlyContinue
            FileCount =   $Folderstats | measure-object | select -ExpandProperty count -ErrorAction SilentlyContinue
            FolderPath = $Item.FullName
            LastAccessTime = $Item.LastAccessTime
        }

        $TempFolderArray += $FolderData
    }

    Return $TempFolderArray
	
}

<#######################################################################################>
#Variable Declarations
$LogFilePath = "$env:windir\temp\LargestFolderDetermination.txt"
[Int]$SizeThreshold = 314572800
[DateTime]$AgeThreshold = (get-date).AddDays(-60)
[Array]$FolderArray = @()

<#######################################################################################>
#Main Script
$LogicalDrives = Get-WmiObject -Class Win32_LogicalDisk | Where { $_.DriveType -eq 3 }

If(!$LogicalDrives)
{
    Write-Log "No logical drives were detected on this machine."
    Return "No Drives Detected"
}

Foreach($Drive in $LogicalDrives)
{
    $DriveResult = Process-Drive -DriveID $Drive.DeviceID

    If($DriveResult -Ne "No Folders")
    {
        $FolderArray += $DriveResult
    }
}

$SuspectFolders = $FolderArray | Where-Object {$_.foldersize -gt $SizeThreshold -and $_.LastAccessTime -lt $AgeThreshold} |Sort-Object -Property FolderSize -Descending  

If(!$SuspectFolders)
{
    Write-log "No large dormant folders exist on this machine."
    Return "No Bad Folders"
}
    
Return $Suspectfolders