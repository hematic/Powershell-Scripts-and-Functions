<#
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

$ListofComputers = Get-ADComputer -Filter  'OperatingSystem -like "*server*"' | Select -Expand name
$Credential = Get-Credential

$ListofComputers | Start-RSjob -Name {$_} -ScriptBlock {
		Param($ListofComputers)
	
	Invoke-command -computername $($ListofComputers) -Credential $Using:Credential -ScriptBlock {
        
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
        Return "No Folders"
    }

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

    [Int]$SizeThreshold = 314572800
    [DateTime]$AgeThreshold = (get-date).AddDays(-60)
    [Array]$FolderArray = @()
    $LogicalDrives = Get-WmiObject -Class Win32_LogicalDisk | Where { $_.DriveType -eq 3 }
        
    If(!$LogicalDrives)
    {
        $TempData = New-Object PSObject -Property @{
            LogicalDrives = "No Drives"
            Suspectfolders   = "No Folders"
            }
    }

    Else
    {
        Foreach($Drive in $LogicalDrives)
        {
            $DriveResult = Process-Drive -DriveID $Drive.DeviceID

            If($DriveResult -Ne "No Folders")
            {
                $FolderArray += $DriveResult
            }
        }

        $SuspectFolders = $FolderArray | Where-Object {$_.foldersize -gt $SizeThreshold -and $_.LastAccessTime -lt $AgeThreshold} |Sort-Object -Property FolderSize -Descending  
    
        $TempData = New-Object PSObject -Property @{
            LogicalDrives  = $LogicalDrives
            Suspectfolders = $SuspectFolders
            }
    }

    $TempData

}

$Tempdata

}

#$Data = Get-RSjob | Receive-RSJob