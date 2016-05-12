<#######################################################################################>
#Function Declarations
Function Write-Log
{

	
	Param
		(
		[Parameter(Mandatory = $True, Position = 0)]
		[String]$Message,
		[Parameter(Mandatory = $False, Position = 1)]
		[INT]$Severity
	)
	
	$Note = "[NOTE]"
	$Warning = "[WARNING]"
	$Problem = "[ERROR]"
	[string]$Date = get-date
	
	switch ($Severity)
	{
		1 { add-content -path $LogFilePath -value ($Date + "`t:`t" + $Note + $Message) }
		2 { add-content -path $LogFilePath -value ($Date + "`t:`t" + $Warning + $Message) }
		3 { add-content -path $LogFilePath -value ($Date + "`t:`t" + $Problem + $Message) }
		default { add-content -path $LogFilePath -value ($Date + "`t:`t" + $Message) }
	}
	
	
}

Function Process-Drive
{

	
	Param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[String]$DriveID
	)

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
#


$Servers = Get-ADComputer -LDAPFilter “(&(objectcategory=computer)(OperatingSystem=*server*))”

If(!$Servers)
{
    Write-Log "Gathering list of servers from Active Directory failed."
    exit  
}

Foreach($Server in $Servers)
{
    Invoke-Command -ComputerName $($Server.name) -ScriptBlock {
    
    
            #[Int]$SizeThreshold = 1073741824 -1 Gig
            [Int]$SizeThreshold = 314572800
            $AgeThreshold = (get-date).AddDays(-60)
            [Array]$FolderArray = @()

            $LogicalDrives = Get-WmiObject -Class Win32_LogicalDisk | Where { $_.DriveType -eq 3 }

            If(!$LogicalDrives)
            {
                Write-Log "No logical drives were detected on this machine."
                exit
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
    
    }
}


