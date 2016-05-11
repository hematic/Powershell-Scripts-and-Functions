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

function Get-DirSize {
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$Root
    )
    
    BEGIN {}
 
    PROCESS
    {

        $size = 0
        $folders = @()
        $Items = Get-ChildItem $Root -Force -ea SilentlyContinue
  
        foreach ($item in $Items) 
        {
            if ($item.PSIsContainer) 
            {
                $subfolders = @(Get-DirSize $($item.FullName))
                $size += $subfolders[-1].Size
                $folders += $subfolders
            } 
        
            else 
            {
                $size += $file.Length
            }
        }
  
        $object = New-Object -TypeName PSObject
        $object | Add-Member -MemberType NoteProperty -Name Folder -Value (Get-Item $Root).FullName
        $object | Add-Member -MemberType NoteProperty -Name Size -Value $size
        $folders += $object
    
        Write-Output $folders
  }
  
    END {}
}
<#######################################################################################>
#Variable Declarations


<#
$Servers = Get-ADComputer -LDAPFilter “(&(objectcategory=computer)(OperatingSystem=*server*))”

If(!$Servers)
{
    Write-Log "Gathering list of servers from Active Directory failed."    
}

Foreach($Server in $Servers)
{
    Invoke-Command -ComputerName $($Server.name) -ScriptBlock $Scriptblock
}
#>

$SizeThreshold = ""
$AgeThreshold = ""
$FolderArray = @()

$LogicalDrives = Get-WmiObject -Class Win32_LogicalDisk | Where { $_.DriveType -eq 3 }
Foreach($Drive in $LogicalDrives)
{
    $root = $drive.deviceid
    $Items = Get-ChildItem $Root -Attributes directory -Recurse -Force -ea SilentlyContinue

    Foreach($item in $items)
    {
        $FolderData = New-Object PSObject -Property @{
        FolderSize = $($Item.FullName) | Get-childitem | Measure-object -Sum Length
        FileCount =  $Item | Get-childitem | Measure-object -Sum Count
        FolderPath = $Item.FullName
        }

        $FolderArray += $FolderData
    }

    
}
