<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2016 v5.2.122
	 Created on:   	6/22/2016 10:53 AM
	 Created by:   	Phillip Marshall
	 Organization: 	White & Case
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>


#region Function Declarations

function New-VirtualMachine
{
<#
	.SYNOPSIS
		This function creates a new virtual machine.
	
	.DESCRIPTION
		Using the VSphere commandlets this function takes input from the GUI and fires 
		off a call to create a new VirtualMachine. It includes support for all parameters 
		that "New-VM" supports.
	
	.NOTES
		For a description of all of the parameters please see:
		https://www.vmware.com/support/developer/PowerCLI/PowerCLI41/html/New-VM.html
#>
	
	param
	(
		[Parameter(Mandatory = $true)]
		[String]$VMHost,
		[Parameter(Mandatory = $false)]
		[ValidateSet('unknown', 'v4', 'v7', 'v8', 'v9', 'v10', 'v11')]
		[String]$Version,
		[Parameter(Mandatory = $true)]
		[String]$Name,
		[Parameter(Mandatory = $false)]
		[String]$ResourcePool,
		[Parameter(Mandatory = $false)]
		[String]$Vapp,
		[Parameter(Mandatory = $false)]
		[String]$Location,
		[Parameter(Mandatory = $false)]
		[String]$Datastore,
		[Parameter(Mandatory = $false)]
		[String]$DiskMB,
		[Parameter(Mandatory = $false)]
		[String]$DiskPath,
		[Parameter(Mandatory = $false)]
		[ValidateSet('EagerZeroedThick', 'Thick', 'Thick2GB', 'Thin', 'Thin2GB')]
		[String]$DiskStorageFormat,
		[Parameter(Mandatory = $false)]
		[String]$MemoryMB,
		[Parameter(Mandatory = $false)]
		[String]$NumCPU,
		[Parameter(Mandatory = $false)]
		[Bool]$Floppy,
		[Parameter(Mandatory = $false)]
		[Bool]$CD,
		[Parameter(Mandatory = $false)]
		[String]$GuestID,
		[Parameter(Mandatory = $false)]
		[String]$AlternateGuestName,
		[Parameter(Mandatory = $false)]
		[String]$NetworkName,
		[Parameter(Mandatory = $false)]
		[String]$HARestartPriority,
		[Parameter(Mandatory = $false)]
		[String]$HAIsolationResponse,
		[Parameter(Mandatory = $false)]
		[String]$DrsAutomationLevel,
		[Parameter(Mandatory = $false)]
		[String]$VMswapFilePolicy,
		[Parameter(Mandatory = $false)]
		[String]$Server,
		[Parameter(Mandatory = $false)]
		[String]$Description,
		[Parameter(Mandatory = $true)]
		[String]$Template,
		[Parameter(Mandatory = $false)]
		[String]$RunAsync,
		[Parameter(Mandatory = $false)]
		[String]$OSCustomizationSpec,
		[Parameter(Mandatory = $true)]
		[String]$VMFilePath,
		[Parameter(Mandatory = $false)]
		[String]$VM
	)
	
	#TODO: Place script here
}


#endregion
