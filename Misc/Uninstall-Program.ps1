Function Uninstall-Program
{
		<#
	.SYNOPSIS
		A function to uninstall a piece of software.
	
	.DESCRIPTION
		This function allows you to pass a software name and it will
		attempt to uninstall it.
	
	.PARAMETER Name
		The name of the software to look for.
	
	.EXAMPLE
		PS C:\> Uninstall-Program -Name 'Screenconnect'
	
	.NOTES
		N/A
#>
	
	Param
		(
			[Parameter(Mandatory = $True, Position = 0)]
			[String]$Name
		)
	
	#Read installation information from the registry and get a list of installed software.
	$RegistryLocation = Get-ChildItem "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
	
    foreach ($RegistryItem in $RegistryLocation)
    {
        #Check for the software we are looking for...
        if ((Get-itemproperty $registryItem.PSPath).DisplayName -like "*$name*")
        {
            # Get the product code if possible
            $productCode = (Get-itemproperty $registryItem.PSPath).ProductCode
            
            # If a product code is available, uninstall using it
            if ([string]::IsNullOrEmpty($productCode) -eq $false)
            {
                Write-Output "Uninstalling $name, ProductCode:$code"
            
                $args="/uninstall $code"
				
				[diagnostics.process]::start("msiexec", $args).WaitForExit()
				
				$Result = 'Uninstall Attempted'
				
			}
			
			# If there is no product code, try to read the uninstall string
            else
            {
                $uninstallString = (Get-itemproperty $registryItem.PSPath).UninstallString
                
                if ([string]::IsNullOrEmpty($uninstallString) -eq $false)
                {
                    # Grab the product key and create an argument string
                    $match = [RegEx]::Match($uninstallString, "{.*?}")
                    $args = "/x $($match.Value) /qb"
					
					[diagnostics.process]::start("msiexec", $args).WaitForExit()
					
					$Result = 'Uninstall Attempted'
                    
                }
                else { $Result = "Unable to uninstall $name" }
            }
        }
    }
}

