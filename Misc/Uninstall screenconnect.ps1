#Read installation information from the registry and get a list of installed software.
$RegistryLocation = Get-ChildItem "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
$Name = "screenconnect client (6f88598b921cda40)"
	
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
            $args="/uninstall /q $code"
				
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
                $args = "/x $($match.Value) /qn"
					
				[diagnostics.process]::start("msiexec", $args).WaitForExit()
					
				$Result = 'Uninstall Attempted'
                }

            else { $Result = "Unable to uninstall $name" }

        }
    }
}


Remove-Item -path "$($env:windir)\servermonitor\packages\screenconnect" -Force -Recurse;
Return $result;
