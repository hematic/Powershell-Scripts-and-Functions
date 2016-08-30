Function EnableRDP {

Param (
        <#
        .PARAMETER Computer = The ComputerName of the computer to enable RDP for.
        #>
                [Parameter(Mandatory=$True,Position=0)]
                [String]$Computer
      )

$RDP = Get-WmiObject -Class Win32_TerminalServiceSetting `
			-Namespace root\CIMV2\TerminalServices `
			-Computer $Computer `
			-Authentication 6 `
			-ErrorAction Stop

$result = $RDP.SetAllowTsConnections(1,1)
   if($result.ReturnValue -eq 0) {
   Write-output "$Computer : Enabled RDP Successfully"

 } else {
   Write-output "$Computer : Failed to enabled RDP"
 }

 }

 EnableRDP