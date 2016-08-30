$ErrorActionPreference = "silentlycontinue"
$filepath = 'D:\Repldata'
$d = [datetime](Get-ItemProperty -Path $filepath -Name LastWriteTime).lastwritetime

if ( test-path $filepath )
{
  if ( (get-item $filepath).LastWriteTime -ge (get-date).Date ) 
  {
    $Difference = ((Get-date)-($d)).minutes
        If ($Difference -lt 15)
        {
            $result = 'MODIFIED'
        }
  }
  else {$result = 'NOT MODIFIED'}  
}
Else {$Result = 'FILE MISSING'}
Return $result