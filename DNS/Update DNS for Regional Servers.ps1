[Array]$americasServers = Get-Content D:\Temp\ServersToChange.txt
$Credential = <credential>
$Results = @()
Foreach($Server in $asiapacServers){
    $Result = Invoke-Command -ComputerName $Server -Credential $Credential -ScriptBlock{
        $NICs = Get-WMIObject Win32_NetworkAdapterConfiguration | Where-Object {$_.IPEnabled -eq "TRUE"}
        Foreach($NIC in $NICs){
            $DNSServers = "10.10.101.77","10.10.114.206"
            $NIC.SetDNSServerSearchOrder($DNSServers) | Out-Null
            $NIC.SetDynamicDNSRegistration("TRUE") | Out-Null
        }
        Try{
            [Array]$NICs = Get-WMIObject Win32_NetworkAdapterConfiguration | Where-Object {$_.IPEnabled -eq "TRUE"}
            $Objs = @()
            Foreach($Nic in $Nics){
                $obj = New-Object -TypeName psobject -Property @{
                    nic = $Nic.Description
                    DNSSearchOrder = $Nic.DNSServerSearchOrder -join ','
                }
                $Objs += $Obj
            }
            Return $Objs
        }
        Catch{
            Write-Error $_
        } 
    }
    $Results += $Result
}

$Results | Select -Property PSComputername,NIC,DNSServerSearchOrder | Export-Excel -Path $env:temp\AmericasDNSChanges.xlsx

$Splat = @{
    Body = 'See attachment for Americas results'
    Attachments = "$env:temp\AmericasDNSChanges.xlsx"
    To = 'some_address@domain.com'
    From = 'jenkins@domain.com'
    CC = 'Phillip.marshall@domain.com','other.guy@domain.com'
    Subject = 'AMERICAS DNS Scripted Changes Results'
    SMTPServer = '<SMTP_Server>'
}
Send-MailMessage @Splat