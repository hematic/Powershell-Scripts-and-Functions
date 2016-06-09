[Array]$SQLInstances = @($ENV:SQLInstances)
$Credential = Get-credential
$TotalInstances = $($Sqlinstances.count)
Write-output "********************************************"
Write-output "Servers to Process  : $($Sqlinstances.count)"
Write-output "********************************************"
[Int]$ProcessedCount = 0
Foreach($Instance in $SQLInstances)
{
    [Int]$ProcessedCount++ | out-null
    If($Instance -like '*\*')
    {
        $ServerName = ([regex]::matches($Instance, "(.+)(?:\\)")).groups[1].value
    }
    Else
    {
        $Servername = $Instance
    }
    
    Write-output ""
    Write-Output "Processing ($ProcessedCount of $TotalInstances)"
    Write-output "Processing Instance : $Instance"
    Write-output "Parsed Servername   : $Servername"
    $TimeTaken = Measure-command {

    $IC = Invoke-Command -ComputerName $Servername -Credential $Credential -ScriptBlock {
    Param($Instance)


    $bJobStatus = 1

    #SQL Statement
    $cSQLStmt = @"
SELECT TABLE_NAME
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_CATALOG='dbName'
"@

    $SqlCon = New-Object System.Data.SqlClient.SqlConnection
    $SqlCon.ConnectionString = "Server = $Instance; Database = MSDB; Integrated Security = True; Trusted_Connection = True"
    $SqlCon.Open()

    #-- SQL command to get instance list
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
    $SqlCmd.CommandTimeout = 10
    $SqlCmd.CommandText = $cSQLStmt
    $SqlCmd.Connection = $SqlCon
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $SqlCmd

    $DataSet = New-Object System.Data.DataSet
    $result = $SqlAdapter.Fill($DataSet)
    $SqlCon.close()

    } -argumentlist $Instance
    }
    Write-Output "Time Taken          : $($Timetaken.seconds) second(s)"
    Write-Output "Result              : $IC"

}