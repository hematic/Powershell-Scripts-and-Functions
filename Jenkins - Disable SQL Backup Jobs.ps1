$PassedInstances = "$ENV:SQLInstances"
[Array]$SQLInstances = $PassedInstances -split ","

$SecPassword = ConvertTo-SecureString "$ENV:SAPassword" -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential ($ENV:SAUsername, $SecPassword)
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

    $Session = New-PsSession -ComputerName $Servername -Credential $Credential -ErrorAction SilentlyContinue
    If($Session)
    {
        Invoke-Command -ComputerName $Servername -Credential $credential -ScriptBlock {
            Param($Instance)


            $bJobStatus = 0

            #SQL Statement
            $cSQLStmt = @"
            DECLARE
	            @bExecSQL BIT
	            , @bJobSwitch BIT
	            , @cSQLStmt NVARCHAR(2048);

            SET @bExecSQL = 1;
            SET @bJobSwitch = $bJobStatus;
            SET @cSQLStmt = '';

            SELECT
	            @cSQLStmt = (
		            @cSQLStmt 
		            + 'EXEC [msdb].[dbo].[sp_update_job] @job_id = N''' + CAST([s].[job_id] AS NVARCHAR(36)) + ''', @enabled = ' + CAST(@bJobSwitch AS NCHAR(1)) + ';' 		
		            + ' /*' 
		            + [s].[name]
		            + '*/' 
		            + NCHAR(13)
	            )
            FROM [msdb].[dbo].[sysjobs] AS [s]
            WHERE [s].[name] LIKE '%FULL%'
            AND [s].[name] NOT LIKE '%SYSTEM_DATABASES%';

            IF @bExecSQL = 1
	            EXEC [sys].[sp_executesql] @cSQLStmt;
            ELSE
	            PRINT @cSQLStmt;
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
            $Result

    } -argumentlist $Instance
    }

    }

    Write-Output "Time Taken          : $($Timetaken.seconds) second(s)"
    If(!$Session)
    {
        Write-Output "Result              : FAILED!! Unable to Connect!"
    }

    Else
    {
        Write-Output "Result              : Success"
    }

}

get-pssession | remove-pssession