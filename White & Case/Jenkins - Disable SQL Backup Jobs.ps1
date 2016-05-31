$secpasswd = ConvertTo-SecureString $ENV:ADMPassword -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential ($ENV:ADMUsername, $secpasswd)
$Target = $ENV:SQLInstance


Invoke-Command -ScriptBlock {

$cSQLInstance = "AM1FMDB903\FINDATA"

<# 
Flag to determine job status
OPTION:
0 - Disabled
1 - Enabled
#>
$bJobStatus = 1

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

Invoke-Sqlcmd -Query $cSQLStmt -ServerInstance $($Target + "\FINDATA") -AbortOnError -Verbose

} -ComputerName $Target -Credential $Credential