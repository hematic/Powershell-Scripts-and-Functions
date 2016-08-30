Function StartProcWithWait 
{

	$InstallApplication = $args[0];
    $InstallProc = New-Object System.Diagnostics.Process;
	$InstallProc.StartInfo = $InstallApplication;
	$InstallProc.Start();
	$InstallProc.WaitForExit();
	
}

$RepairNetPath = "$env:windir\temp\NetFxRepairTool.exe"
$LogFilePath = "$env:windir\temp\Netrepairlog.txt"

<#######################################################################################>
#Begin .net Repair
	
$RepairNet = New-Object System.Diagnostics.ProcessStartInfo("$RepairNetPath","/q /l $env:windir\temp");
$RepairNet.UseShellExecute = $False 
StartProcWithWait $RepairNet;
