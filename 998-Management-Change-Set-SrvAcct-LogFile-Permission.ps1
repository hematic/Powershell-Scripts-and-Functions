#/* 2005,2008,2008R2,2012 */

###############################################################################################################
# PowerShell Script Template
###############################################################################################################
# For use with the Auto-Install process
#
# When called, the script will recieve a single hashtable object passed as a parameter.
# This object contains the following items by default:
#	SqlVersion - version being installed
#	SqlEdition - edition being installed
#	ServiceAccount - service account being used
#	ServicePassword - password for service account
#	SysAdminPassword - SA password
#	FilePath - working folder for auto-install
#	DataCenter - data center the server is located in
#	DbaTeam - responsible DBA team
#	InstanceName - SQL instance name
#	ProductStringName - product features being installed
#	Environment - environment for server (dev, qa, bcp, prod)
#
# Additional items can be added by using the Add method on the $ht object in the Start-SqlSpade.ps1 script
#
# You can access any of these items using the following syntax: $configParams["SqlVersion"]
#
# This script should be placed in the appropriate scripts folder and will be automatically called during the 
# auto-install process.
#
# The script should be saved using the following naming convention:
# 
# Pre Install Script (Save to PreScripts Folder) -
# -------------------------------------------------
# Pre-100-OS-[ScriptName].ps1
# Pre-200-Service-[ScriptName].ps1
#
# Post Install Script (Save to SQLScripts Folder) -
# -------------------------------------------------
# 100-OS-[ScriptName].ps1
# 200-Service-[ScriptName].ps1
# 300-Server-[ScriptName].ps1
# 400-Database-[ScriptName].ps1
# 500-Table-[ScriptName].ps1
# 600-View-[ScriptName].ps1
# 700-Procedure-[ScriptName].ps1
# 800-Agent-[ScriptName].ps1
# 900-Management-ScriptName.ps1
###############################################################################################################

function enable-privilege
{
	param (
		## The privilege to adjust. This set is taken from
		## http://msdn.microsoft.com/en-us/library/bb530716(VS.85).aspx
		[ValidateSet(
					 "SeAssignPrimaryTokenPrivilege", "SeAuditPrivilege", "SeBackupPrivilege",
					 "SeChangeNotifyPrivilege", "SeCreateGlobalPrivilege", "SeCreatePagefilePrivilege",
					 "SeCreatePermanentPrivilege", "SeCreateSymbolicLinkPrivilege", "SeCreateTokenPrivilege",
					 "SeDebugPrivilege", "SeEnableDelegationPrivilege", "SeImpersonatePrivilege", "SeIncreaseBasePriorityPrivilege",
					 "SeIncreaseQuotaPrivilege", "SeIncreaseWorkingSetPrivilege", "SeLoadDriverPrivilege",
					 "SeLockMemoryPrivilege", "SeMachineAccountPrivilege", "SeManageVolumePrivilege",
					 "SeProfileSingleProcessPrivilege", "SeRelabelPrivilege", "SeRemoteShutdownPrivilege",
					 "SeRestorePrivilege", "SeSecurityPrivilege", "SeShutdownPrivilege", "SeSyncAgentPrivilege",
					 "SeSystemEnvironmentPrivilege", "SeSystemProfilePrivilege", "SeSystemtimePrivilege",
					 "SeTakeOwnershipPrivilege", "SeTcbPrivilege", "SeTimeZonePrivilege", "SeTrustedCredManAccessPrivilege",
					 "SeUndockPrivilege", "SeUnsolicitedInputPrivilege")]
		$Privilege,
		## The process on which to adjust the privilege. Defaults to the current process.

		$ProcessId = $pid,
		## Switch to disable the privilege, rather than enable it.

		[Switch]$Disable
	)
	
	## Taken from P/Invoke.NET with minor adjustments.
	$definition = @'
 using System;
 using System.Runtime.InteropServices;
  
 public class AdjPriv
 {
  [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
  internal static extern bool AdjustTokenPrivileges(IntPtr htok, bool disall,
   ref TokPriv1Luid newst, int len, IntPtr prev, IntPtr relen);
  
  [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
  internal static extern bool OpenProcessToken(IntPtr h, int acc, ref IntPtr phtok);
  [DllImport("advapi32.dll", SetLastError = true)]
  internal static extern bool LookupPrivilegeValue(string host, string name, ref long pluid);
  [StructLayout(LayoutKind.Sequential, Pack = 1)]
  internal struct TokPriv1Luid
  {
   public int Count;
   public long Luid;
   public int Attr;
  }
  
  internal const int SE_PRIVILEGE_ENABLED = 0x00000002;
  internal const int SE_PRIVILEGE_DISABLED = 0x00000000;
  internal const int TOKEN_QUERY = 0x00000008;
  internal const int TOKEN_ADJUST_PRIVILEGES = 0x00000020;
  public static bool EnablePrivilege(long processHandle, string privilege, bool disable)
  {
   bool retVal;
   TokPriv1Luid tp;
   IntPtr hproc = new IntPtr(processHandle);
   IntPtr htok = IntPtr.Zero;
   retVal = OpenProcessToken(hproc, TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref htok);
   tp.Count = 1;
   tp.Luid = 0;
   if(disable)
   {
    tp.Attr = SE_PRIVILEGE_DISABLED;
   }
   else
   {
    tp.Attr = SE_PRIVILEGE_ENABLED;
   }
   retVal = LookupPrivilegeValue(null, privilege, ref tp.Luid);
   retVal = AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
   return retVal;
  }
 }
'@
	
	$processHandle = (Get-Process -id $ProcessId).Handle
	$type = Add-Type $definition -PassThru
	$type[0]::EnablePrivilege($processHandle, $Privilege, $Disable)
}

enable-privilege -Privilege "SeRestorePrivilege" | out-null

$configParams = $args[0]

$homeFolder = "C:\Windows\System32\LogFiles\Sum"
$username = $configParams["ServiceAccount"]
$aryRights = @("Read", "Write", "Modify")

$Acl = (Get-Item $homeFolder).GetAccessControl('Access')

foreach ($right in $aryRights) {    
    $Ar = New-Object System.Security.AccessControl.FileSystemAccessRule($username, $right, 'ContainerInherit,ObjectInherit', 'None', 'Allow')
    $Acl.SetAccessRule($Ar)
    Set-Acl -path $homeFolder -AclObject $Acl
}

Write-Log -level "Info" -message "Read, Write and Modify settings have been sucessfully set to foler $homeFolder for login $username"