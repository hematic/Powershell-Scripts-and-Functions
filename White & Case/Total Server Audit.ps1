#########################################################################
#Function Declarations

Function Write-HTML
{
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [String]$Message
    )
    $FilePath = "C:\windows\temp\" + $Target + " " + $date.day + "-" + $date.Month + "-" + $Date.year +".htm"
    Out-file -InputObject $Message -Append -encoding ascii -FilePath $Filepath
}

$secpasswd = ConvertTo-SecureString $ENV:ADMPassword -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential ($ENV:ADMUsername, $secpasswd)
$Target = $ENV:Targetmachine
$Date = Get-Date


####################
#CREATE HTML OUTPUT#
####################

Write-HTML -Message " "
Write-HTML -Message '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">'
Write-HTML -Message "<html ES_auditInitialized='false'><head><title>Audit</title>"
Write-HTML -Message "<META http-equiv=Content-Type content='text/html; charset=windows-1252'>"

###################################	
#Start of Style Definition Section#
###################################
	
Write-HTML -Message "<STYLE type=text/css>"
Write-HTML -Message "	DIV .expando {DISPLAY: block; FONT-WEIGHT: normal; FONT-SIZE: 8pt; RIGHT: 10px; COLOR: #ffffff; FONT-FAMILY: Tahoma; POSITION: absolute; TEXT-DECORATION: underline}"
Write-HTML -Message "	TABLE {TABLE-LAYOUT: fixed; FONT-SIZE: 100%; WIDTH: 100%}"
Write-HTML -Message "	#objshowhide {PADDING-RIGHT: 10px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; Z-INDEX: 2; CURSOR: hand; COLOR: #000000; MARGIN-RIGHT: 0px; FONT-FAMILY: Tahoma; TEXT-ALIGN: right; TEXT-DECORATION: underline; WORD-WRAP: normal}"
Write-HTML -Message "	.heading0_expanded {BORDER-RIGHT: #bbbbbb 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #bbbbbb 1px solid; DISPLAY: block; PADDING-LEFT: 8px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #bbbbbb 1px solid; WIDTH: 100%; CURSOR: hand; COLOR: #FFFFFF; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #bbbbbb 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; BACKGROUND-COLOR: #cc0000}"
Write-HTML -Message "	.heading1 {BORDER-RIGHT: #bbbbbb 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #bbbbbb 1px solid; DISPLAY: block; PADDING-LEFT: 16px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 5px; BORDER-LEFT: #bbbbbb 1px solid; WIDTH: 100%; CURSOR: hand; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #bbbbbb 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; BACKGROUND-COLOR: #7BA7C7}"
Write-HTML -Message "	.heading2 {BORDER-RIGHT: #bbbbbb 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #bbbbbb 1px solid; DISPLAY: block; PADDING-LEFT: 16px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 5px; BORDER-LEFT: #bbbbbb 1px solid; WIDTH: 100%; CURSOR: hand; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #bbbbbb 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; BACKGROUND-COLOR: #A5A5A5}"
Write-HTML -Message "	.tableDetail {BORDER-RIGHT: #bbbbbb 1px solid; BORDER-TOP: #bbbbbb 1px solid; DISPLAY: block; PADDING-LEFT: 16px; FONT-SIZE: 8pt;MARGIN-BOTTOM: -1px; PADDING-BOTTOM: 5px; MARGIN-LEFT: 5px; BORDER-LEFT: #bbbbbb 1px solid; WIDTH: 100%; COLOR: #000000; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #bbbbbb 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; BACKGROUND-COLOR: #f9f9f9}"
Write-HTML -Message "	.filler {BORDER-RIGHT: medium none; BORDER-TOP: medium none; DISPLAY: block; BACKGROUND: none transparent scroll repeat 0% 0%; MARGIN-BOTTOM: -1px; FONT: 100%/8px Tahoma; MARGIN-LEFT: 43px; BORDER-LEFT: medium none; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: medium none; POSITION: relative}"
Write-HTML -Message "	.Solidfiller {BORDER-RIGHT: medium none; BORDER-TOP: medium none; DISPLAY: block; BACKGROUND: none transparent scroll repeat 0% 0%; MARGIN-BOTTOM: -1px; FONT: 100%/8px Tahoma; MARGIN-LEFT: 0px; BORDER-LEFT: medium none; COLOR: #000000; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: medium none; POSITION: relative; BACKGROUND-COLOR: #000000}"
Write-HTML -Message "	td {VERTICAL-ALIGN: TOP; FONT-FAMILY: Tahoma}"
Write-HTML -Message "	th {VERTICAL-ALIGN: TOP; COLOR: #cc0000; TEXT-ALIGN: left}"
Write-HTML -Message "</STYLE>"

#################################	
#Start of Control Script Section#
#################################
	
Write-HTML -Message "<SCRIPT language=vbscript>"

#######################################	
#Declare Global Variables for Routines#
#######################################	

Write-HTML -Message "	strShowHide = 1"
Write-HTML -Message '	strShow = "show"'
Write-HTML -Message '	strHide = "hide"'
Write-HTML -Message '	strShowAll = "show all"'
Write-HTML -Message '	strHideAll = "hide all"'
Write-HTML -Message "Function window_onload()"
Write-HTML -Message '	If UCase(document.documentElement.getAttribute("ES_auditInitialized")) <> "TRUE" Then'
Write-HTML -Message "		Set objBody = document.body.all"
Write-HTML -Message "		For Each obji in objBody"
Write-HTML -Message "			If IsSectionHeader(obji) Then"
Write-HTML -Message "				If IsSectionExpandedByDefault(obji) Then"
Write-HTML -Message "					ShowSection obji"
Write-HTML -Message "				Else"
Write-HTML -Message "					HideSection obji"
Write-HTML -Message "				End If"
Write-HTML -Message "			End If"
Write-HTML -Message "		Next"
Write-HTML -Message "		objshowhide.innerText = strShowAll"
Write-HTML -Message '		document.documentElement.setAttribute "ES_auditInitialized", "true"'
Write-HTML -Message "	End If"
Write-HTML -Message "End Function"
Write-HTML -Message "Function IsSectionExpandedByDefault(objHeader)"
Write-HTML -Message '	IsSectionExpandedByDefault = (Right(objHeader.className, Len("_expanded")) = "_expanded")'
Write-HTML -Message "End Function"
	
Write-HTML -Message "Function document_onclick()"
Write-HTML -Message "	Set strsrc = window.event.srcElement"
Write-HTML -Message '	While (strsrc.className = "sectionTitle" or strsrc.className = "expando")'
Write-HTML -Message "		Set strsrc = strsrc.parentElement"
Write-HTML -Message "	Wend"
Write-HTML -Message "	If Not IsSectionHeader(strsrc) Then Exit Function"
Write-HTML -Message "	ToggleSection strsrc"
Write-HTML -Message "	window.event.returnValue = False"
Write-HTML -Message "End Function"
	
Write-HTML -Message "Sub ToggleSection(objHeader)"
Write-HTML -Message '	SetSectionState objHeader, "toggle"'
Write-HTML -Message "End Sub"
	
Write-HTML -Message "Sub SetSectionState(objHeader, strState)"
Write-HTML -Message "	i = objHeader.sourceIndex"
Write-HTML -Message "	Set all = objHeader.parentElement.document.all"
Write-HTML -Message '	While (all(i).className <> "container")'
Write-HTML -Message "		i = i + 1"
Write-HTML -Message "	Wend"
Write-HTML -Message "	Set objContainer = all(i)"
Write-HTML -Message '	If strState = "toggle" Then'
Write-HTML -Message '		If objContainer.style.display = "none" Then'
Write-HTML -Message '			SetSectionState objHeader, "show" '
Write-HTML -Message "		Else"
Write-HTML -Message '			SetSectionState objHeader, "hide" '
Write-HTML -Message "		End If"
Write-HTML -Message "	Else"
Write-HTML -Message "		Set objExpando = objHeader.children.item(1)"
Write-HTML -Message '		If strState = "show" Then'
Write-HTML -Message '			objContainer.style.display = "block" '
Write-HTML -Message "			objExpando.innerText = strHide"
	
Write-HTML -Message '		ElseIf strState = "hide" Then'
Write-HTML -Message '			objContainer.style.display = "none" '
Write-HTML -Message "			objExpando.innerText = strShow"
Write-HTML -Message "		End If"
Write-HTML -Message "	End If"
Write-HTML -Message "End Sub"
	
Write-HTML -Message "Function objshowhide_onClick()"
Write-HTML -Message "	Set objBody = document.body.all"
Write-HTML -Message "	Select Case strShowHide"
Write-HTML -Message "		Case 0"
Write-HTML -Message "			strShowHide = 1"
Write-HTML -Message "			objshowhide.innerText = strShowAll"
Write-HTML -Message "			For Each obji In objBody"
Write-HTML -Message "				If IsSectionHeader(obji) Then"
Write-HTML -Message "					HideSection obji"
Write-HTML -Message "				End If"
Write-HTML -Message "			Next"
Write-HTML -Message "		Case 1"
Write-HTML -Message "			strShowHide = 0"
Write-HTML -Message "			objshowhide.innerText = strHideAll"
Write-HTML -Message "			For Each obji In objBody"
Write-HTML -Message "				If IsSectionHeader(obji) Then"
Write-HTML -Message "					ShowSection obji"
Write-HTML -Message "				End If"
Write-HTML -Message "			Next"
Write-HTML -Message "	End Select"
Write-HTML -Message "End Function"
	
Write-HTML -Message 'Function IsSectionHeader(obj) : IsSectionHeader = (obj.className = "heading0_expanded") Or (obj.className = "heading1_expanded") Or (obj.className = "heading1") Or (obj.className = "heading2"): End Function'
Write-HTML -Message 'Sub HideSection(objHeader) : SetSectionState objHeader, "hide" : End Sub'
Write-HTML -Message 'Sub ShowSection(objHeader) : SetSectionState objHeader, "show": End Sub'
	
Write-HTML -Message "</SCRIPT>"

Write-HTML -Message "	</HEAD>"
	
#######################
#Start of Body Section#
#######################	
	
Write-HTML -Message "<BODY>"
Write-HTML -Message "	<p><b><font face=`"Arial`" size=`"5`">$Target Audit<hr size=`"8`" color=`"#CC0000`"></font></b>"
Write-HTML -Message '	<font face="Arial" size="1"><b><i>Version 1.0 by Phillip Marshall</i></b></font><br>'
Write-HTML -Message "	<font face=`"Arial`" size=`"1`">Report generated on ' $(Get-Date) '</font></p>"
	
Write-HTML -Message "<TABLE cellSpacing=0 cellPadding=0>"
Write-HTML -Message "	<TBODY>"
Write-HTML -Message "		<TR>"
Write-HTML -Message "			<TD>"
Write-HTML -Message "				<DIV id=objshowhide tabIndex=0><FONT face=Arial></FONT></DIV>"
Write-HTML -Message "			</TD>"
Write-HTML -Message "		</TR>"
Write-HTML -Message "	</TBODY>"
Write-HTML -Message "</TABLE>"

write-output "Writing Detail for $Target"
$ComputerSystem = Invoke-Command -ScriptBlock{Get-WmiObject Win32_ComputerSystem} -ComputerName $target -Credential $Credential	

switch ($ComputerSystem.DomainRole)
    {
	    0 { $ComputerRole = "Standalone Workstation" }
	    1 { $ComputerRole = "Member Workstation" }
	    2 { $ComputerRole = "Standalone Server" }
	    3 { $ComputerRole = "Member Server" }
	    4 { $ComputerRole = "Domain Controller" }
	    5 { $ComputerRole = "Domain Controller" }
	    default { $ComputerRole = "Information not available" }
    }

$OperatingSystems = Invoke-Command -ScriptBlock{Get-WmiObject Win32_OperatingSystem} -ComputerName $target -Credential $Credential		
$TimeZone = Invoke-Command -ScriptBlock{Get-WmiObject Win32_Timezone} -ComputerName $target -Credential $Credential		
$Keyboards = Invoke-Command -ScriptBlock{Get-WmiObject Win32_Keyboard} -ComputerName $target -Credential $Credential		
$SchedTasks = Invoke-Command -ScriptBlock{Get-WmiObject Win32_ScheduledJob} -ComputerName $target -Credential $Credential		
$RecoveryOptions = Invoke-Command -ScriptBlock{Get-WmiObject Win32_OSRecoveryConfiguration} -ComputerName $target -Credential $Credential		
$BootINI = $OperatingSystems.SystemDrive + "boot.ini"


#############################################
#Start of COMPUTER DETAILS Section HTML Code#
#############################################

	
Write-HTML -Message "<DIV class=heading0_expanded>"
Write-HTML -Message "	<SPAN class=sectionTitle tabIndex=0>$target Details</SPAN>"
Write-HTML -Message "	<A class=expando href='#'></A>"
Write-HTML -Message "</DIV>"
Write-HTML -Message "<DIV class=filler></DIV>"

###########################################################	
#Start of COMPUTER DETAILS - GENERAL Sub Section HTML Code#
###########################################################

Write-HTML -Message "..Computer Details"

	
Write-HTML -Message "<DIV class=container>"
Write-HTML -Message "	<DIV class=heading1>"
Write-HTML -Message "		<SPAN class=sectionTitle tabIndex=0>General</SPAN>"
Write-HTML -Message "		<A class=expando href='#'></A>"
Write-HTML -Message "	</DIV>"
Write-HTML -Message "	<DIV class=container>"
Write-HTML -Message "		<DIV class=tableDetail>"
Write-HTML -Message "			<TABLE>"
Write-HTML -Message "				<tr>"
Write-HTML -Message "					<th width='25%'><b>Computer Name</b></font></th>"
Write-HTML -Message "					<td width='75%'>$($ComputerSystem.Name)</font></td>"
Write-HTML -Message "				</tr>"
Write-HTML -Message "				<tr>"
Write-HTML -Message "					<th width='25%'><b>Computer Role</b></font></th>"
Write-HTML -Message "					<td width='75%'>$ComputerRole</font></td>"
Write-HTML -Message "				</tr>"
Write-HTML -Message "				<tr>"
	
switch ($ComputerRole)
{
	"Member Workstation" { Write-HTML -Message "					<th width='25%'><b>Computer Domain</b></font></th>"	}
	"Domain Controller" { Write-HTML -Message "					<th width='25%'><b>Computer Domain</b></font></th>"	 }
	"Member Server" { Write-HTML -Message "					<th width='25%'><b>Computer Domain</b></font></th>"	 }
	default { Write-HTML -Message "					<th width='25%'><b>Computer Workgroup</b></font></th>"}
}
	
Write-HTML -Message "					<td width='75%'>$($ComputerSystem.Domain)</font></td>"
Write-HTML -Message "				</tr>"
Write-HTML -Message "				<tr>"
Write-HTML -Message "					<th width='25%'><b>Operating System</b></font></th>"
Write-HTML -Message "					<td width='75%'>$($OperatingSystems.Caption)</font></td>"
Write-HTML -Message "				</tr>"
Write-HTML -Message "				<tr>"
Write-HTML -Message "					<th width='25%'><b>Service Pack</b></font></th>"
Write-HTML -Message "					<td width='75%'>$($OperatingSystems.CSDVersion)</font></td>"
Write-HTML -Message "				</tr>"
Write-HTML -Message "				<tr>"
Write-HTML -Message "					<th width='25%'><b>System Root</b></font></th>"
Write-HTML -Message "					<td width='75%'>$($OperatingSystems.SystemDrive)</font></td>"
Write-HTML -Message "				</tr>"
Write-HTML -Message "				<tr>"
Write-HTML -Message "					<th width='25%'><b>Manufacturer</b></font></th>"
Write-HTML -Message "					<td width='75%'>$($ComputerSystem.Manufacturer)</font></td>"
Write-HTML -Message " 				</tr>"
Write-HTML -Message "				<tr>"
Write-HTML -Message "					<th width='25%'><b>Model</b></font></th>"
Write-HTML -Message "					<td width='75%'>$($ComputerSystem.Model)</font></td>"
Write-HTML -Message " 				</tr>"
Write-HTML -Message "				<tr>"
Write-HTML -Message "					<th width='25%'><b>Number of Processors</b></font></th>"
Write-HTML -Message "					<td width='75%'>$($ComputerSystem.NumberOfProcessors)</font></td>"
Write-HTML -Message " 				</tr>"
Write-HTML -Message "				<tr>"
Write-HTML -Message "					<th width='25%'><b>Memory</b></font></th>"
Write-HTML -Message "					<td width='75%'>$($ComputerSystem.TotalPhysicalMemory)</font></td>"
Write-HTML -Message " 				</tr>"
Write-HTML -Message "				<tr>"
Write-HTML -Message "					<th width='25%'><b>Registered User</b></font></th>"
Write-HTML -Message "					<td width='75%'>$($ComputerSystem.PrimaryOwnerName)</font></td>"
Write-HTML -Message " 				</tr>"
Write-HTML -Message " 				<tr>"
Write-HTML -Message "					<th width='25%'><b>Registered Organisation</b></font></th>"
Write-HTML -Message "					<td width='75%'>$($OperatingSystems.Organization)</font></td>"
Write-HTML -Message "				</tr>"
Write-HTML -Message "  				<tr>"
Write-HTML -Message "   				<th width='25%'><b>Last System Boot</b></font></th>"

If(!$OperatingSystems.Lastbootuptime)
{
    $LBTime = 'System has never rebooted.'
}

Else
{
    $LBTime = $OperatingSystems.ConvertToDateTime($OperatingSystems.Lastbootuptime)
}

Write-HTML -Message "					<td width='75%'>$($LBTime)</font></td>"
Write-HTML -Message "				</tr>"
Write-HTML -Message "			</TABLE>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "	</DIV>"
Write-HTML -Message "	<DIV class=filler></DIV>"
	
###########################################################
#Start of COMPUTER DETAILS - HOFIXES Sub-section HTML Code#
###########################################################
	
write-output "..Hotfix Information"
	
$colQuickFixes = Invoke-command -Scriptblock {Get-WmiObject Win32_QuickFixEngineering} -ComputerName $Target -Credential $Credential
	
Write-Output "	<DIV class=container>"
Write-Output "		<DIV class=heading1>"
Write-Output "			<SPAN class=sectionTitle tabIndex=0>HotFixes</SPAN>"
Write-Output "			<A class=expando href='#'></A>"
Write-Output "		</DIV>"
Write-Output "		<DIV class=container>"
Write-Output "			<DIV class=tableDetail>"
Write-Output "				<TABLE>"
Write-Output "					<tr>"
Write-Output "  						<th width='25%'><b>HotFix Number</b></font></th>"
Write-Output "  						<th width='75%'><b>Description</b></font></th>"
Write-Output "					</tr>"
	
ForEach ($objQuickFix in $colQuickFixes)
{
	if ($objQuickFix.HotFixID -ne "File 1")
	{
		Write-Output "				<tr>"
		Write-Output "					<td width='25%'>$($objQuickFix.HotFixID)</font></td>"
		Write-Output "					<td width='75%'>$($objQuickFix.Description)</font></td>"
		Write-Output "				</tr>"
	}
}

Write-HTML -Message "				</TABLE>"
Write-HTML -Message "			</DIV>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "	</DIV>"
Write-HTML -Message "	<DIV class=filler></DIV>"

##############################################################################	
#Start of COMPUTER DETAILS - LOGICAL DISK CONFIGURATION Sub-section HTML Code#
##############################################################################
	
Write-HTML -Message "..Logical Disks"
	
$colDisks = Invoke-command -Scriptblock {Get-WmiObject Win32_LogicalDisk} -ComputerName $Target -Credential $Credential
	
Write-HTML -Message "	<DIV class=container>"
Write-HTML -Message "		<DIV class=heading1>"
Write-HTML -Message "			<SPAN class=sectionTitle tabIndex=0>Logical Disk Configuration</SPAN>"
Write-HTML -Message "			<A class=expando href='#'></A>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "		<DIV class=container>"
Write-HTML -Message "			<DIV class=tableDetail>"
Write-HTML -Message "				<TABLE>"
Write-HTML -Message "					<tr>"
Write-HTML -Message "  						<th width='15%'><b>Drive Letter</b></font></th>"
Write-HTML -Message "  						<th width='20%'><b>Label</b></font></th>"
Write-HTML -Message "  						<th width='20%'><b>File System</b></font></th>"
Write-HTML -Message "  						<th width='15%'><b>Disk Size</b></font></th>"
Write-HTML -Message "  						<th width='15%'><b>Disk Free Space</b></font></th>"
Write-HTML -Message "  						<th width='15%'><b>% Free Space</b></font></th>"
Write-HTML -Message "  					</tr>"
	
Foreach ($objDisk in $colDisks)
{
	if ($objDisk.DriveType -eq 3)
	{
		Write-HTML -Message "					<tr>"
		Write-HTML -Message "						<td width='15%'>$($objDisk.DeviceID)</font></td>"
		Write-HTML -Message " 						<td width='20%'>$($objDisk.VolumeName)</font></td>"
		Write-HTML -Message " 						<td width='20%'>$($objDisk.FileSystem)</font></td>"
		
        $disksize = [math]::round(($objDisk.size / 1048576))
		
        Write-HTML -Message " 						<td width='15%'>$($disksize) MB</font></td>"
		
        $freespace = [math]::round(($objDisk.FreeSpace / 1048576))
		
        Write-HTML -Message " 						<td width='15%'>$($Freespace) MB</font></td>"
		
        $percFreespace=[math]::round(((($objDisk.FreeSpace / 1048576)/($objDisk.Size / 1048676)) * 100),0)
		
        Write-HTML -Message " 						<td width='15%'>$($percFreespace) %</font></td>"
		Write-HTML -Message "					</tr>"
	}
}
	
Write-HTML -Message "				</TABLE>"
Write-HTML -Message "			</DIV>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "	</DIV>"
Write-HTML -Message "	<DIV class=filler></DIV>"

#####################################################################	
#Start of COMPUTER DETAILS - NIC CONFIGURATION Sub-section HTML Code#
#####################################################################	

Write-HTML -Message "..Network Configuration"
	
$NICCount = 0
$colAdapters = Invoke-Command -scriptblock {Get-WmiObject Win32_NetworkAdapterConfiguration} -ComputerName $Target -Credential $Credential
	
Write-HTML -Message "	<DIV class=container>"
Write-HTML -Message "		<DIV class=heading1>"
Write-HTML -Message "			<SPAN class=sectionTitle tabIndex=0>NIC Configuration</SPAN>"
Write-HTML -Message "			<A class=expando href='#'></A>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "		<DIV class=container>"
Write-HTML -Message "			<DIV class=tableDetail>"
Write-HTML -Message "				<TABLE>"
	
$NICCount = 0
Foreach ($objAdapter in $colAdapters)
{
	if ($objAdapter.IPEnabled -eq "True")
	{
		$NICCount = $NICCount + 1
		If ($NICCount -gt 1)
		{
			Write-HTML -Message "			</TABLE>"
			Write-HTML -Message "				<DIV class=Solidfiller></DIV>"
			Write-HTML -Message "			<TABLE>"
		}
	Write-HTML -Message "  					<tr>"
	Write-HTML -Message "	 					<th width='25%'><b>Description</b></font></th>"
	Write-HTML -Message "    					<td width='75%'>$($objAdapter.Description)</font></td>"
	Write-HTML -Message "  					</tr>"
	Write-HTML -Message "  					<tr>"
	Write-HTML -Message "						<th width='25%'><b>Physical address</b></font></th>"
	Write-HTML -Message "						<td width='75%'>$($objAdapter.MACaddress)</font></td>"
	Write-HTML -Message " 					</tr>"
	If ($objAdapter.IPAddress -ne $Null)
	{
		Write-HTML -Message "					<tr>"
		Write-HTML -Message "						<th width='25%'><b>IP Address / Subnet Mask</b></font></th>"
		Write-HTML -Message "						<td width='75%'>$($objAdapter.IPAddress) /  $($objAdapter.IPSubnet)</font></td>"
		Write-HTML -Message "					</tr>"
		Write-HTML -Message "					</tr>"
		Write-HTML -Message "					<tr>"
		Write-HTML -Message "						<th width='25%'><b>Default Gateway</b></font></th>"
		Write-HTML -Message "						<td width='75%'>$($objAdapter.DefaultIPGateway)</font></td>"
		Write-HTML -Message "					</tr>"
		
	}
	Write-HTML -Message "					<tr>"
	Write-HTML -Message "						<th width='25%'><b>DHCP enabled</b></font></th>"
	If ($objAdapter.DHCPEnabled -eq "True")
	{
		Write-HTML -Message "						<td width='75%'>Yes</font></td>"
	}
	Else
	{
		Write-HTML -Message "						<td width='75%'>No</font></td>"
	}
	Write-HTML -Message "					</tr>"
	Write-HTML -Message "					<tr>"
	Write-HTML -Message "							<th width='25%'><b>DNS Servers</b></font></th>"
	Write-HTML -Message "							<td width='75%'>"
	If ($objAdapter.DNSServerSearchOrder -ne $Null)
	{
		Write-HTML -Message $objAdapter.DNSServerSearchOrder 
	}
	Write-HTML -Message "					</tr>"
	Write-HTML -Message "					<tr>"
	Write-HTML -Message "						<th width='25%'><b>Primary WINS Server</b></font></th>"
	Write-HTML -Message "						<td width='75%'>$($objAdapter.WINSPrimaryServer)</font></td>"
	Write-HTML -Message "					</tr>"
	Write-HTML -Message "					<tr>"
	Write-HTML -Message "						<th width='25%'><b>Secondary WINS Server</b></font></th>"
	Write-HTML -Message "						<td width='75%'>$($objAdapter.WINSSecondaryServer)</font></td>"
	Write-HTML -Message "					</tr>"
	$NICCount = $NICCount + 1
	}
}

Write-HTML -Message "				</TABLE>"
Write-HTML -Message "			</DIV>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "	</DIV>"
Write-HTML -Message "	<DIV class=filler></DIV>"

############################################################	
#Start of COMPUTER DETAILS - Software Sub-section HTML Code#
############################################################
$Namespace = Invoke-Command -ScriptBlock {(get-wmiobject -namespace "root/cimv2" -list)} -ComputerName $target -Credential $Credential	

if ($Namespace | ? {$_.name -match "Win32_Product"})
{
	Write-HTML -Message "..Installed Software"
	
	$colShares = Invoke-Command -scriptblock {get-wmiobject Win32_Product | select Name,Version,Vendor,InstallDate} -ComputerName $target -Credential $Credential
		
	Write-HTML -Message "	<DIV class=container>"
	Write-HTML -Message "		<DIV class=heading1>"
	Write-HTML -Message "			<SPAN class=sectionTitle tabIndex=0>Software</SPAN>"
	Write-HTML -Message "			<A class=expando href='#'></A>"
	Write-HTML -Message "		</DIV>"
	Write-HTML -Message "		<DIV class=container>"
	Write-HTML -Message "			<DIV class=tableDetail>"
	Write-HTML -Message "				<TABLE>"
	Write-HTML -Message "					<tr>"
	Write-HTML -Message "  						<th width='25%'><b>Name</b></font></th>"
	Write-HTML -Message "  						<th width='25%'><b>Version</b></font></th>"
	Write-HTML -Message "  						<th width='25%'><b>Vendor</b></font></th>"
	Write-HTML -Message "  						<th width='25%'><b>Install Date</b></font></th>"
	Write-HTML -Message "					</tr>"
		
	Foreach ($objShare in $colShares)
	{
		Write-HTML -Message "					<tr>"
		Write-HTML -Message "						<td width='50%'>$($objShare.Name)</font></td>"
		Write-HTML -Message "						<td width='20%'>$($objShare.Version)</font></td>"
		Write-HTML -Message "						<td width='15%'>$($objShare.Vendor)</font></td>"
		Write-HTML -Message "						<td width='15%'>$($objShare.InstallDate)</font></td>"
		Write-HTML -Message "					</tr>"
	}
	Write-HTML -Message "				</TABLE>"
	Write-HTML -Message "			</DIV>"
	Write-HTML -Message "		</DIV>"
	Write-HTML -Message "	</DIV>"
	Write-HTML -Message "	<DIV class=filler></DIV>"
		
}

################################################################
#Start of COMPUTER DETAILS - LOCAL SHARES Sub-section HTML Code#
################################################################
	
Write-HTML -Message "..Local Shares"
	
$colShares = Invoke-Command -Scriptblock {Get-wmiobject Win32_Share} -ComputerName $target -Credential $Credential
	
Write-HTML -Message "	<DIV class=container>"
Write-HTML -Message "		<DIV class=heading1>"
Write-HTML -Message "			<SPAN class=sectionTitle tabIndex=0>Local Shares</SPAN>"
Write-HTML -Message "			<A class=expando href='#'></A>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "		<DIV class=container>"
Write-HTML -Message "			<DIV class=tableDetail>"
Write-HTML -Message "				<TABLE>"
Write-HTML -Message "					<tr>"
Write-HTML -Message "  						<th width='25%'><b>Share</b></font></th>"
Write-HTML -Message "  						<th width='25%'><b>Path</b></font></th>"
Write-HTML -Message "  						<th width='50%'><b>Comment</b></font></th>"
Write-HTML -Message "					</tr>"
	
Foreach ($objShare in $colShares)
{
	Write-HTML -Message "					<tr>"
	Write-HTML -Message "						<td width='25%'>$($objShare.Name)</font></td>"
	Write-HTML -Message "						<td width='25%'>$($objShare.Path)</font></td>"
	Write-HTML -Message "						<td width='50%'>$($objShare.Caption)</font></td>"
	Write-HTML -Message "					</tr>"
}

Write-HTML -Message "				</TABLE>"
Write-HTML -Message "			</DIV>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "	</DIV>"
Write-HTML -Message "	<DIV class=filler></DIV>"

############################################################	
#Start of COMPUTER DETAILS - PRINTERS Sub-section HTML Code#
############################################################
	
Write-HTML -Message "..Printers"
	
$colInstalledPrinters = Invoke-Command -scriptblock { Get-WmiObject Win32_Printer} -ComputerName $target -Credential $Credential
	
Write-HTML -Message "	<DIV class=container>"
Write-HTML -Message "		<DIV class=heading1>"
Write-HTML -Message "			<SPAN class=sectionTitle tabIndex=0>Printers</SPAN>"
Write-HTML -Message "			<A class=expando href='#'></A>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "		<DIV class=container>"
Write-HTML -Message "			<DIV class=tableDetail>"
Write-HTML -Message "				<TABLE>"
Write-HTML -Message "					<tr>"
Write-HTML -Message "						<th width='25%'><b>Printer</b></font></th>"
Write-HTML -Message "						<th width='25%'><b>Location</b></font></th>"
Write-HTML -Message "						<th width='25%'><b>Default Printer</b></font></th>"
Write-HTML -Message "						<th width='25%'><b>Portname</b></font></th>"
Write-HTML -Message "					</tr>"
	
Foreach ($objPrinter in $colInstalledPrinters)
{
	If ($objPrinter.Name -eq "")
	{
		Write-HTML -Message "					<tr>"
		Write-HTML -Message "						<td width='100%'>No Printers Installed</font></td>"
	}
	Else
	{
		Write-HTML -Message "					<tr>"
		Write-HTML -Message "						<td width='25%'>$($objPrinter.Name)</font></td>"
		Write-HTML -Message "						<td width='25%'>$($objPrinter.Location)</font></td>"
		if ($objPrinter.Default -eq "True")
		{
			Write-HTML -Message "						<td width='25%'>Yes</font></td>"
		}
		Else
		{
			Write-HTML -Message "						<td width='25%'>No</font></td>"
		}
		Write-HTML -Message "						<td width='25%'>$($objPrinter.Portname)</font></td>"
	}
	Write-HTML -Message "					</tr>"
}

Write-HTML -Message "				</TABLE>"
Write-HTML -Message "			</DIV>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "	</DIV>"
Write-HTML -Message "	<DIV class=filler></DIV>"

############################################################	
#Start of COMPUTER DETAILS - SERVICES Sub-section HTML Code#
############################################################
	
Write-HTML -Message "..Services"
	
$colListOfServices = Invoke-command -Scriptblock { Get-WmiObject Win32_Service} -ComputerName $target -Credential $Credential
	
Write-HTML -Message "	<DIV class=container>"
Write-HTML -Message "		<DIV class=heading1>"
Write-HTML -Message "			<SPAN class=sectionTitle tabIndex=0>Services</SPAN>"
Write-HTML -Message "			<A class=expando href='#'></A>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "		<DIV class=container>"
Write-HTML -Message "			<DIV class=tableDetail>"
Write-HTML -Message "				<TABLE>"
Write-HTML -Message "  					<tr>"
Write-HTML -Message "	 					<th width='20%'><b>Name</b></font></th>"
Write-HTML -Message "	 					<th width='20%'><b>Account</b></font></th>"
Write-HTML -Message "	 					<th width='20%'><b>Start Mode</b></font></th>"
Write-HTML -Message "	 					<th width='20%'><b>State</b></font></th>"
Write-HTML -Message "	 					<th width='20%'><b>Expected State</b></font></th>"
Write-HTML -Message "  					</tr>"
	
Foreach ($objService in $colListOfServices)
{
	Write-HTML -Message " 					<tr>"
	Write-HTML -Message "	 					<td width='20%'>$($objService.Caption)</font></td>"
	Write-HTML -Message "	 					<td width='20%'>$($objService.Startname)</font></td>"
	Write-HTML -Message "	 					<td width='20%'>$($objService.StartMode)</font></td>"
	If ($objService.StartMode -eq "Auto")
	{
		if ($objService.State -eq "Stopped")
		{
			Write-HTML -Message "						<td width='20%'><font color='#FF0000'>$($objService.State)</font></td>"
			Write-HTML -Message "						<td width='25%'><font face='Wingdings'color='#FF0000'>û</font></td>"
		}
	}
	If ($objService.StartMode -eq "Auto")
	{
		if ($objService.State -eq "Running")
		{
			Write-HTML -Message "						<td width='20%'><font color='#009900'>$($objService.State)</font></td>"
			Write-HTML -Message "						<td width='20%'><font face='Wingdings'color='#009900'>ü</font></td>"
		}
	}
	If ($objService.StartMode -eq "Disabled")
	{
		If ($objService.State -eq "Running")
		{
			Write-HTML -Message "						<td width='20%'><font color='#FF0000'>$($objService.State)</font></td>"
			Write-HTML -Message "						<td width='25%'><font face='Wingdings'color='#FF0000'>û</font></td>"
		}
	}
	If ($objService.StartMode -eq "Disabled")
	{
		if ($objService.State -eq "Stopped")
		{
			Write-HTML -Message "						<td width='20%'><font color='#009900'>$($objService.State)</font></td>"
			Write-HTML -Message "						<td width='20%'><font face='Wingdings'color='#009900'>ü</font></td>"
		}
	}
	If ($objService.StartMode -eq "Manual")
	{
		Write-HTML -Message "						<td width='20%'><font color='#009900'>$($objService.State)</font></td>"
		Write-HTML -Message "						<td width='20%'><font face='Wingdings'color='#009900'>ü</font></td>"
	}
	If ($objService.State -eq "Paused")
	{
		Write-HTML -Message "						<td width='20%'><font color='#FF9933'>$($objService.State)</font></td>"
		Write-HTML -Message "						<td width='20%'><font face='Wingdings'color='#009900'>ü</font></td>"
	}
	Write-HTML -Message "  					</tr>"
}

Write-HTML -Message "				</TABLE>"
Write-HTML -Message "			</DIV>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "	</DIV>"
Write-HTML -Message "	<DIV class=filler></DIV>"

#####################################################################	
#Start of COMPUTER DETAILS - REGIONAL SETTINGS Sub-section HTML Code#
#####################################################################	

Write-HTML -Message "..Regional Options"
	
$ObjKeyboards = Invoke-Command -scriptblock {Get-WmiObject Win32_Keyboard} -ComputerName $target -Credential $Credential
	
Write-HTML -Message "	<DIV class=container>"
Write-HTML -Message "		<DIV class=heading1>"
Write-HTML -Message "			<SPAN class=sectionTitle tabIndex=0>Regional Settings</SPAN>"
Write-HTML -Message "			<A class=expando href='#'></A>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "		<DIV class=container>"
Write-HTML -Message "			<DIV class=tableDetail>"
Write-HTML -Message "				<TABLE>"
Write-HTML -Message " 					<tr>"
Write-HTML -Message "	 					<th width='25%'><b>Time Zone</b></font></th>"
Write-HTML -Message "	 					<td width='75%'>$($TimeZone.Description)</font></td>"
Write-HTML -Message " 					</tr>"
Write-HTML -Message " 					<tr>"
Write-HTML -Message "	 					<th width='25%'><b>Country Code</b></font></th>"
Write-HTML -Message "	 					<td width='75%'>$($OperatingSystems.Countrycode)</font></td>"
Write-HTML -Message " 					</tr>"
Write-HTML -Message " 					<tr>"
Write-HTML -Message "		 				<th width='25%'><b>Locale</b></font></th>"
Write-HTML -Message "		 				<td width='75%'>$($OperatingSystems.Locale)</font></td>"
Write-HTML -Message " 					</tr>"
Write-HTML -Message " 					<tr>"
Write-HTML -Message "		 				<th width='25%'><b>Operating System Language</b></font></th>"
Write-HTML -Message "		 				<td width='75%'>$($OperatingSystems.OSLanguage)</font></td>"
Write-HTML -Message " 					</tr>"
Write-HTML -Message " 					<tr>"
	
switch ($ObjKeyboards.Layout)
{
	"00000402"{ $keyb = "BG" }
	"00000404"{ $keyb = "CH" }
	"00000405"{ $keyb = "CZ" }
	"00000406"{ $keyb = "DK" }
	"00000407"{ $keyb = "GR" }
	"00000408"{ $keyb = "GK" }
	"00000409"{ $keyb = "US" }
	"0000040A"{ $keyb = "SP" }
	"0000040B"{ $keyb = "SU" }
	"0000040C"{ $keyb = "FR" }
	"0000040E"{ $keyb = "HU" }
	"0000040F"{ $keyb = "IS" }
	"00000410"{ $keyb = "IT" }
	"00000411"{ $keyb = "JP" }
	"00000412"{ $keyb = "KO" }
	"00000413"{ $keyb = "NL" }
	"00000414"{ $keyb = "NO" }
	"00000415"{ $keyb = "PL" }
	"00000416"{ $keyb = "BR" }
	"00000418"{ $keyb = "RO" }
	"00000419"{ $keyb = "RU" }
	"0000041A"{ $keyb = "YU" }
	"0000041B"{ $keyb = "SL" }
	"0000041C"{ $keyb = "US" }
	"0000041D"{ $keyb = "SV" }
	"0000041F"{ $keyb = "TR" }
	"00000422"{ $keyb = "US" }
	"00000423"{ $keyb = "US" }
	"00000424"{ $keyb = "YU" }
	"00000425"{ $keyb = "ET" }
	"00000426"{ $keyb = "US" }
	"00000427"{ $keyb = "US" }
	"00000804"{ $keyb = "CH" }
	"00000809"{ $keyb = "UK" }
	"0000080A"{ $keyb = "LA" }
	"0000080C"{ $keyb = "BE" }
	"00000813"{ $keyb = "BE" }
	"00000816"{ $keyb = "PO" }
	"00000C0C"{ $keyb = "CF" }
	"00000C1A"{ $keyb = "US" }
	"00001009"{ $keyb = "US" }
	"0000100C"{ $keyb = "SF" }
	"00001809"{ $keyb = "US" }
	"00010402"{ $keyb = "US" }
	"00010405"{ $keyb = "CZ" }
	"00010407"{ $keyb = "GR" }
	"00010408"{ $keyb = "GK" }
	"00010409"{ $keyb = "DV" }
	"0001040A"{ $keyb = "SP" }
	"0001040E"{ $keyb = "HU" }
	"00010410"{ $keyb = "IT" }
	"00010415"{ $keyb = "PL" }
	"00010419"{ $keyb = "RU" }
	"0001041B"{ $keyb = "SL" }
	"0001041F"{ $keyb = "TR" }
	"00010426"{ $keyb = "US" }
	"00010C0C"{ $keyb = "CF" }
	"00010C1A"{ $keyb = "US" }
	"00020408"{ $keyb = "GK" }
	"00020409"{ $keyb = "US" }
	"00030409"{ $keyb = "USL" }
	"00040409"{ $keyb = "USR" }
	"00050408"{ $keyb = "GK" }
	default { $keyb = "Unknown" }
}


Write-HTML -Message "		 				<th width='25%'><b>Keyboard Layout</b></font></th>"
Write-HTML -Message "		 				<td width='75%'>$($keyb)</font></td>"
Write-HTML -Message " 					</tr>"
Write-HTML -Message "				</TABLE>"
Write-HTML -Message "			</div>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "	</DIV>"
Write-HTML -Message "	<DIV class=filler></DIV>"

##############################################################	
#Start of COMPUTER DETAILS - EVENT LOGS Sub-section HTML Code#
##############################################################

Write-HTML -Message "	<DIV class=container>"
Write-HTML -Message "		<DIV class=heading1>"
Write-HTML -Message "			<SPAN class=sectionTitle tabIndex=0>Event Logs</SPAN>"
Write-HTML -Message "			<A class=expando href='#'></A>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "		<DIV class=container>"
Write-HTML -Message "			<DIV class=tableDetail>"

###################################################################################	
#Start of COMPUTER DETAILS - EVENT LOGS - EVENT LOG SETTINGS Sub-section HTML Code#
###################################################################################	

Write-HTML -Message "..Event Log Settings"
	
$colLogFiles = Invoke-Command -Scriptblock {Get-WmiObject Win32_NTEventLogFile} -ComputerName $Target -Credential $Credential
	
Write-HTML -Message "	<DIV class=container>"
Write-HTML -Message "		<DIV class=heading2>"
Write-HTML -Message "			<SPAN class=sectionTitle tabIndex=0>Event Log Settings</SPAN>"
Write-HTML -Message "			<A class=expando href='#'></A>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "		<DIV class=container>"
Write-HTML -Message "			<DIV class=tableDetail>"
Write-HTML -Message "				<TABLE>"
Write-HTML -Message "  					<tr>"
Write-HTML -Message "    					<th width='25%'><b>Log Name</b></font></th>"
Write-HTML -Message "    					<th width='25%'><b>Overwrite Outdated Records</b></font></th>"
Write-HTML -Message "  					  	<th width='25%'><b>Maximum Size</b></font></th>"
Write-HTML -Message " 					   	<th width='25%'><b>Current Size</b></font></th>"
Write-HTML -Message " 					</tr>"
	
ForEach ($objLogFile in $colLogfiles)
{
    Write-HTML -Message " 					<tr>"
    Write-HTML -Message "	 					<td width='25%'>$($objLogFile.LogFileName)</font></td>"
    If ($objLogfile.OverWriteOutdated -lt 0)
    {
	    Write-HTML -Message "	 					<td width='25%'>Never</font></td>"
    }
    if ($objLogFile.OverWriteOutdated -eq 0)
    {
	    Write-HTML -Message "	 					<td width='25%'>As needed</font></td>"
    }
    Else
    {
	    Write-HTML -Message "	 					<td width='25%'>After $($objLogFile.OverWriteOutdated) days</font></td>"
    }
    Write-HTML -Message "	 					<td width='25%'> $(($objLogfile.MaxFileSize)/1024) KB</font></td>"
    Write-HTML -Message "	 					<td width='25%'> $(($objLogfile.FileSize)/1024) KB</font></td>"
    Write-HTML -Message "  					</tr>"
}

Write-HTML -Message "				</TABLE>"
Write-HTML -Message "			</DIV>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "	</DIV>"
Write-HTML -Message "	<DIV class=filler></DIV>"


##############################################################################	
#Start of COMPUTER DETAILS - EVENT LOGS - ERROR ENTRIES Sub-section HTML Code#
##############################################################################	
Write-HTML -Message "..Event Log Errors"
		

$colLoggedEvents = Invoke-Command -Scriptblock {
                    $WmidtQueryDT = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime([DateTime]::Now.AddDays(-14))
                    Get-WmiObject -query ("Select * from Win32_NTLogEvent Where Type='Error' and TimeWritten >='" + $WmidtQueryDT + "'") 
                    } -ComputerName $Target -Credential $Credential

Write-HTML -Message "	<DIV class=container>"
Write-HTML -Message "		<DIV class=heading2>"
Write-HTML -Message "			<SPAN class=sectionTitle tabIndex=0>ERROR Entries</SPAN>"
Write-HTML -Message "			<A class=expando href='#'></A>"
Write-HTML -Message "		</DIV>"
Write-HTML -Message "		<DIV class=container>"
Write-HTML -Message "			<DIV class=tableDetail>"
Write-HTML -Message "				<TABLE>"
Write-HTML -Message "  					<tr>"
Write-HTML -Message "    					<th width='10%'><b>Event Code</b></font></th>"
Write-HTML -Message "   					<th width='10%'><b>Source Name</b></font></th>"
Write-HTML -Message "    					<th width='15%'><b>Time</b></font></th>"
Write-HTML -Message "    					<th width='10%'><b>Log</b></font></th>"
Write-HTML -Message "    					<th width='55%'><b>Message</b></font></th>"
Write-HTML -Message "  					</tr>"

ForEach ($objEvent in $colLoggedEvents)
{
	$dtmEventDate = [Management.ManagementDateTimeConverter]::ToDateTime($objEvent.TimeWritten)
	Write-HTML -Message " 					<tr>"
	Write-HTML -Message "	 					<td width='10%'>$($objEvent.EventCode)</font></td>"
	Write-HTML -Message "	 					<td width='10%'>$($objEvent.SourceName)</font></td>"
	Write-HTML -Message "	 					<td width='15%'>$($dtmEventDate)</font></td>"
	Write-HTML -Message "	 					<td width='10%'>$($objEvent.LogFile)</font></td>"
	Write-HTML -Message "	 					<td width='55%'>$($objEvent.Message)</font></td>"
	Write-HTML -Message "  					</tr>"
}

Write-HTML -Message "				</TABLE>"
Write-HTML -Message "			</DIV>" 

