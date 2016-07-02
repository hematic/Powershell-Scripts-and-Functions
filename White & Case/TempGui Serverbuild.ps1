#region Function Declarations
#######################################################
Function Process-Choices
{
	<#
	.SYNOPSIS
		A function to process the data the user put into the form.
	
	.DESCRIPTION
		This function validates data that the user inputted into the form to
        verify it is acceptable for the folder creation function. It then
		calls the Process Queries function.
	
	.EXAMPLE
		Process-Choices
	
	.NOTES
		N/A
#>
	
	$Invalid = $False
	$Loggingbox.clear()
	
	#Validate Share
	
	#If the share dropdown is null this will output the problem to the Rich Text Box.
	Try
	{
		$ChosenShare = $ShareDropDownBox.SelectedItem.ToString()
	}
	
	Catch
	{
		Write-RichText -LogType Error -LogMsg "No Share has been chosen from the dropdown."
		$Invalid = $True
	}
	
	#Validate Region
	
	#If the region dropdown is null this will output the problem to the Rich Text Box.
	Try
	{
		$ChosenRegion = $RegionDropDownBox.SelectedItem.ToString()
	}
	
	Catch
	{
		Write-RichText -LogType Error -LogMsg "No Region has been chosen from the Dropdown."
		$Invalid = $True
	}
	
	#Validate Client Number
	
    <#
    If the Client Number Text Box is null or empty this will output the problem 
    to the Rich Text Box.
    #>
	
	$ChosenClientNumber = $ClientNumberTextBox.text
	IF ([string]::IsNullOrWhiteSpace($ChosenClientNumber))
	{
		Write-RichText -LogType Error -LogMsg "No Client Number was provided."
		$Invalid = $True
	}
	
	If (($ChosenClientNumber -as [Int]) -isnot [Int])
	{
		Write-RichText -LogType Error -LogMsg "The Client Number given is not all numeric."
		$Invalid = $True
	}
	
	#Validate Matter Number
	
    <#
    If the Matter NumberText Box is null or empty we display a message stating that only
    a client folder will be created.
    #>
	
	$ChosenMatterNumber = $MatterNumberTextBox.text
	IF ([string]::IsNullOrWhiteSpace($ChosenMatterNumber))
	{
		Write-RichText -LogType Informational -LogMsg "No Matter Number was provided. Just making client folder."
		$SkipMatter = $True
	}
	
	If (!$SkipMatter)
	{
		If (($ChosenMatterNumber -as [Int]) -isnot [Int])
		{
			Write-RichText -LogType Error -LogMsg "The Matter Number given is not all numeric."
			$Invalid = $True
		}
	}
	
	If ($Invalid -eq $True)
	{
		Return
	}
	
	Else
	{
		Write-RichText -LogType Informational -LogMsg "Chosen Share`t: $ChosenShare"
		Write-RichText -LogType Informational -LogMsg "Chosen Region`t: $ChosenRegion"
		Write-RichText -LogType Informational -LogMsg "ClientNumber`t: $ChosenClientnumber"
		Write-RichText -LogType Informational -LogMsg "MatterNumber`t: $ChosenMatterNumber"
		
		if ($SkipMatter)
		{
			Process-Queries -SkipMatter $True
		}
		else
		{
			Process-Queries -SkipMatter $false
		}
	}
}

Function Process-Queries
{
	param
	(
		[parameter(Mandatory = $true)]
		[Bool]$SkipMatter
	)
	
	$Invalid = $False
	$RetrievedClientname = Query-ClientName -credential $Glocredential -Clnum $($ClientNumbertextBox.text)
	
	If ($RetrievedClientname -eq $False)
	{
		Write-richtext -LogType Error -LogMsg "`t No client name found in the database matching that number."
		Return
	}
	
	ElseIf ([string]::IsNullOrWhiteSpace($RetrievedClientname))
	{
		Write-RichText -LogType error -LogMsg "`t No client name found in the database matching that number."
		Return
	}
	
	Else
	{
		Write-RichText -LogType Success -LogMsg "Retrieved Client name was : $RetrievedClientname"
	}
	
	If ($SkipMatter -eq $false)
	{
		$RetrievedMattername = Query-MatterName -Credential $Glocredential -Clnum $($ClientNumbertextBox.text) -Matternum $($matterNumbertextBox.text)
		
		If ($RetrievedMattername -eq $False)
		{
			Write-richtext -LogType Error -LogMsg "`t Unable to connect to the database."
			Return
		}
		
		ElseIf ([string]::IsNullOrWhiteSpace($RetrievedMattername))
		{
			Write-RichText -LogType error -LogMsg "`t No matter name found in the database matching that number."
			Return
		}
		
		Else
		{
			Write-RichText -LogType Success -LogMsg "Retrieved Matter name was :  $RetrievedMattername"
		}
		
		Create-Share -Share $($ShareDropDownBox.SelectedItem.ToString()) `
					 -Region $($RegionDropDownBox.SelectedItem.ToString()) `
					 -Clientname $RetrievedClientname `
					 -MatterName $RetrievedMattername
	}
	
	Else
	{
		Create-Share -Share $($ShareDropDownBox.SelectedItem.ToString()) `
					 -Region $($RegionDropDownBox.SelectedItem.ToString()) `
					 -Clientname $RetrievedClientname
	}
	
}

Function Clear-Choices
{
	<#
	.SYNOPSIS
		A function to clear all user input
	
	.DESCRIPTION
		This function is designed to clear all user-generated input from the form.
	
	.EXAMPLE
		Clear-Choices
	
	.NOTES
		N/A
#>
	
	$ShareDropDownBox.SelectedIndex = -1
	$ShareDropDownBox.SelectedIndex = -1
	$RegionDropDownBox.SelectedIndex = -1
	$RegionDropDownBox.SelectedIndex = -1
	$ClientNumberTextBox.Text = ''
	$MatterNumberTextBox.Text = ''
	$LoggingBox.Text = ''
}

Function Write-RichText
{
	<#
	.SYNOPSIS
		A function to output text to a Rich Text Box.
	
	.DESCRIPTION
		This function appends text to a Rich Text Box and colors it based 
        upon the type of message being displayed.

    .PARAM Logtype
        Used to determine if the text is a success or error message or purely
        informational.

    .PARAM LogMSG
        The message to be added to the RichTextBox.
	
	.EXAMPLE
		Write-Richtext -LogType Error -LogMsg "This is an Error."
		Write-Richtext -LogType Success -LogMsg "This is a Success."
		Write-Richtext -LogType Informational -LogMsg "This is Informational."
	
	.NOTES
		N/A
#>
	
	Param
	(
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
		[String]$LogType,
		[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline, ValueFromPipelineByPropertyName)]
		[String]$LogMsg
	)
	
	switch ($logtype)
	{
		Error {
			$Loggingbox.SelectionColor = 'Red'
			$Loggingbox.AppendText("`n $logtype : $logmsg")
			
		}
		Success {
			$Loggingbox.SelectionColor = 'Green'
			$Loggingbox.AppendText("`n $logmsg")
			
		}
		Informational {
			$Loggingbox.SelectionColor = 'Blue'
			$Loggingbox.AppendText("`n $logmsg")
			
		}
		
	}
	
}

Function Create-Share
{
	
	<#
	.SYNOPSIS
		A function to create the requested DFS share.
	
	.DESCRIPTION
		This function takes all of the validated user input from the form and
        creates the required share if it does not already exist.
    
    .PARAMETER Share
        This is the user selected share from the dropdown box.

    .PARAMETER Region
        This is the user selected region from the dropdown box.

    .PARAMETER FolderName
        This is the combination of the client number and name.

    .PARAMETER Credential
        The Credential Object used to create the share.
	
	.EXAMPLE
		Create-Share -Share Admin -Region Americas -Foldername 10101_That_Client -Credential $Credential
	
	.NOTES
		N/A
#>
	
	Param
	(
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
		[String]$Share,
		[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline, ValueFromPipelineByPropertyName)]
		[String]$Region,
		[Parameter(Mandatory = $true, Position = 2, ValueFromPipeline, ValueFromPipelineByPropertyName)]
		[String]$ClientName,
		[Parameter(Mandatory = $false, Position = 3, ValueFromPipeline, ValueFromPipelineByPropertyName)]
		[String]$Mattername
	)
	
	
	If (($ClientName.length) -gt 20)
	{
		$ClientName = $ClientName.substring(0, 20)
	}
	
	$pattern = '[^a-zA-Z\d]'
	$ClientName = $ClientName -replace $Pattern, '_'
	$ClientFolder = "$($ClientNumbertextBox.text)" + '_' + $ClientName
	$ClientFolder = $ClientFolder.trimend("_")
	$ClientPath = "\\WCNET\Firm\$Share\$Region\$ClientFolder"
	Write-RichText -LogType Informational -LogMsg "The Client folder path is $ClientPath"
	
	If ($MatterName)
	{
		If (($Mattername.length) -gt 20)
		{
			$Mattername = $Mattername.substring(0, 20)
		}
		
		$Mattername = $Mattername -replace $Pattern, '_'
		$MatterFolder = "$($MatterNumbertextBox.text)" + '_' + $MatterName
		$MatterFolder = $MatterFolder.trimend("_")
		$Matterpath = "\\WCNET\Firm\$Share\$Region\$ClientFolder\$MatterFolder"
		Write-RichText -LogType Informational -LogMsg "The Matter folder path is $MatterPath"
	}
	
	If (Test-Path $ClientPath)
	{
		Write-RichText -LogType Informational -LogMsg "The client folder already exists."
	}
	
	Else
	{
		New-Item -ItemType Directory -Path "$ClientPath" -Force
		
		If (Test-Path $ClientPath)
		{
			Write-RichText -LogType Success -LogMsg "The client folder was created successfully."
		}
		
		Else
		{
			Write-RichText -LogType Error -LogMsg "`t Failed to create the client folder."
			Return
		}
	}
	
	If ($MatterName)
	{
		If (Test-Path $MatterPath)
		{
			Write-RichText -LogType Informational -LogMsg "The matter folder already exists."
		}
		
		Else
		{
			New-Item -ItemType Directory -Path $MatterPath -Force
			
			If (Test-Path $MatterPath)
			{
				Write-RichText -LogType Success -LogMsg "The matter folder was created successfully."
			}
			
			Else
			{
				Write-RichText -LogType Error -LogMsg "`t Failed to create the matter folder."
				Return
			}
		}
	}
	
	$Message = @"

User: $ENV:Username
Host: $env:COMPUTERNAME
Date: $(Get-Date)
Client: $ClientName
Client#: $($ClientNumbertextBox.text)
Matter: $MatterName
Matter#: $($MatterNumbertextBox.text)
ClientPath: $ClientPath
MatterPath: $Matterpath
"@
	
	Write-log -Message $Message
	Record-Stats -Body $Message -Subject "Share Creation Tool"
	
}

Function Write-Log
{
	<#
	.SYNOPSIS
		A function to write ouput messages to a logfile.
	
	.DESCRIPTION
		This function is designed to send timestamped messages to a logfile of your choosing.
		Use it to replace something like write-host for a more long term log.
	
	.PARAMETER Message
		The message being written to the log file.
	
	.EXAMPLE
		PS C:\> Write-Log -Message 'This is the message being written out to the log.' 
	
	.NOTES
		N/A
#>
	
	Param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[String]$Message
	)
	
	
	add-content -path $LogFilePath -value ($Message)
	Write-Output $Message
}

Function Query-ClientName
{
	<#
	.SYNOPSIS
		This function pulls the Client Name from the remote DB.
	
	.DESCRIPTION
		This function relies on the client number to find the client name from the 
		global replication db on a remote server.
	
	.PARAMETER Credential
		This credential must be for a user who has access to Invoke-Command to
		the remote sql server and must have rights to the sql DB itself.
	
	.PARAMETER Clnum
		This is the client number entered by the user.
	
	.EXAMPLE
		PS C:\> Query-MatterName -Credential $value1 -Clnum 'Value2' -Matternum 'Value3'
	
	.NOTES
		Will Return either $False (if the connection failed) or $Result.clname which will
		either be the client name or $null.
#>
	
	Param
	(
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
		$Credential,
		[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline, ValueFromPipelineByPropertyName)]
		[String]$Clnum
	)
	
	Write-Richtext -LogType Informational -LogMsg "Querying the database for the client name. Please be patient."
	$SQLStmt = "select  TOP 1 CLNAME from mhgroup.wc_ImpCM where CLNO = '$Clnum';"
	
	$Result = Invoke-Command -ComputerName 'AM1APDB720' `
							 -credential $Credential `
							 -ArgumentList $SQLStmt `
							 -ScriptBlock {
		Param
		($SQLStmt)
		
		Invoke-Sqlcmd -ServerInstance 'AM1APDB720\appdata' `
					  -Database 'globalreplication' `
					  -Query $SQLStmt
	}
	
	If (!$Result)
	{
		Return $False
	}
	
	Else
	{
		Return $Result.clname
	}
}

Function Query-MatterName
{
	<#
	.SYNOPSIS
		This function pulls the Matter Name from the remote DB.
	
	.DESCRIPTION
		This function relies on the matter number to find the matter name from the 
		global replication db on a remote server.
	
	.PARAMETER Credential
		This credential must be for a user who has access to Invoke-Command to
		the remote sql server and must have rights to the sql DB itself.
	
	.PARAMETER Clnum
		This is the client number entered by the user.
	
	.PARAMETER Matternum
		This is the matter number entered by the user.
	
	.EXAMPLE
		PS C:\> Query-MatterName -Credential $value1 -Clnum 'Value2' -Matternum 'Value3'
	
	.NOTES
		Will Return either $False (if the connection failed) or $Result.mtname which will
		either be the matter name or $null.
#>
	Param
	(
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
		$Credential,
		[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline, ValueFromPipelineByPropertyName)]
		[String]$Clnum,
		[Parameter(Mandatory = $true, Position = 2, ValueFromPipeline, ValueFromPipelineByPropertyName)]
		[String]$Matternum
	)
	
	Write-Richtext -LogType Informational -LogMsg "Querying the database for the matter name. Please be patient."
	$SQLStmt = "select TOP 1 MTNAME from mhgroup.wc_ImpCM where MTNO = '$Matternum' and CLNO = '$Clnum';"
	
	$Result = Invoke-Command -ComputerName 'AM1APDB720' `
							 -credential $Credential `
							 -ArgumentList $SQLStmt `
							 -ScriptBlock {
		Param
		($SQLStmt)
		
		Invoke-Sqlcmd -ServerInstance 'AM1APDB720\appdata' `
					  -Database 'globalreplication' `
					  -Query $SQLStmt
	}
	
	$Result.mtname
}

Function Record-Stats
{
	param
	(
		[parameter(Mandatory = $true)]
		[String]$Body,
		[parameter(Mandatory = $true)]
		[String]$Subject
	)
	
	$Splat = @{
		
		To = "phillip.marshall@whitecase.com"
		From = "ShareCreation@whitecase.com"
		Body = $Body
		Subject = $Subject
		SmtpServer = 'AM1SMTP'
		BodyAsHtml = $False
	}
	
	Send-MailMessage @Splat
}

#endregion

#region Variable Declarations

[Array]$ShareList = @("Admin", "Client")
[Array]$RegionList = @("Americas", "Emea", "Asiapac")
[String]$LogfilePath = "\\wcnet\firm\Applications\GTS\Software\Scripts\_Logs\Folder Creation Tool\Folder Creation Tool.txt"
[String]$DatabaseServer = 'AM1APDB720'
$SecPassword = ConvertTo-SecureString "Dont4get#56" -AsPlainText -Force
$GLOCredential = New-Object System.Management.Automation.PSCredential ('GLO-SA-Powershell', $SecPassword)
#endregion

#region DeclareForm

#################
#Create the Form#
#################
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(900, 700)
$Form.text = "White & Case Share Creation Tool"
$Form.BackColor = [System.Drawing.Color]::FromArgb(255, 185, 209, 234)

#endregion

#region ChooseShare

####################################
#Create the Dropdown - Choose Share#
####################################
$ShareDropDownBox = New-Object System.Windows.Forms.ComboBox
$ShareDropDownBox.Location = New-Object System.Drawing.Size(20, 50)
$ShareDropDownBox.Size = New-Object System.Drawing.Size(120, 25)
$ShareDropDownBox.DropDownHeight = 200
$ShareDropDownBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$Form.Controls.Add($ShareDropDownBox)

Foreach ($Share in $ShareList)
{
	$ShareDropDownBox.Items.Add($Share)
}

###################################
#Create the Label for Choose Share#
###################################
$ShareLabel = New-Object System.Windows.Forms.Label
$ShareLabel.Location = New-Object System.Drawing.Point(20, 30)
$ShareLabel.Size = New-Object System.Drawing.Size(150, 25)
$ShareLabel.Text = "Choose the Share"
$Form.Controls.Add($ShareLabel)

#endregion

#region ChooseRegion

#####################################
#Create the Dropdown - Choose Region#
#####################################
$RegionDropDownBox = New-Object System.Windows.Forms.ComboBox
$RegionDropDownBox.Location = New-Object System.Drawing.Size(175, 50)
$RegionDropDownBox.Size = New-Object System.Drawing.Size(120, 25)
$RegionDropDownBox.DropDownHeight = 200
$RegionDropDownBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$Form.Controls.Add($RegionDropDownBox)

Foreach ($Region in $RegionList)
{
	$RegionDropDownBox.Items.Add($Region)
}

####################################
#Create the Label for Choose Region#
####################################
$RegionLabel = New-Object System.Windows.Forms.Label
$RegionLabel.Location = New-Object System.Drawing.Point(175, 30)
$RegionLabel.Size = New-Object System.Drawing.Size(150, 25)
$RegionLabel.Text = "Choose the Region"
$Form.Controls.Add($RegionLabel)

#endregion

#region SubmitButton

##########################
#Create the Submit Button#
##########################
$SubmitButton = New-Object System.Windows.Forms.Button
$SubmitButton.Location = New-Object System.Drawing.Size(740, 30)
$SubmitButton.Size = New-Object System.Drawing.Size(130, 50)
$SubmitButton.Text = "Submit"
$SubmitButton.Enabled = $True
$SubmitButton.Add_Click({ Process-Choices })

$Form.Controls.Add($SubmitButton)

#endregion

#region ClearButton

#########################
#Create the Clear Button#
#########################
$ClearButton = New-Object System.Windows.Forms.Button
$ClearButton.Location = New-Object System.Drawing.Size(740, 90)
$ClearButton.Size = New-Object System.Drawing.Size(130, 50)
$ClearButton.Text = "Clear"
$ClearButton.Enabled = $True
$ClearButton.Add_Click({ Clear-Choices })

$Form.Controls.Add($ClearButton)

#endregion

#region ClientNumberInputBox

####################################
#Create the Client Number Input Box#
####################################
$ClientNumberTextBox = New-Object System.Windows.Forms.TextBox
$ClientNumberTextBox.Location = New-Object System.Drawing.Size(325, 50)
$ClientNumberTextBox.Size = New-Object System.Drawing.Size(250, 20)
$ClientNumberTextBox.BackColor = 'Gainsboro'
$Form.Controls.Add($ClientNumberTextBox)

##################################################
#Create the Label for the Client Number Input Box#
##################################################
$ClientNumberLabel = New-Object System.Windows.Forms.Label
$ClientNumberLabel.Location = New-Object System.Drawing.Point(325, 30)
$ClientNumberLabel.Size = New-Object System.Drawing.Size(250, 25)
$ClientNumberLabel.Text = "Type in the Client Number"
$Form.Controls.Add($ClientNumberLabel)

#endregion

#region MatternumberInputBox

####################################
#Create the Matter Number Input Box#
####################################
$MatterNumberTextBox = New-Object System.Windows.Forms.TextBox
$MatterNumberTextBox.Location = New-Object System.Drawing.Size(325, 110)
$MatterNumberTextBox.Size = New-Object System.Drawing.Size(250, 20)
$MatterNumberTextBox.BackColor = 'Gainsboro'
$Form.Controls.Add($MatterNumberTextBox)

################################################
#Create the Label for the Matter Number Input Box#
################################################
$MatterNumberLabel = New-Object System.Windows.Forms.Label
$MatterNumberLabel.Location = New-Object System.Drawing.Point(325, 90)
$MatterNumberLabel.Size = New-Object System.Drawing.Size(250, 25)
$MatterNumberLabel.Text = "Type in the Matter Number (Optional)"
$MatterNumberLabel.ForeColor = 'Black'
$Form.Controls.Add($MatterNumberLabel)

#endregion

#region LoggingBox

########################
#Create the Logging Box#
########################
$Loggingbox = New-Object 'System.Windows.Forms.RichTextBox'
$Loggingbox.Location = New-Object System.Drawing.Size(20, 150)
$Loggingbox.Size = New-Object System.Drawing.Point(850, 500)
$Loggingbox.Name = "Output Box"
$Loggingbox.TabIndex = 2
$Loggingbox.Text = ""
$Loggingbox.font = "Arial"
$Loggingbox.BackColor = 'Gainsboro'
$Form.Controls.Add($Loggingbox)

#endregion

$Form.Add_Shown({ $Form.Activate() })

#region ADM Check

if ($env:USERNAME -notlike '*-adm*')
{
	Write-Richtext -LogType Error -LogMsg "Tool was not launched as Admin! Relaunch with your -ADM account!"
}

Else
{
	Write-richtext -LogType Informational -LogMsg "Welcome $Env:username. Please fill out the appropriate fields and hit submit."
	Write-richtext -LogType Informational -LogMsg "All usage of this tool is loggged."
}

#endregion

[void]$Form.ShowDialog()