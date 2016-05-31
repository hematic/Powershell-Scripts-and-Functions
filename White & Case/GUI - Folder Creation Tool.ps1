#Function Declarations
#######################################################
Function Process-Choices
{

	<#
	.SYNOPSIS
		A function to process the data the user put into the form.
	
	.DESCRIPTION
		This function validates data that the user inputted into the form to
        verify it is acceptable for the folder creation function.
	
	.EXAMPLE
		Process-Choices
	
	.NOTES
		N/A
#>

    $Invalid = $False
    $Loggingbox.clear()

    ################
    #Validate Share#
    ################
    
    #If the share dropdown is null this will output the problem to the Rich Text Box.
    Try
    {
        $ChosenShare = $ShareDropDownBox.SelectedItem.ToString()
    }

    Catch
    {
        Write-RichText -LogType Error -LogMsg "`t No Share has been chosen from the dropdown."
        $Invalid =$True    
    }

    #################
    #Validate Region#
    #################

    #If the region dropdown is null this will output the problem to the Rich Text Box.
    Try
    {
        $ChosenRegion = $RegionDropDownBox.SelectedItem.ToString()
    }

    Catch
    {
        Write-RichText -LogType Error -LogMsg "`t No Region has been chosen from the Dropdown."
        $Invalid =$True    
    }

    ########################
    #Validate Client Number#
    ########################

    <#
    If the Client Number Text Box is null or empty this will output the problem 
    to the Rich Text Box. Client Numbers are allowed to have hyphens so if one is
    present we temporarily strip it out to validate the rest of the string is a 
    valid INT. If either of these processes fail we output the problem to the Rich
    Text Box.
    #>

    $ChosenClientNumber = $ClientNumberTextBox.text
    IF([string]::IsNullOrWhiteSpace($ChosenClientNumber))
    {    
        Write-RichText -LogType Error -LogMsg "`t No Client Number was provided."
        $Invalid =$True  
    }

    $Intcheck = $ChosenClientNumber -replace "-",""
    If(($Intcheck -as [Int]) -isnot [Int])
    {
        Write-RichText -LogType Error -LogMsg "`t The Client Number given is not all numeric."
        $Invalid =$True         
    }

    ######################
    #Validate Client Name#
    ######################

    <#
    If the Client Name Text Box is null or empty this will output the problem
    to the Rich Text Box. We also trim the entry to 20 characters as requested 
    and remove any special characters and replace them with underscores. 
    However this is only temporary as we should be pulling this from DMS directly 
    in the near future.
    #>

    $ChosenClientName = $ClientNameTextBox.text
    IF([string]::IsNullOrWhiteSpace($ChosenClientName))
    {    
        Write-RichText -LogType Error -LogMsg "`t No Client Name was provided."
        $Invalid =$True 
    }

    Else
    {
        If(($ChosenClientName.length) -gt 20)
        {
            $ChosenClientName = $ChosenClientName.substring(0,20)
        }
        

        <#
        Handle all possible special characters that should be replaced.
        This regex replaces any character that is not a letter or a number with an underscore.
        #>

        $pattern = '[^a-zA-Z\d]'
        $ChosenClientName =  $ChosenClientName -replace $Pattern, '_' 

    }

    If($Invalid -eq $True)
    {
        Return
    }

    Else
    {
        Write-RichText -LogType Informational -LogMsg "`n `t Chosen Share `t : $ChosenShare `n `t Chosen Region `t : $ChosenRegion `n `t ClientNumber `t : $ChosenClientnumber `n `t ClientName `t : $ChosenClientname"
        Write-RichText -LogType Informational -LogMsg "`n `t Prompting for credentials..."

        $Foldername = $ChosenClientNumber + '_' + $ChosenClientName
        $Creds = Get-Credential

        Create-share -Share $ChosenShare -Region $ChosenRegion -FolderName $Foldername -Credential $Creds
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
    $ClientNameTextBox.Text = ''
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
	
	.EXAMPLE
		Write-Richtext -LogType Error -LogMsg "This is an Error."
		Write-Richtext -LogType Success -LogMsg "This is a Success."
		Write-Richtext -LogType Informational -LogMsg "This is Informational."
	
	.NOTES
		N/A
#>

    Param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String]$LogType, 
    	[Parameter(Mandatory = $true,Position = 1,ValueFromPipeline,ValueFromPipelineByPropertyName)]
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
                $Loggingbox.AppendText("`n $logtype : $logmsg")
                 
            }
            Informational {
                $Loggingbox.SelectionColor = 'Blue'
                $Loggingbox.AppendText("`n $logtype : $logmsg")
                 
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
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String]$Share, 
    	[Parameter(Mandatory = $true,Position = 1,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String]$Region,
    	[Parameter(Mandatory = $true,Position = 2,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String]$FolderName,
    	[Parameter(Mandatory = $true,Position = 3,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[Object]$Credential
    )
    
    $CompletePath = "\\WCNET\Firm\$Share\$Region\$FolderName"
    
    
    If(Test-Path $CompletePath)
    {
        Write-RichText -LogType Error -LogMsg "`n `t The specified path ($CompletePath) already exists."
        Return
    }

    New-Item -ItemType Directory -Path "$CompletePath" -Credential $Credential -Force

    If(Test-Path $CompletePath)
    {
        Write-RichText -LogType Success -LogMsg "`n `t The specified path was successfully created. `n `t ($CompletePath)"

        $message = @"
******************************
User        : $Env:username
Share       : $Share
Region      : $Region
Client #    : $($ClientNumberTextBox.text)
Client Name : $($ChosenClientName.substring(0,20))
"@

        Write-Log -Message $message

        #Clear the selections
        $ShareDropDownBox.SelectedIndex = -1
        $ShareDropDownBox.SelectedIndex = -1
        $RegionDropDownBox.SelectedIndex = -1
        $RegionDropDownBox.SelectedIndex = -1
        $ClientNumberTextBox.Text = ''
        $ClientNameTextBox.Text = ''

        Return
    }

    Else
    {
        Write-RichText -LogType Error -LogMsg "`n `t Creation of the specified path ($CompletePath) failed."
        Write-RichText -LogType Error -LogMsg "`n `t Reason : $?"
        Return    
    }
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

#######################
#Variable Declarations#
#######################

[Array]$ShareList = @("Admin","Client")
[Array]$RegionList = @("Americas","Emea","Asiapac")
[String]$LogfilePath = "\\wcnet\firm\Applications\GTS\Software\Scripts\_Logs\Folder Creation Tool.txt"

#region DeclareForm

#################
#Create the Form#
#################
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  
$Form = New-Object System.Windows.Forms.Form    
$Form.Size = New-Object System.Drawing.Size(800,500)
$Form.text = "White & Case Share Creation Tool"
$Form.backcolor = "gray"

#endregion

#region ChooseShare

####################################
#Create the Dropdown - Choose Share#
####################################
$ShareDropDownBox = New-Object System.Windows.Forms.ComboBox 
$ShareDropDownBox.Location = New-Object System.Drawing.Size(20,50)
$ShareDropDownBox.Size = New-Object System.Drawing.Size(120,25)
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
$ShareLabel.Location = New-Object System.Drawing.Point(20,30)
$ShareLabel.Size = New-Object System.Drawing.Size(150, 25)
$ShareLabel.Text = "Choose the Share"
$Form.Controls.Add($ShareLabel)

#endregion

#region ChooseRegion

#####################################
#Create the Dropdown - Choose Region#
#####################################
$RegionDropDownBox = New-Object System.Windows.Forms.ComboBox 
$RegionDropDownBox.Location = New-Object System.Drawing.Size(175,50)       
$RegionDropDownBox.Size = New-Object System.Drawing.Size(120,25)            
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
$ShareLabel = New-Object System.Windows.Forms.Label
$ShareLabel.Location = New-Object System.Drawing.Point(175,30)
$ShareLabel.Size = New-Object System.Drawing.Size(150, 25)
$ShareLabel.Text = "Choose the Region"
$Form.Controls.Add($ShareLabel)

#endregion

#region SubmitButton

##########################
#Create the Submit Button#
##########################
$SubmitButton = New-Object System.Windows.Forms.Button 
$SubmitButton.Location = New-Object System.Drawing.Size(640,30) 
$SubmitButton.Size = New-Object System.Drawing.Size(130,50) 
$SubmitButton.Text = "Submit"
$SubmitButton.Enabled = $True
$SubmitButton.Add_Click({Process-Choices})

$Form.Controls.Add($SubmitButton)

#endregion

#region ClearButton

#########################
#Create the Clear Button#
#########################
$ClearButton = New-Object System.Windows.Forms.Button 
$ClearButton.Location = New-Object System.Drawing.Size(640,90) 
$ClearButton.Size = New-Object System.Drawing.Size(130,50) 
$ClearButton.Text = "Clear"
$ClearButton.Enabled = $True
$ClearButton.Add_Click({Clear-Choices})

$Form.Controls.Add($ClearButton)

#endregion

#region ClientNumberInputBox

####################################
#Create the Client Number Input Box#
####################################
$ClientNumberTextBox = New-Object System.Windows.Forms.TextBox 
$ClientNumberTextBox.Location = New-Object System.Drawing.Size(325,50) 
$ClientNumberTextBox.Size = New-Object System.Drawing.Size(260,20) 
$ClientNumberTextBox.BackColor = "White"
$Form.Controls.Add($ClientNumberTextBox) 

##################################################
#Create the Label for the Client Number Input Box#
##################################################
$ShareLabel = New-Object System.Windows.Forms.Label
$ShareLabel.Location = New-Object System.Drawing.Point(325,30)
$ShareLabel.Size = New-Object System.Drawing.Size(260, 25)
$ShareLabel.Text = "Type in the Client Number"
$Form.Controls.Add($ShareLabel)

#endregion

#region ClientNameInputBox

##################################
#Create the Client Name Input Box#
##################################
$ClientNameTextBox = New-Object System.Windows.Forms.TextBox 
$ClientNameTextBox.Location = New-Object System.Drawing.Size(325,110) 
$ClientNameTextBox.Size = New-Object System.Drawing.Size(260,20)
$ClientNameTextBox.BackColor = "White"
$Form.Controls.Add($ClientNameTextBox) 

################################################
#Create the Label for the Client Name Input Box#
################################################
$ShareLabel = New-Object System.Windows.Forms.Label
$ShareLabel.Location = New-Object System.Drawing.Point(325,90)
$ShareLabel.Size = New-Object System.Drawing.Size(260, 25)
$ShareLabel.Text = "Type in the Client Name"
$ShareLabel.ForeColor = 'Black'
$Form.Controls.Add($ShareLabel)

#endregion

#region LoggingBox

########################
#Create the Logging Box#
########################
$Loggingbox = New-Object 'System.Windows.Forms.RichTextBox'
$Loggingbox.Location = New-Object System.Drawing.Size(20,150) 
$Loggingbox.Size = New-Object System.Drawing.Point(750,300) 
$Loggingbox.Name = "Output Box"
$Loggingbox.TabIndex = 2
$Loggingbox.Text = ""
$Loggingbox.font = "Arial"
$Loggingbox.BackColor = "White"
$Form.Controls.Add($Loggingbox)

#endregion

$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
