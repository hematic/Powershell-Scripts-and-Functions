#Function Declarations
#######################################################
Function Process-Choices
{
    $Invalid = $False
    $Loggingbox.clear()

    ################
    #Validate Share#
    ################

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

    $ChosenClientNumber = $ClientNumberTextBox.text
    IF([string]::IsNullOrWhiteSpace($ChosenClientNumber))
    {    
        Write-RichText -LogType Error -LogMsg "`t No Client Number was provided."
        $Invalid =$True  
    }

    If(($ChosenClientNumber -as [Int]) -isnot [Int])
    {
        Write-RichText -LogType Error -LogMsg "`t The Client Number given is not all numeric."
        $Invalid =$True         
    }

    ######################
    #Validate Client Name#
    ######################

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
        
        #################################################################
        #Handle all possible special characters that should be replaced.#
        #################################################################

        $ChosenClientName = $ChosenClientName.replace("!","_")
        $ChosenClientName = $ChosenClientName.replace("@","_")
        $ChosenClientName = $ChosenClientName.replace("#","_")
        $ChosenClientName = $ChosenClientName.replace("$","_")
        $ChosenClientName = $ChosenClientName.replace("%","_")
        $ChosenClientName = $ChosenClientName.replace("^","_")
        $ChosenClientName = $ChosenClientName.replace("&","_")
        $ChosenClientName = $ChosenClientName.replace("*","_")
        $ChosenClientName = $ChosenClientName.replace("(","_")
        $ChosenClientName = $ChosenClientName.replace(")","_")
        $ChosenClientName = $ChosenClientName.replace("{","_")
        $ChosenClientName = $ChosenClientName.replace("}","_")
        $ChosenClientName = $ChosenClientName.replace("[","_")
        $ChosenClientName = $ChosenClientName.replace("]","_")
        $ChosenClientName = $ChosenClientName.replace("`"","_")
        $ChosenClientName = $ChosenClientName.replace("`'","_")
        $ChosenClientName = $ChosenClientName.replace(":","_")
        $ChosenClientName = $ChosenClientName.replace(";","_")
        $ChosenClientName = $ChosenClientName.replace(",","_")
        $ChosenClientName = $ChosenClientName.replace(".","_")
        $ChosenClientName = $ChosenClientName.replace("<","_")
        $ChosenClientName = $ChosenClientName.replace(">","_")
        $ChosenClientName = $ChosenClientName.replace("/","_")
        $ChosenClientName = $ChosenClientName.replace("|","_")
        $ChosenClientName = $ChosenClientName.replace("\","_")
        $ChosenClientName = $ChosenClientName.replace("-","_")
        $ChosenClientName = $ChosenClientName.replace("+","_")
        $ChosenClientName = $ChosenClientName.replace("=","_")
        $ChosenClientName = $ChosenClientName.replace("``","_")
        $ChosenClientName = $ChosenClientName.replace("~","_")
        $ChosenClientName = $ChosenClientName.replace(" ","_")

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

function Write-RichText
{

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

function Create-Share
{
    Param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String]$Share, 
    	[Parameter(Mandatory = $true,Position = 1,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String]$Region,
    	[Parameter(Mandatory = $true,Position = 1,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String]$FolderName,
    	[Parameter(Mandatory = $true,Position = 1,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[Object]$Credential
    )
    
    #$CompletePath = "\\WCNET\Firm\$Share\$Region\$FolderName"
    $CompletePath = "C:\windows\temp\$Share\$Region\$FolderName"
    
    If(Test-Path $CompletePath)
    {
        Write-RichText -LogType Error -LogMsg "`n `t The specified path ($CompletePath) already exists."
        Return
    }

    New-Item -ItemType Directory -Path "$CompletePath" -Credential $Credential -Force

    If(Test-Path $CompletePath)
    {
        Write-RichText -LogType Success -LogMsg "`n `t The specified path was successfully created. `n `t ($CompletePath)"

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

#######################
#Variable Declarations#
#######################

[Array]$ShareList = @("Admin","Client")
[Array]$RegionList = @("Americas","Emea","Asiapac")

#region DeclareForm

#################
#Create the Form#
#################
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  
$Form = New-Object System.Windows.Forms.Form    
$Form.Size = New-Object System.Drawing.Size(800,500)
$Form.text = "White & Case Share Creation Tool"

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
$RegionDropDownBox.Location = New-Object System.Drawing.Size(20,100)       
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
$ShareLabel.Location = New-Object System.Drawing.Point(20, 80)
$ShareLabel.Size = New-Object System.Drawing.Size(150, 25)
$ShareLabel.Text = "Choose the Region"
$Form.Controls.Add($ShareLabel)

#endregion

#region SubmitButton

##########################
#Create the Submit Button#
##########################
$SubmitButton = New-Object System.Windows.Forms.Button 
$SubmitButton.Location = New-Object System.Drawing.Size(600,30) 
$SubmitButton.Size = New-Object System.Drawing.Size(130,100) 
$SubmitButton.Text = "Submit"
$SubmitButton.Enabled = $True
$SubmitButton.Add_Click({Process-Choices})

$Form.Controls.Add($SubmitButton)

#endregion

#region ClientNumberInputBox
####################################
#Create the Client Number Input Box#
####################################
$ClientNumberTextBox = New-Object System.Windows.Forms.TextBox 
$ClientNumberTextBox.Location = New-Object System.Drawing.Size(20,150) 
$ClientNumberTextBox.Size = New-Object System.Drawing.Size(260,20) 
$Form.Controls.Add($ClientNumberTextBox) 

##################################################
#Create the Label for the Client Number Input Box#
##################################################
$ShareLabel = New-Object System.Windows.Forms.Label
$ShareLabel.Location = New-Object System.Drawing.Point(20,130)
$ShareLabel.Size = New-Object System.Drawing.Size(260, 25)
$ShareLabel.Text = "Type in the Client Matter Number"
$Form.Controls.Add($ShareLabel)

#endregion

#region ClientNameInputBox

##################################
#Create the Client Name Input Box#
##################################
$ClientNameTextBox = New-Object System.Windows.Forms.TextBox 
$ClientNameTextBox.Location = New-Object System.Drawing.Size(20,200) 
$ClientNameTextBox.Size = New-Object System.Drawing.Size(260,20) 
$Form.Controls.Add($ClientNameTextBox) 

################################################
#Create the Label for the Client Name Input Box#
################################################
$ShareLabel = New-Object System.Windows.Forms.Label
$ShareLabel.Location = New-Object System.Drawing.Point(20,180)
$ShareLabel.Size = New-Object System.Drawing.Size(260, 25)
$ShareLabel.Text = "Type in the Client Matter Name"
$Form.Controls.Add($ShareLabel)

#endregion

#region LoggingBox

$Loggingbox = New-Object 'System.Windows.Forms.RichTextBox'
$Loggingbox.Location = New-Object System.Drawing.Size(300,150) 
$Loggingbox.Size = New-Object System.Drawing.Point(475,300) 
$Loggingbox.Name = "Output Box"
$Loggingbox.TabIndex = 2
$Loggingbox.Text = ""
$Loggingbox.font = "Arial"
$Form.Controls.Add($Loggingbox)

#endregion

$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
