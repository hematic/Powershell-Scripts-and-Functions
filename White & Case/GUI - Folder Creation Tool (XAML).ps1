Function Get-FormVariables
{
    write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
    get-variable WPF*
}

Function Process-Choices
{
    $Invalid = $False
    $WPFrichTextBox.clear()

    ################
    #Validate Share#
    ################

    Try
    {
        $ChosenShare = $WPFShare_Dropdown.SelectedItem.ToString()
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
        $ChosenRegion = $WPFRegion_Dropdown.SelectedItem.ToString()
    }

    Catch
    {
        Write-RichText -LogType Error -LogMsg "`t No Region has been chosen from the Dropdown."
        $Invalid =$True    
    }

    ########################
    #Validate Client Number#
    ########################

    $ChosenClientNumber = $WPFClient_Number_TextBox.text
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
                $WPFrichTextBox.SelectionColor = 'Red'
                $WPFrichTextBox.AppendText("`n $logtype : $logmsg")
                 
            }
            Success {
                $WPFrichTextBox.SelectionColor = 'Green'
                $WPFrichTextBox.AppendText("`n $logtype : $logmsg")
                 
            }
            Informational {
                $WPFrichTextBox.SelectionColor = 'Blue'
                $WPFrichTextBox.AppendText("`n $logtype : $logmsg")
                 
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
    
    $CompletePath = "\\WCNET\Firm\$Share\$Region\$FolderName"
    #$CompletePath = "C:\windows\temp\$Share\$Region\$FolderName"
    
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
Client #    : $($WPFClient_Number_TextBox.text)
Client Name : $($ChosenClientName.substring(0,20))
"@

        Write-Log -Message $message

        #Clear the selections
        $WPFShare_Dropdown.SelectedIndex = -1
        $WPFShare_Dropdown.SelectedIndex = -1
        $WPFRegion_Dropdown.SelectedIndex = -1
        $WPFRegion_Dropdown.SelectedIndex = -1
        $WPFClient_Number_TextBox.Text = ''
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

#region GUI XAML is in here
$inputXML = @"
<Window x:Class="WpfApplication1.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApplication1"
        mc:Ignorable="d"
        Title="White &amp; Case Share Creation Tool" Height="410.553" Width="650" Background="#FF7A7878" ResizeMode="CanResizeWithGrip">
    <Grid Height="322">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="37*"/>
            <ColumnDefinition Width="22*"/>
            <ColumnDefinition Width="25*"/>
            <ColumnDefinition Width="401*"/>
            <ColumnDefinition Width="157*"/>
        </Grid.ColumnDefinitions>
        <Image x:Name="image" HorizontalAlignment="Left" Height="100" Margin="332.783,0,0,0" VerticalAlignment="Top" Width="100" Grid.ColumnSpan="2" Grid.Column="3"/>
        <Image x:Name="WC_Logo" HorizontalAlignment="Left" Height="101" Margin="57,-27,0,0" VerticalAlignment="Top" Width="100" Source="C:\Users\Phillip\Downloads\WandC.png" Grid.Column="4"/>
        <ComboBox x:Name="Share_Dropdown" HorizontalAlignment="Left" Margin="0,-7,0,0" VerticalAlignment="Top" Width="120" IsSynchronizedWithCurrentItem="True" Grid.ColumnSpan="4">
            <ComboBox.Background>
                <LinearGradientBrush EndPoint="0,1" StartPoint="0,0">
                    <GradientStop Color="#FFF0F0F0" Offset="0"/>
                    <GradientStop Color="#FF7A7878" Offset="1"/>
                </LinearGradientBrush>
            </ComboBox.Background>
        </ComboBox>
        <ComboBox x:Name="Region_Dropdown" HorizontalAlignment="Left" Margin="0,57,0,0" VerticalAlignment="Top" Width="120" Grid.ColumnSpan="4"/>
        <Label x:Name="Share_Dropdown_Label" Content="Choose the Share" HorizontalAlignment="Left" VerticalAlignment="Top" Width="120" Margin="0,-33,0,0" Grid.ColumnSpan="4"/>
        <Label x:Name="Region_Dropdown_Label" Content="Choose the Region" HorizontalAlignment="Left" Margin="0,31,0,0" VerticalAlignment="Top" Width="120" Grid.ColumnSpan="4"/>
        <TextBox x:Name="Client_Number_TextBox" HorizontalAlignment="Left" Height="23" Margin="0,121,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="200" Grid.ColumnSpan="4"/>
        <Label x:Name="Client_Number_Label" Content="Type in the Client Number" HorizontalAlignment="Left" Margin="0,95,0,0" VerticalAlignment="Top" Width="200" Height="26" Grid.ColumnSpan="4"/>
        <RichTextBox x:Name="richTextBox" HorizontalAlignment="Left" Height="245" Margin="167,99,0,-22" VerticalAlignment="Top" Width="381" Grid.ColumnSpan="2" Grid.Column="3">
            <FlowDocument>
                <Paragraph>
                    <Run Text="RichTextBox"/>
                </Paragraph>
            </FlowDocument>
        </RichTextBox>
        <Button x:Name="Submit_Button" Content="Submit" HorizontalAlignment="Left" Margin="0,245,0,0" VerticalAlignment="Top" Height="67" ToolTip="Once all 4 required fields are filled out click here." Width="93" FontFamily="Times New Roman" FontSize="24" FontWeight="Bold" Grid.ColumnSpan="4"/>
        <Label x:Name="label" Content="Type in the Matter Number" HorizontalAlignment="Left" Margin="0,162,0,0" VerticalAlignment="Top" Width="200" Grid.ColumnSpan="4"/>
        <TextBox x:Name="textBox" HorizontalAlignment="Left" Height="23" Margin="0,188,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="200" Grid.ColumnSpan="4"/>

    </Grid>
</Window>
"@       
#endregion 

#This tweaks the XAML to work with PowerShell properly by taking it back to XML.
$inputXML = $inputXML   -replace 'mc:Ignorable="d"',''`
                        -replace "x:N",'N'`
                        -replace '^<Win.*', '<Window'
 
 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML

#Read the XML and error if .net is missing.
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try
{
    $Form=[Windows.Markup.XamlReader]::Load($reader)
}
catch
{
    Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
}

[Array]$ShareList = @("Admin","Client")
[Array]$RegionList = @("Americas","Emea","Asiapac")
[String]$LogfilePath = "\\wcnet\firm\Applications\GTS\Software\Scripts\_Logs\Folder Creation Tool.txt"

Foreach ($Share in $ShareList) 
{
    $WPFShare_Dropdown.Items.Add($Share)
}

Foreach ($Region in $RegionList) 
{
    $WPFRegion_Dropdown.Items.Add($Region)
}
 
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}

#===========================================================================
# Actually make the objects work
#===========================================================================
 
$WPFSubmitButton.Add_Click({Process-Choices})

#===========================================================================
# Shows the form
#===========================================================================

$Form.ShowDialog() | out-null
