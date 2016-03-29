# Load the Winforms assembly
[reflection.assembly]::LoadWithPartialName( "System.Windows.Forms")
[reflection.assembly]::loadwithpartialname("System.Drawing")

# Close function
function do_exit
{
     $form.close()
}

# Create form, size, position and border style
$form= New-Object Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(400,400) 
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = 'Fixed3D'
$form.Text = " A Powershell GUI by Juan Hernandez"

## Embed logo
$file = (get-item 'C:\tmp\logo.png')
$img = [System.Drawing.Image]::Fromfile($file);
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Width =  $img.Size.Width;
$pictureBox.Height =  $img.Size.Height;
$pictureBox.Image = $img;
$pictureBox.Location = New-Object Drawing.Point 120,5
$form.controls.add($pictureBox)

## Create a close button, add text, set location, action and add it to the form
$B_close = New-Object Windows.Forms.Button
$B_close.text = "Close"
$B_close.Location = New-Object Drawing.Point 305,330
$B_close.add_click({do_exit})
$form.controls.add($B_close)

## create a button, tekst, position, action and add it to the form
$B_ont = New-Object Windows.Forms.Button
$B_ont.text = "Submit"
$B_ont.Location = New-Object Drawing.Point 20,210
$B_ont.add_click({$label.Text = $InputTextBox.text;$label.refresh})
$form.controls.add($B_ont)

# Create a label with text, size, location and add it to the form
$label = New-Object Windows.Forms.Label
$label.Location = New-Object Drawing.Point 230,183
$label.Size = New-Object Drawing.Point 150,20
$label.text = "Submit your text..."
$form.controls.add($label)

##create an input box
$InputTextBox = New-Object System.Windows.Forms.TextBox 
$InputTextBox.Location = New-Object System.Drawing.Size(20,180) 
$InputTextBox.Size = New-Object System.Drawing.Size(200,20) 
$form.Controls.Add($InputTextBox) 

## create a button, tekst, position, action and add it to the form
$B_cho = New-Object Windows.Forms.Button
$B_cho.text = "Choose"
$B_cho.Location = New-Object Drawing.Point 20,240
$B_cho.add_click({Return-DropDown})
$form.controls.add($B_cho)

# Create a label with text, size, location and add it to the form
$label2 = New-Object Windows.Forms.Label
$label2.Location = New-Object Drawing.Point 230,245
$label2.Size = New-Object Drawing.Point 150,20
$label2.text = "Make a choice..."
$form.controls.add($label2)

## function to retunr dropdown choice
function Return-DropDown {

	$Choice = $DropDown.SelectedItem.ToString();
	$label2.Text = $Choice;
	$label2.Refresh;
}

## create dropdown button
[array]$DropDownArray = "First", "Second", "Third"
$DropDown = new-object System.Windows.Forms.ComboBox
$DropDown.Location = new-object System.Drawing.Size(100,240)
$DropDown.Size = new-object System.Drawing.Size(80,30)

ForEach ($Item in $DropDownArray) {
	$DropDown.Items.Add($Item)
}

$form.Controls.Add($DropDown)

## Create application menu
function addMenuItem 
{ 
param([ref]$ParentItem, 
[string]$ItemName='', 
[string]$ItemText='', 
[scriptblock]$ScriptBlock=$null 
) 
[System.Windows.Forms.ToolStripMenuItem]$private:menuItem=`
New-Object System.Windows.Forms.ToolStripMenuItem; 
$private:menuItem.Name =$ItemName; 
$private:menuItem.Text =$ItemText; 
if ($ScriptBlock -ne $null) 
{ $private:menuItem.add_Click(([System.EventHandler]$handler=`
$ScriptBlock)); } 
if (($ParentItem.Value) -is [System.Windows.Forms.MenuStrip]) 
{ ($ParentItem.Value).Items.Add($private:menuItem); } 
if (($ParentItem.Value) -is [System.Windows.Forms.ToolStripItem]) 
{ ($ParentItem.Value).DropDownItems.Add($private:menuItem); } 
return $private:menuItem; 
}

[System.Windows.Forms.MenuStrip]$mainMenu=New-Object System.Windows.Forms.MenuStrip; 
$form.Controls.Add($mainMenu);

[scriptblock]$sb1= {Write-Host "You want to open something";}
[scriptblock]$sb2= {do_exit;}

#region File 
(addMenuItem -ParentItem ([ref]$mainMenu) -ItemName 'mnuFile' -ItemText 'File' -ScriptBlock $null) | %{ 
$null=addMenuItem -ParentItem ([ref]$_) -ItemName 'mnuFileOpen' -ItemText 'Open' -ScriptBlock $sb1; 
$null=addMenuItem -ParentItem ([ref]$_) -ItemName 'mnuFileSave' -ItemText 'Save' -ScriptBlock $null; 
$null=addMenuItem -ParentItem ([ref]$_) -ItemName 'mnuFileExit' -ItemText 'Exit' -ScriptBlock $sb2;} | Out-Null; 
#endregion File

## show the form
$form.ShowDialog()