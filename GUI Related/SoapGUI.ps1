## SOAP test GUI

# Load the Winforms assembly
[reflection.assembly]::LoadWithPartialName( "System.Windows.Forms")
[reflection.assembly]::loadwithpartialname("System.Drawing")

# Close function
function do_exit
{
     $form.close()
}

function Execute-SOAPRequest 
( 
        [Xml]    $SOAPRequest, 
        [String] $URL 
) 
{ 
        write-host "Sending SOAP Request To Server: $URL" 
        $soapWebRequest = [System.Net.WebRequest]::Create($URL) 
      	$soapWebRequest.Headers.Add("SOAPAction","`"http://www.webserviceX.NET/ChangeMetricWeightUnit`"")
        $soapWebRequest.ContentType = "application/soap+xml;charset=`"utf-8`""
		$soapWebRequest.Accept      = "text/xml" 
        $soapWebRequest.Method      = "POST" 
		
        write-host "Initiating Send." 
        $requestStream = $soapWebRequest.GetRequestStream() 
        $SOAPRequest.Save($requestStream) 
        $requestStream.Close() 
        
        write-host "Send Complete, Waiting For Response." 
        $resp = $soapWebRequest.GetResponse() 
        $responseStream = $resp.GetResponseStream() 
        $soapReader = [System.IO.StreamReader]($responseStream) 
        $ReturnXml = [Xml] $soapReader.ReadToEnd() 
        $responseStream.Close() 
        write-host "Response Received."
        return $ReturnXml 

}

# Create form, size, position and border style
$form= New-Object Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(400,400) 
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = 'Fixed3D'
$form.Text = " A Powershell GUI by Phillip Marshall"

## Create a close button, add text, set location, action and add it to the form
$B_close = New-Object Windows.Forms.Button
$B_close.text = "Close"
$B_close.Location = New-Object Drawing.Point 305,330
$B_close.add_click({do_exit})
$form.controls.add($B_close)

# Create a label with text, size, location and add it to the form
$label = New-Object Windows.Forms.Label
$label.Location = New-Object Drawing.Point 20,20
$label.Size = New-Object Drawing.Point 150,20
$label.text = "Enter endpoint:"
$form.controls.add($label)

##create an input box
$InputTextBox = New-Object System.Windows.Forms.TextBox 
$InputTextBox.Location = New-Object System.Drawing.Size(20,40) 
$InputTextBox.Size = New-Object System.Drawing.Size(350,20)

$form.Controls.Add($InputTextBox) 

# Create a label with text, size, location and add it to the form
$label2 = New-Object Windows.Forms.Label
$label2.Location = New-Object Drawing.Point 20,70
$label2.Size = New-Object Drawing.Point 150,20
$label2.text = "Enter SOAP body:"
$form.controls.add($label2)

##create an input box
$InputTextBox2 = New-Object System.Windows.Forms.TextBox 
$InputTextBox2.Location = New-Object System.Drawing.Size(20,90) 
$InputTextBox2.Size = New-Object System.Drawing.Size(350,200)
$InputTextBox2.Multiline = "true";

$form.Controls.Add($InputTextBox2) 

## create a button, tekst, position, action and add it to the form
$B_sub = New-Object Windows.Forms.Button
$B_sub.text = "Submit Test"
$B_sub.Location = New-Object Drawing.Point 20,330
$B_sub.add_click(
{

$soap = [xml] $InputTextBox2.Text;
$url = $InputTextBox.Text;
$ret = Execute-SOAPRequest $soap $url;
$ret | Export-Clixml  "c:\tmp\response.xml";
}
)
$form.controls.add($B_sub)

## show the form
$form.ShowDialog()