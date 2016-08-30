Function Show-Diff {

    param([String[]]$OldString,[String[]]$NewString)

    $WebPage = @"
<!DOCTYPE HTML>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <!-- when using the mode "code", it's important to specify charset utf-8 -->
  <meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<script type="text/javascript"
 src="http://ajax.googleapis.com/ajax/libs/jquery/1.9.0/jquery.min.js"></script>

<script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.17.0/codemirror.js"></script>
<link rel="stylesheet" type="text/css" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.17.0/codemirror.css"/>

<script type="text/javascript" src="https://cdn.rawgit.com/wickedest/Mergely/master/lib/mergely.min.js"></script>
<link type="text/css" rel="stylesheet" href="https://cdn.rawgit.com/wickedest/Mergely/master/lib/mergely.css" />

<script id="OldString" type="multistring">$($OldString -join "`n")</script>
<script id="NewString" type="multistring">$($NewString -join "`n")</script>
</head>
<body>

<div id="mergely-resizer">
		<div id="compare">
		</div>
	</div>

<script>
`$(document).ready(function () {
	`$('#compare').mergely({
        editor_width: "400px",
		cmsettings: { readOnly: false, lineNumbers: true },
		lhs: function(setValue) {
			setValue(document.getElementById('OldString').innerHTML);
		},
		rhs: function(setValue) {
			setValue(document.getElementById('NewString').innerHTML);
		}
	});
});
</script>

</body>
</html>
"@

Write-Host $WebPage

 [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = @'
    <Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PowerShell HTML GUI" WindowStartupLocation="CenterScreen">
            <WebBrowser Name="WebBrowser"></WebBrowser>
    </Window>
'@

    #Read XAML
    $reader=(New-Object System.Xml.XmlNodeReader $xaml) 
    $Form=[Windows.Markup.XamlReader]::Load( $reader )
    #===========================================================================
    # Store Form Objects In PowerShell
    #===========================================================================
    $WebBrowser = $Form.FindName("WebBrowser")

    $WebBrowser.NavigateToString($WebPage)

    $Form.ShowDialog()

}