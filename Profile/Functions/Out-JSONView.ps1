  Function Out-JSONView  {

  [cmdletbinding()]

  Param (

  [parameter(ValueFromPipeline)]

  [psobject]$InbutObject,
  [int]$Depth

  )

  $JSON = @($Input) | ConvertTo-Json -Depth $Depth | Out-String

  $WebPage = @"
<!DOCTYPE HTML>
<html>
<head>
  <title>JSONEditor | Switch mode</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <!-- when using the mode "code", it's important to specify charset utf-8 -->
  <meta http-equiv="Content-Type" content="text/html;charset=utf-8">
  <link href="https://cdnjs.cloudflare.com/ajax/libs/jsoneditor/5.5.6/jsoneditor.min.css" rel="stylesheet" type="text/css">
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jsoneditor/5.5.6/jsoneditor.min.js"></script>
  <style type="text/css">
    body {
      font: 10.5pt arial;
      color: #4d4d4d;
      line-height: 150%;
      width: 100%;
    }
    code {
      background-color: #f5f5f5;
    }
    #jsoneditor {
      width: 90%;
    }
  </style>
</head>
<body>
<p>
  Switch editor mode using the mode box.
  Note that the mode can be changed programmatically as well using the method
  <code>editor.setMode(mode)</code>, try it in the console of your browser.
</p>
<div id="jsoneditor"></div>
<script>
  var container = document.getElementById('jsoneditor');
  var options = {
    mode: 'tree',
    modes: ['code', 'form', 'text', 'tree', 'view'], // allowed modes
    onError: function (err) {
      alert(err.toString());
    },
    onModeChange: function (newMode, oldMode) {
      console.log('Mode switched from', oldMode, 'to', newMode);
    }
  };
  var json = $JSON
  var editor = new JSONEditor(container, options, json);
</script>
</body>
</html>
"@

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