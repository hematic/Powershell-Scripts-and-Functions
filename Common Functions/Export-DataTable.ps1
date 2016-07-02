#From Doug Finke

Add-Type -AssemblyName System.Web

function New-Title {
    param(
        [Parameter(ValueFromPipeline=$true)]
        $title
    )
    Process {
        @{title=$title}
    }
}

function Export-DT{
    param(
        [Parameter(ValueFromPipeline=$true)]
        $TargetData,
        [Switch]$Raw
    )

    Begin {
        $dataSet=@()
        $htmlFileName=[System.IO.Path]::GetTempFileName() -replace "tmp","html"
    }

    Process {
        $columnNames=$TargetData.psobject.properties.name

        $element = foreach ($name in $columnNames) {

            $data=$TargetData.$name
            if(!$data) {
                $data="[not displayable]"
            }

            "`"{0}`"" -f ([System.Web.HttpUtility]::JavaScriptStringEncode($data))
        }

        $dataSet+="`t[{0}]" -f ($element -join ',')
    }

    End {
        $dataset="[`n{0}`n];" -f ($dataSet -join ",`n")
        $columns=$columnNames | New-Title | ConvertTo-Json -Compress

$html=@"
<html>
<head>
  <link href="https://cdn.datatables.net/1.10.11/css/jquery.dataTables.min.css" rel="stylesheet" />
  <script src="https://code.jquery.com/jquery-1.12.0.min.js"></script>
  <script src="https://cdn.datatables.net/1.10.11/js/jquery.dataTables.min.js"></script>

  <script >
  `$(document).ready(function() {

var dataSet= $($dataset)

  var table = `$('#targetTable').DataTable({
    data: dataSet,
     columns: $($columns)
  });

});
  </script>
</head>

<body>
  <table id="targetTable" class="display" width="100%" />
</body>
</html>
"@
        if($Raw) {
            $html
        } else {
            $html | Set-Content -Encoding Ascii $htmlFileName
            ii $htmlFileName
        }
    }
}

Set-Alias odt export-dt
