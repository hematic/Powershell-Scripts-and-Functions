Function Export-HtmlReport
{

<#
.SYNOPSIS
	Creates a HTML report
.DESCRIPTION
	Creates an eye-friendly HTML report with an inline cascading style sheet from given PowerShell objects and saves it to a file.
.PARAMETER InputObject
	Hashtable containing data to be converted into a HTML report.
	
	HashTable Proberties:
	
	[array]  .Object       Any PowerShell object to be converted into a HTML table.

	[string] .Property     Select a set of properties from the object. Default is "*".
	
	[sting]  .As           Use "List" to create a vertical table instead of horizontal alignment.
			
	[string] .Title        Add a table headline.
	
	[string] .Description  Add a table description.

.PARAMETER Title
	Title of the HTML report.
.PARAMETER OutputFileName
	Full path of the output file to create, e.g. "C:\temp\output.html".
.EXAMPLE
	@{Object = Get-ChildItem "Env:"} | Export-HtmlReport -OutputFile "HtmlReport.html" | Invoke-Item

	Creates a directory item HTML report and opens it with the default browser.
.EXAMPLE

	$ReportTitle = "HTML-Report"
	$OutputFileName = "HtmlReport.html"
	
	$InputObject =  @{ 
	                   Title  = "Directory listing for C:\";
	                   Object = Get-Childitem "C:\" | Select -Property FullName,CreationTime,LastWriteTime,Attributes
					},
					@{
	                   Title  = "PowerShell host details";
					   Object = Get-Host;
					   As     = 'List'
					},
					@{
	                   Title       = "Running processes";
					   Description = 'Information about the first 2 running processes.'
					   Object      = Get-Process | Select -First 2
					},
					@{
					   Object = Get-ChildItem "Env:"
					}
					
	Export-HtmlReport -InputObject $InputObject -ReportTitle $ReportTitle -OutputFile $OutputFileName

	Creates a HTML report with separated tables for each given object.
.INPUTS
	Data object, title and alignment parameter
.OUTPUTS
	File object for the created HTML report file.
#>

	[CmdletBinding()]
	Param(
		[Parameter(ValueFromPipeline=$True, Mandatory=$True)]
		[ValidateNotNullOrEmpty()]
		[Array]$InputObject,

		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[String]$ReportTitle = 'Generic HTML-Report',

		[Parameter(Mandatory=$True)]
		[ValidateNotNullOrEmpty()]
		[String]$OutputFileName
	)

	BEGIN
	{
		$HtmlTable		= ''
	}

	PROCESS
	{
		ForEach ($InputElement in $InputObject)
		{
			If ($InputElement.ContainsKey('Title') -eq $False)
			{
				$InputElement.Title = ''
			}

			If ($InputElement.ContainsKey('As') -eq $False)
			{
				$InputElement.As = 'Table'
			}

			If ($InputElement.ContainsKey('Property') -eq $False)
			{
				$InputElement.Property = '*'
			}

			If ($InputElement.ContainsKey('Description') -eq $False)
			{
				$InputElement.Description = ''
			}

			$HtmlTable += $InputElement.Object | ConvertTo-HtmlTable -Title $InputElement.Title -Description $InputElement.Description -Property $InputElement.Property -As $InputElement.As
			$HtmlTable += '<br>'
		}
	}

	END
	{
		$HtmlTable | New-HtmlReport -Title $ReportTitle | Set-Content $OutputFileName
		Get-Childitem $OutputFileName | Write-Output
	}
}

Function ConvertTo-HtmlTable
{

<#
.SYNOPSIS
	Converts a PowerShell object into a HTML table
.DESCRIPTION
	Converts a PowerShell object into a HTML table.	Then use "New-HtmlReport" to create an eye-friendly HTML report with an inline cascading style sheet.
.PARAMETER InputObject
	Any PowerShell object to be converted into a HTML table.
.PARAMETER Property
	Select object properties to be used for table creation. Default is "*"
.PARAMETER As
	Use "List" to create a vertical table. All other values will create a horizontal table.
.PARAMETER Title
	Adds an additional table with a title. Very useful for multi-table-reports!
.PARAMETER Description
	Adds an additional table with a description. Very useful for multi-table-reports!
EXAMPLE
	Get-Process | ConvertTo-HtmlTable
	
	Returns a HTML table as a string.
.EXAMPLE
	Get-Process | ConvertTo-HtmlTable | New-HtmlReport | Set-Content "HtmlReport.html"

	Returns a HTML report and saves it as a file.
.EXAMPLE
	$body =	ConvertTo-HtmlTable -InputObject $(Get-Process) -Property Name,ID,Path -As "List" -Title "Process list" -Description "Shows running processes as a list"
	New-HtmlReport -Body $body | Set-Content "HtmlReport.html"

	Returns a HTML report and saves it as a file.
.INPUTS
	Any PowerShell object
.OUTPUTS
	HTML table as String
#>

	[CmdletBinding()]
	Param(
		[Parameter(ValueFromPipeline=$True, Mandatory=$True)]
		[ValidateNotNullOrEmpty()]
		[PSObject]$InputObject,

		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[Object[]]$Property = '*',

		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[String]$As = 'TABLE',
		
		[Parameter()]
		[String]$Title,

		[Parameter()]
		[String]$Description
	)
	
	BEGIN
	{
		$InputObjectList = @()
		
		If ($As -ne 'LIST')
		{
			$As = 'TABLE'
		}
	}

	PROCESS
	{
		$InputObjectList += $InputObject
	}

	END
	{
		$ofs = "`r`n"	# Set separator for string-convertion to carrige return
		[String]$HtmlTable = $InputObjectList | ConvertTo-HTML -Property $Property -As $As -Fragment
		Remove-Variable ofs -force

		# Add description table
		If ($Description)
		{
			$Html		= '<table id="TableDescription"><tr><td>' + "$Description</td></tr></table>`n"
			$Html 		+= '<table id="TableSpacer"></table>' + "`n"
			$HtmlTable	= $Html + $HtmlTable
		}
		Else
		{
			$Html 		= '<table id="TableSpacer"></table>' + "`n"
			$HtmlTable	= $Html + $HtmlTable
		}
		
		# Add title table
		If ($Title)
		{
			$Html		= '<table id="TableHeader"><tr><td>' + "$Title</td></tr></table>`n"
			$HtmlTable	= $Html + $HtmlTable
		}

		# Add missing data separator tag <hr> to second column (on list-tables only)
		$HtmlTable = $HtmlTable -Replace '<hr>', '<hr><td><hr></td>'
				
		Write-Output $HtmlTable	
	}
}

Function New-HtmlReport
{

<#
.SYNOPSIS
	Creates a HTML report
.DESCRIPTION
	Creates an eye-friendly HTML report with an inline cascading style sheet for a given HTML body.
	Usage of "ConvertTo-HtmlTable" is recommended to create the HTML body.
.PARAMETER Body
	Any HTML body, e.g. a table. Usage of "ConvertTo-HtmlTable" is recommended
	to create an according table from any PowerShell object.
.PARAMETER Title
	Title of the HTML report.
.PARAMETER Head
	Any HTML code to be inserted into the head-tag, e.g. scripts or meta-information.
.PARAMETER CssUri
	Path to a CSS-File to be included as an inline css.
	If CssUri is invalid or not provided, a default css is used instead.
.EXAMPLE
	Get-Process | ConvertTo-HtmlTable | New-HtmlReport
	
	Returns a HTML report as a string.
.EXAMPLE
	Get-Process | ConvertTo-HtmlTable | New-HtmlReport -Title "HTML Report with CSS" | Set-Content "HtmlReport.html"

	Returns a HTML report and saves it as a file.
.EXAMPLE
	$body =	Get-Process | ConvertTo-HtmlTable
	New-HtmlReport -Body $body -Title "HTML Report with CSS" -Head '<meta name="author" content="Thomas Franke">' -CssUri "stylesheet.css" | Set-Content "HtmlReport.html"

	Returns a HTML report with an alternative CSS and saves it as a file.
.INPUTS
	HTML body as String
.OUTPUTS
	HTML page as String
#>

	[CmdletBinding()]
	Param(
		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[String]$CssUri,

		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[String]$Title = 'HTML Report',

		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[String]$Head = '',

		[Parameter(ValueFromPipeline=$True, Mandatory=$True)]
		[ValidateNotNullOrEmpty()]
		[Array]$Body
	)

	# Add title to head because -Title parameter is ignored if -Head parameter is used
	If ($Title)
	{
		$Head = "<title>$Title</title>`n$Head`n"
	}

	# Add inline stylesheet
	If (($CssUri) -And (Test-Path $CssUri))
	{
		$Head += "<style>`n" + $(Get-Content $CssUri | ForEach {"`t$_`n"}) + "</style>`n"
	}
	Else
	{
		$Head += @'
<style>
	table
		{
			Margin: 0px 0px 0px 4px;
			Border: 1px solid rgb(190, 190, 190);
			Font-Family: Tahoma;
			Font-Size: 8pt;
			Background-Color: rgb(252, 252, 252);
		}

	tr:hover td
		{
			Background-Color: rgb(150, 150, 220);
			Color: rgb(255, 255, 255);
		}

	tr:nth-child(even)
		{
			Background-Color: rgb(242, 242, 242);
		}
	  
	th
		{
			Text-Align: Left;
			Color: rgb(150, 150, 220);
			Padding: 1px 4px 1px 4px;
		}


	td
		{
			Vertical-Align: Top;
			Padding: 1px 4px 1px 4px;
		}

	#TableHeader
		{
			Margin-Bottom: -1px;
			Background-Color: rgb(255, 255, 225);
			Width: 30%;
		}
		
	#TableDescription
		{
			Background-Color: rgb(252, 252, 252);
			Width: 30%;
		}

	#TableSpacer
		{
			Height: 6px;
			Border: 0px;
		}
</style>
'@
	}

	$HtmlReport = ConvertTo-Html -Head $Head
	
	# Delete empty table that is created because no input object was given
	$HtmlReport = $($HtmlReport -Replace '<table>', '') -Replace '</table>', $Body

	Write-Output $HtmlReport
}

$ListofComputers = Get-ADComputer -Filter {OperatingSystem -like "*Server*"} -Property *
$ListofComputers | Start-RSjob -Name {"$($_.name)"} -ScriptBlock {
        Param($ListofComputers)

        $TempData = New-Object PSObject -Property @{
            Name        = $ListofComputers.name
            OS          = $ListofComputers.item('OperatingSystem')[0]
            ServicePack = $ListofComputers.item('OperatingSystemServicePack')[0]
            Reachable   = $(Test-Connection -ComputerName $ListofComputers.name -Count 1 -quiet)
        }

        $TempData
}

While ((get-rsjob | Where-object {$_.State -ne 'completed'}).count -gt 0 )   
{
    start-sleep -Seconds 5
}
       
$Data = (Get-RSjob | Receive-RSJob) | Where-object {$_.reachable -eq $True}| Group-Object -Property OS | select count,name | sort-object -Property count -Descending

$Month = $(Get-Date -Format MMMM)

$OutputFileName = "$Env:windir\temp\OS Report $Month" + '.html'
$ReportTitle = "Windows Server OS Report for $Month"
$InputObject =  @{
                    Title  = "Windows Server OS Report for $(Get-Date -Format MMMM)";
                    Object = $Data
                }

Export-HtmlReport -InputObject $InputObject -ReportTitle $ReportTitle -OutputFile $OutputFileName | out-null

$Splat = @{

To = "$ENV:MailTo"
From = "$ENV:MailFrom"
Body = $(get-content -Path $OutputFileName -raw)
Subject =  "Windows Server OS Report for $Month" 
SmtpServer = 'AM1SMTP' 
BodyAsHtml = $True
}

Send-MailMessage @Splat

$Data | FT