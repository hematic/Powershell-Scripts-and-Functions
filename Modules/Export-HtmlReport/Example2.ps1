#Requires -Version 2.0
Set-StrictMode -Version Latest

Push-Location $(Split-Path $Script:MyInvocation.MyCommand.Path)
. .\include\Export-HtmlReport.ps1

# A more complex example for the usage of Export-HtmlReport:
#
# A report is generated from several PowerShell objects with 
# additional parameters provided as an array of hashtables

$OutputFileName	= "Example2.html"
$ReportTitle	= "Example2"

$InputObject =  @{ 
				   Object = Get-ChildItem "Env:" | Select -First 5
				},
				@{
				   Title  = "Directory listing for C:\";
				   Object = Get-Childitem "C:\" | Select -Property FullName,CreationTime,LastWriteTime,Attributes
				},
				@{
				   Title  = "Directory listing for C:\ as a list table (first 3 items only)";
				   Object = Get-Childitem "C:\" | Select -Property FullName,CreationTime,LastWriteTime,Attributes | Select -First 3
				   As     = 'List'
				},
				@{
				   Title       = "Running processes";
				   Description = 'The following table contains information about the first 5 running processes.'
				   Object      = Get-Process | Select -First 5
				}
					
Export-HtmlReport -InputObject $InputObject -ReportTitle $ReportTitle -OutputFile $OutputFileName
Invoke-Item $OutputFileName

Pop-Location
