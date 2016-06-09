#Requires -Version 2.0
Set-StrictMode -Version Latest

Push-Location $(Split-Path $Script:MyInvocation.MyCommand.Path)
. .\include\Export-HtmlReport.ps1

# A simple example for the usage of Export-HtmlReport:
#
# A report is generated from a single PowerShell object

@{Object = Get-ChildItem "Env:"} | Export-HtmlReport -OutputFile "Example1.html" | Invoke-Item

Pop-Location
