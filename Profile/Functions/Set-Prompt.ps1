function Set-Prompt
{
    $LastCmd = Get-History -Count 1
    $nextCommand = $lastId + 1
    $currentDirectory = Split-Path (Get-Location) -Leaf
    $host.UI.RawUI.WindowTitle = Get-Location
    "$nextCommand PS:$currentDirectory>"
} 