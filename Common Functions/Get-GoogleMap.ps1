function Show-Map {
    <#
        .Synopsis
        Launches a map in the browser using an address from the command line or the clipboard
    #>
    param(
        [string]$address,
        [Switch]$UseBing
    )

    if(!$address) {
        Add-Type -AssemblyName System.Windows.Forms
        $address=[System.Windows.Forms.Clipboard]::GetText()
    }

    if($UseBing) {
        start "http://www.bing.com/maps?q=$address"
    } else {
        start "http://www.google.com/maps/place/$address"
    }
}