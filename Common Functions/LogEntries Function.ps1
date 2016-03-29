Function SendTo-LogEntries
{
    Param
    (
		[Parameter(Mandatory = $true,Position = 0)]
		[STRING]$Token,
		[Parameter(Mandatory = $true,Position = 0)]
		[STRING]$Message   
    )
    $tcpConnection = New-Object System.Net.Sockets.TcpClient('data.logentries.com', '80')
    $tcpStream = $tcpConnection.GetStream()
    $reader = New-Object System.IO.StreamReader($tcpStream)
    $writer = New-Object System.IO.StreamWriter($tcpStream)
    $writer.AutoFlush = $true
    $buffer = new-object System.Byte[] 1024
    $encoding = new-object System.Text.AsciiEncoding 
    $writer.WriteLine("$Token $Message")
    $reader.Close()
    $writer.Close()
    $tcpConnection.Close()
}

