Function Get-MD5Hash{

Param (
        <#
        .PARAMETER NewName = The path of the file to get a hash for.
        #>
                [Parameter(Mandatory=$True,Position=0)]
                [String]$FilePath
      

$md5 = new-object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
$hash = ([System.BitConverter]::ToString($md5.ComputeHash([System.IO.File]::ReadAllBytes($FilePath)))).replace('-','')
$hash.ToLower()
}
]