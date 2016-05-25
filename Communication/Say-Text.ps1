function Say-Text {

Param (
        <#
        .PARAMETER NewName = The phrase to say.
        #>
                [Parameter(Mandatory=$True,Position=0)]
                [String]$text

              )
    
    [Reflection.Assembly]::LoadWithPartialName('System.Speech') | Out-Null   
    $object = New-Object System.Speech.Synthesis.SpeechSynthesizer 
    $object.Speak($Text) 
}

Say-text