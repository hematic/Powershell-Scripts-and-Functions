function Say-Text 
{

    Param (
    <#
    .PARAMETER Message = The phrase to say.
    #>
            [Parameter(Mandatory=$True,Position=0)]
            [String]$Message

            )
    
        [Reflection.Assembly]::LoadWithPartialName('System.Speech') | Out-Null   
        $object = New-Object System.Speech.Synthesis.SpeechSynthesizer 
        $object.Speak($Message)
        $object.Speak($Message) 
        $object.Speak($Message)  
}