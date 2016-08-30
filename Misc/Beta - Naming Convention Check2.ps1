#Variable declaration

$path     = "C:\Users\marshph\Desktop\Device Naming policy.doc"
$application = New-Object -comobject word.application
$application.visible = $False
$LocationCodes = @()
$RoleCodes = @()
$FunctionCodes = @()


$Document = $application.documents.open($path,$false,$true)
$Doctext = $document.content.FormattedText.Text
$DocLength = $Doctext.Length


################################################
#Pull The Location Codes From the Word Document#
################################################

#Strips out the first half or so of the document we don't need to parse.
$UsefulText = $Doctext.Substring(6000,$($DocLength - 6000))

#This delimits the beginning and end  of the Location Codes section.
$LocStart = $UsefulText.IndexOf("Appendix A - City/Location Code")
$LocEnd = $UsefulText.IndexOf("Appendix B")

#This defines the exact range of all the location data.
$LocRange = $LocEnd - $LocStart

#This grabs just the location text.
$LocText = $UsefulText.Substring($LocStart, $LocRange)

#The Regex to parse location codes.
$LocRegex = "([A-Z]{2}[A-z0-9])"

#The pulls all the location codes and adds them to an array.
$LocationMatches = ([regex]::matches($LocText, $LocRegex))

Foreach ($LocCode in $LocationMatches)
{
    $LocationCodes += $($LocCode.groups[1].value)
}

############################################
#Pull The Role Codes From the Word Document#
############################################

#This delimits the beginning and end of the Role Codes section.
$RoleStart = $UsefulText.IndexOf("Appendix B")
$RoleEnd = $UsefulText.IndexOf("Appendix C")

#This defines the exact range of all the role data.
$RoleRange = $RoleEnd - $RoleStart

#This grabs just the role text.
$RoleText = $UsefulText.Substring($RoleStart, $RoleRange)

#The Regex to parse location codes.
$RoleRegex = "([A-Z]{2})"

#The pulls all the location codes and adds them to an array.
$RoleMatches = ([regex]::matches($RoleText, $RoleRegex))

Foreach ($RoleCode in $RoleMatches)
{
    $RoleCodes += $($RoleCode.groups[1].value)
}


################################################
#Pull The Function Codes From the Word Document#
################################################

#This delimits the beginning and end of the Function Codes section.
$FunctionStart = $UsefulText.IndexOf("Appendix C")
$FunctionEnd = $UsefulText.IndexOf("Appendix D")

#This defines the exact range of all the function data.
$FunctionRange = $FunctionEnd - $FunctionStart

#This grabs just the role text.
$FunctionText = $UsefulText.Substring($FunctionStart, $FunctionRange)

#This splits each line of the text into its own value in an array
$FunctionText = $Functiontext -split '[\r\n]'

Foreach($Line in $FunctionText)
{
    If ($Line -like '*-*')
    {
        $Code = ([regex]::matches($Line, '([A-Z]{2})(?:[\s])')).groups[1].value
        $Description = (([regex]::matches($Line, '(?:[\s][-]+[\s])(.*)(?:[(])')).groups[1].value).trim()
        $AppliesTo = ([regex]::matches($Line, '(?:[^(]+\()([^)]+)(?:\))')).groups[1].value
        $AppliestoArray = $AppliesTo -split ','
        
        $FunctionObject = New-Object PSObject -Property @{
            Code = $Code
            Description = $Description
            AppliesTo = $AppliestoArray
        }    

        $FunctionCodes += $FunctionObject
    }    
}

