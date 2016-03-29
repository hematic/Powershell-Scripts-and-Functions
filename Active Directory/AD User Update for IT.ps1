Import-Module ActiveDirectory
$Users = import-csv -path "C:\Users\PMarshall\Documents\My Received Files\TitleImport.csv"

Foreach($User in $Users)
{
    $Identity = $User.'Legal First Name'.Substring(0,1) + $User.'Legal Last Name'
    Write-Output "Beginning User : $Identity"
    Write-Output "Manager = $($User.manager)"
    $UserCheck = Get-ADUser $Identity -Properties *

    If($UserCheck -eq $Null)
    {
        Write-Output "User $Identity not found in Active Directory."
        Write-Output "---------------------------------"
        Continue;
    }
    #Get the manager DN
    $RegexResult = ([regex]::matches($($User.manager), "(.*)(?:,\s)(.)"))
    $ManagerSearchName = $RegexResult.groups[2].value + $RegexResult.groups[1].value
    $ManagerResult = (Get-ADUser $ManagerSearchName -Properties *).distinguishedname
    Write-Output "Manager DN = $ManagerResult"

    If($ManagerResult -notmatch $($RegexResult.groups[1].value))
    {
        Write-Output "Seems like the manager result might not have pulled correctly."
        Write-Output "---------------------------------"
        Continue;
    }

    Set-ADUser  -Identity "$Identity" `
                -Company $($User.'Company Working For') `
                -Department $($User.Dept) `
                -Title $($User.Title) `
                -Manager $ManagerResult

    $UpdateResults = Get-ADUser -Identity $Identity -Properties *

    If($($UpdateResults.company) -ne $($UserCheck.company))
    {
        Write-Output "Company Updated! | Old Value : $($UserCheck.company) | New Value $($Updateresults.company)"
    }

    If($($UpdateResults.department) -ne $($UserCheck.department))
    {
        Write-Output "Department Updated! | Old Value : $($UserCheck.department) | New Value $($Updateresults.department)"
    }

    If($($UpdateResults.title) -ne $($UserCheck.title))
    {
        Write-Output "Title Updated! | Old Value : $($UserCheck.title) | New Value $($Updateresults.title)"
    }

    If($($UpdateResults.manager) -ne $($UserCheck.manager))
    {
        Write-Output "Manager Updated! | Old Value : $($UserCheck.manager) | New Value $($Updateresults.manager)"
    }

    Write-Output "SUCCESS"
    Write-Output "---------------------------------"
}