#region Variable Declarations
#$SecPassword = ConvertTo-SecureString "$ENV:SAPassword" -AsPlainText -Force
#$GLOCred = New-Object System.Management.Automation.PSCredential ($ENV:SAUsername, $SecPassword)
$GLOCred = Get-Credential
$Date = Get-Date -Uformat %Y-%m-%d
#$Reportpath = "D:\Job Output\Automated Security Groups\SGCReport - $Date.xlsx"
$Reportpath = "C:\Temp\SGCReport - $Date.xlsx"
#endregion

#region Import Modules
#AD Module Used for Queries. PSExcel module used for reporting and outputting all the data to excel.

Import-Module ActiveDirectory
Import-Module PSExcel

#endregion

#region Function Declarations

function Get-SiteCode{

    Param(
    
        [String]$Office
    )
        $Office = Fix-Jabroni -Office $Office
        $Description = Get-ADOrganizationalUnit -Filter {Name -eq $Office } -Properties Description  | select -ExpandProperty Description
        $SiteCode = ([regex]::matches($Description, "(?:Site Code:)(\w{3})")).groups[1].value

        If($Sitecode){
            Return $SiteCode
        }

        Else{
            Return "Failed Query or Parse"
        }
}

function Fix-Jabroni{
    
    #Thanks Kervin
    Param(
    
        [String]$Office
    )

    switch ($Office)
    {
        'Düsseldorf'      {$Office = 'Duesseldorf'}
        'Manila (DSM)'    {$Office = 'Manila'}
        'Silicon Valley'  {$office = 'Palo Alto'}
        'Washington D.C.' {$Office = 'Washington DC'}
        Default {}
    }

    Return $Office

}

#endregion

#region Gather SQl Data

$Session = New-PsSession -ComputerName 'am1mfdb001' -Credential $GLOCred

If (!$Session){
    Write-output "Unable to connect to the remote machine."
    exit;
}

$Query = Invoke-Command -ComputerName 'am1mfdb001' -Credential $GLOCred -ScriptBlock {

		$bJobStatus = 0
		
		#SQL Statement
		$cSQLStmt = @"
            SELECT [PersonID]
      ,[LoginID]
      ,[GivenName]
      ,[FamilyName]
      ,[EmailAddress]
      ,[PhysicalOfficeCode]
      ,[OfficeName]
      ,[RegionalSectionID]
      ,[Description]
      ,[PublicTitle]
  FROM [dbo].[vw_SecurityGroup]

"@
		
		$SqlCon = New-Object System.Data.SqlClient.SqlConnection
		$SqlCon.ConnectionString = "Server = am1mfdb001\wcdata; Database = ODS; Integrated Security = True; Trusted_Connection = True"
		$SqlCon.Open()
		
		#-- SQL command to get instance list
		$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
		$SqlCmd.CommandTimeout = 10
		$SqlCmd.CommandText = $cSQLStmt
		$SqlCmd.Connection = $SqlCon
		$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$SqlAdapter.SelectCommand = $SqlCmd
		
		$DataSet = New-Object System.Data.DataSet
		[Void]$SqlAdapter.Fill($DataSet)
		$Result = $DataSet.tables[0]
		$SqlCon.close()
		$Result
	}

Get-PSSession | Remove-PSSession

#endregion

#region Verify SQL Data

If(!$Query[0].OfficeName -or !$Query[0].emailaddress)
{
    Write-output "Returned Data appears to be in the wrong format."
    exit;
}

#endregion

#region Gather Unique offices and Security Groups

$UniqueOffices = $Query.officename | Where-object {$_ -ne 'remote'} | select -Unique | Sort-Object
$SecurityGroups = Get-ADGroup -Filter {Name -like "*-P-AllOffice"}

#endregion

#region Create Missing Security Groups

$newgroups = @()
$LocCodes = @()

Foreach ($Office in $UniqueOffices)
{
    [String]$SiteCode = Get-SiteCode -Office $Office

    If($SiteCode -eq "Failed Query or Parse"){
    
        Write-Output "$Office was not able to be parsed."
    }
    
    Else {

        $Object = New-Object PSObject -Property @{

        Office       = $Office                
        SiteCode     = $SiteCode 
        }

        $LocCodes += $Object

        If((($SecurityGroups | Where-Object {$_.samaccountname -eq $SiteCode + '-p-AllOffice'} | measure-object).count) -gt 0)
        {}

        Else{
        
            $NewGroups += $SiteCode + '-P-AllOffice'
        
            New-ADGroup -Name $($SiteCode + '-P-AllOffice') `
			    -SamAccountName $($SiteCode + '-P-AllOffice') `
			    -GroupCategory Security `
			    -GroupScope Global `
			    -DisplayName $($SiteCode + '-P-AllOffice') `
			    -Path "OU=Automated Groups,OU=Security Groups,OU=FIRMWIDE,DC=WCNET,DC=whitecase,DC=com" `
			    -Description "Automated Group created by leveraging ODS." `
			    -Credential $GLOCred
        }

    }

}


#endregion

#region Verify New Creations

If($NewGroups)
{
    Write-output "Checking new creations."

    Foreach($Group in $NewGroups)
    {
        $Created = Get-ADgroup -Identity $Group

        If($Created)
        {
            Write-output "$($Group) Created Successfully"
        }

        Else
        {
            Write-output "Failed to create $($Group)"
        }
    }
}

#endregion

#region Determine Valid AD Users

$ADUsers = @()
$nonadusers = @()
$FilteredUsers = $Query | Where-Object {$_.officename -ne 'remote'}

Foreach($User in $FilteredUsers)
{
    Try
    {
        $UserCheck = Get-ADUser -Identity $User.LoginID -ErrorAction Stop -ErrorVariable UserError
    }

    Catch
    {
        $UserError = $_.exception.message
        Continue;
    }

    Finally
    {
        If($UserError)
        {
            $nonadusers += $User
        }

        Else
        {
            $SiteCode = $LocCodes | Where-Object {$_.office -eq $User.OfficeName} | select -ExpandProperty SiteCode
            Add-Member -InputObject $UserCheck -MemberType NoteProperty -Name SiteCode -Value $Sitecode -Force
            $adusers += $UserCheck
        }
    }
        
} 

#endregion

#region Add missing members to each group

$ADGroups = Get-ADGroup -Filter {Name -like "*-P-AllOffice*"}

Foreach($Group in $ADGroups)
{
    $LocationUsers = $ADUsers | Where-Object {$Group.name -like "*$($_.SiteCode)*"}

    Try
    {
        Add-ADGroupMember -Identity $Group -Members $LocationUsers -Credential $GLOCred -ErrorAction Stop
    }

    Catch
    {
        Continue;
    }
}

#endregion

#region Reporting Stuff
$ADGroups = $ADGroups | Sort-Object
Foreach($Group in $ADGroups)
{
    $Members = Get-ADGroupMember -Identity $Group | select -Property Name,SamAccountName
    $members | Export-XLSX -Path $Reportpath -worksheetname $Group.Name
}

#endregion

#To Delete the test groups
#get-adgroup -Filter {name -like "*-P-AllOffice*"} | Remove-ADGroup -Confirm:$False -Credential $GLOCred
