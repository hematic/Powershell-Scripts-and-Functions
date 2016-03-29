<#
.Synopsis
   Logon History
.DESCRIPTION
   Show interactive user logon history of the users on a target computer from the Security log.

.EXAMPLE
   Get-LogonHistory "Desktop" "3" "3"
   Returns the User Name, Firstname, Surename, Logon time, logoff time
.NOTES
   LOGON event Log Name: Security, Source: Microsoft-Windows-Security-Auditing, ID: 4624
   LOGOFF event Log Name: Security, Source: Microsoft-Windows-Security-Auditing, ID: 4634
   WORKSTATION_LOCKED event Log Name: Security, Source: Microsoft-Windows-Security-Auditing, ID: 4800
   WORKSTATION_UNLOCKED event Log Name: Security, Source: Microsoft-Windows-Security-Auditing, ID: 4801
   Logon Types: [ref]http://www.windowsecurity.com/articles-tutorials/misc_network_security/Logon-Types.html
   #>

    Param (

        <#
        .PARAMETER PastDays = The number of days back to search for logons.
        #>
                [Parameter(Mandatory=$True,Position=0)]
                [String]$PastDays,

        <# 
        .PARAMETER Days = How many days logons would you like to See
        #>
                [Parameter(Mandatory=$True, position=1)]
                [int]$Days

      )

        $LogonType = @{
            Interactive = 2
            Network = 3
            Batch = 4
            Service = 5
            Unlock = 7
            NetworkCleartext = 8
            NewCredentials = 9
            RemoteInteractive = 10
            CachedInteractive = 11
                     }



        # Use the .Date property to reset the time to 00:00:00
        [datetime]$StartDay = (Get-Date).AddDays( - $PastDays).Date
        [datetime]$StopDay = $StartDay.AddDays($Days).Date

        
        $EventLog = Get-WinEvent -ComputerName localhost -FilterHashtable @{
                            Logname='Security';
                            Id=4624;
                            StartTime=$StartDay;
                            EndTime=$StopDay
                        }



        ForEach ($Event in $EventLog) 
        {

            $xml = [xml]$Event.ToXml()
            $ShortName = $xml.Event.EventData

            ForEach ($data in $ShortName.Data)
            {

                $Event | Add-Member -Force -NotePropertyName $data.name -NotePropertyValue $data.'#text'
            }
        }

        # Select only "real user at the keyboard" logon types.
        $EventLog |
            Where logonType -In $LogonType.Interactive, $LogonType.CachedInteractive |
            Select-Object @{name='Username'; expression={ $_.TargetUserName }},
                          @{name='Date/Time'; expression={ $_.TimeCreated }} |
            Export-csv C:\Users\pmarshall\Desktop\test.csv -NoTypeInformation

                          
