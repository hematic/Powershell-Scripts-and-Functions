Function Get-YoutubePlaylist 
{

 Param (

        #.PARAMETER Playlisturl = The URL of the youtube playlist.
        
            [Parameter(Mandatory=$True,Position=0)]
            [String]$Playlisturl

          )

        
        $VideoUrls= (invoke-WebRequest -uri $Playlisturl).Links
        $DownloadLinks = @()
        $x=0

        While ($x -lt $VideoUrls.Count)
        {
            if($VideoUrls[$x].innerhtml -notmatch 'SPAN class')
            {
                $DownloadLinks += ('www.youtube.com' + $videourls[$x].href)
                $X++
            }
            Else
            {
                $X++
            }
        }

        Foreach ($link in $DownloadLinks)
        {
			C:\Users\Phillip\Downloads\Tabletop $link -o 'D:\CGP Grey\%(title)s.%(ext)s'
        }
	
}

$PlaylistUrl = "https://www.youtube.com/playlist?list=PL7atuZxmT956cWFGxqSyRdn6GWhBxiAwE"