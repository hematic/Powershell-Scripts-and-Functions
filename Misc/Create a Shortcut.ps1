#Proper Form

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$Home\Desktop\Calc.lnk")
$shortcut.iconlocation = "C:\Users\phillip\Downloads\Mario-icon.ico"
$Shortcut.TargetPath = "Calc"
$Shortcut.Save()

#One liner - $S = (New-Object -ComObject WScript.Shell).CreateShortcut("c:\users\phillip\Desktop\Test Link.lnk");$S.iconlocation = "C:\Users\phillip\Downloads\Mario-icon.ico";$S.TargetPath = "http://www.google.com";$S.Save()