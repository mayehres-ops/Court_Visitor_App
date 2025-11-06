' Create Desktop Shortcut for GuardianShip App
Set WshShell = CreateObject("WScript.Shell")
DesktopPath = WshShell.SpecialFolders("Desktop")
Set oShellLink = WshShell.CreateShortcut(DesktopPath & "\GuardianShip App.lnk")

oShellLink.TargetPath = "C:\GoogleSync\GuardianShip_App\launch_app.bat"
oShellLink.WorkingDirectory = "C:\GoogleSync\GuardianShip_App"
oShellLink.Description = "Court Visitor App (NEW VERSION)"
oShellLink.IconLocation = "shell32.dll,165"
oShellLink.Save

WScript.Echo "Desktop shortcut created: GuardianShip App.lnk"
