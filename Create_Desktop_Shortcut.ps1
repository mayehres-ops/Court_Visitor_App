# PowerShell script to create Desktop shortcut for Court Visitor App

$WshShell = New-Object -comObject WScript.Shell
$Desktop = [System.Environment]::GetFolderPath('Desktop')
$ShortcutPath = Join-Path $Desktop "Court Visitor App.lnk"

$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = "C:\GoogleSync\GuardianShip_App\Launch Court Visitor App.vbs"
$Shortcut.WorkingDirectory = "C:\GoogleSync\GuardianShip_App"
$Shortcut.Description = "Launch the Court Visitor App"
$Shortcut.IconLocation = "C:\Windows\System32\shell32.dll,21"  # Folder icon
$Shortcut.Save()

Write-Host "[OK] Desktop shortcut created: $ShortcutPath" -ForegroundColor Green
Write-Host "You can now double-click 'Court Visitor App' on your Desktop to launch the app."
Write-Host ""
Write-Host "Press any key to close..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
