' Launch Court Visitor App without showing console window
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the directory where this script is located
strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Change to the app directory and run the Python app
objShell.CurrentDirectory = strScriptPath
objShell.Run "pythonw.exe guardianship_app.py", 0, False

' Clean exit
Set objShell = Nothing
Set objFSO = Nothing
