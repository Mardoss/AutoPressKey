' This VBS script is used to start the ScrollLockTray PowerShell script.
' It checks if the ScrollLockTray.ps1 file exists in the same directory as this script.
' ScrollLockTray.ps1 enables and disables Scroll Lock key functionality in defined intervals.

Option Explicit

Dim objFSO, objShell, objWshShell, strScriptPath, strPSScriptPath
Dim strFileName

strFileName = "AutoPressKey.ps1"

' Create file system and shell objects
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")
Set objWshShell = WScript.CreateObject("WScript.Shell")

' Get the directory where this VBS script is located
strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
strPSScriptPath = objFSO.BuildPath(strScriptPath, strFileName )

' Check if the PowerShell script exists
If Not objFSO.FileExists(strPSScriptPath) Then
    WScript.Echo "Error: " & strFileName & " not found at: " & strPSScriptPath
    WScript.Quit(1)
End If

' Run the PowerShell script with hidden window
Dim strCommand, intWindowStyle, bWaitOnReturn, intErrorLevel
strCommand = "powershell.exe -NoProfile -WindowStyle Hidden -File """ & strPSScriptPath & """"
intWindowStyle = 0 ' 0 = Hidden window
bWaitOnReturn = False ' Don't wait for PowerShell to exit

On Error Resume Next
intErrorLevel = objWshShell.Run(strCommand, intWindowStyle, bWaitOnReturn)
If Err.Number <> 0 Then
    WScript.Echo "Error: Failed to start " & strFileName & " Error: " & Err.Description
    WScript.Quit(1)
End If

' No need to display success message as script runs silently
' If you want to see a message, uncomment the following line:
' WScript.Echo "ScrollLockTray started successfully."