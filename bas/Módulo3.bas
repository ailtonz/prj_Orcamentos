Attribute VB_Name = "Módulo3"
Option Explicit

Sub Compact(NameFile As String, NameZipFile As String)
    Dim PathZipProgram As String
    Dim ShellStr As String

    'Path of the Zip program
    PathZipProgram = "C:\program files\7-Zip\"
    If Right(PathZipProgram, 1) <> "\" Then
        PathZipProgram = PathZipProgram & "\"
    End If

    'Check if this is the path where 7z is installed.
    If Dir(PathZipProgram & "7z.exe") = "" Then
        MsgBox "Please find your copy of 7z.exe and try again"
        Exit Sub
    End If

    ShellStr = PathZipProgram & "7z.exe a" _
             & " " & Chr(34) & NameZipFile & Chr(34) _
             & " " & NameFile

    ShellAndWait ShellStr, vbHide



End Sub

Sub DesCompact(FileNameZip As Variant, NameUnZipFolder As String)
    Dim PathZipProgram As String
    Dim ShellStr As String

    'Path of the Zip program
    PathZipProgram = "C:\program files\7-Zip\"
    If Right(PathZipProgram, 1) <> "\" Then
        PathZipProgram = PathZipProgram & "\"
    End If

    'Check if this is the path where 7z is installed.
    If Dir(PathZipProgram & "7z.exe") = "" Then
        MsgBox "Please find your copy of 7z.exe and try again"
        Exit Sub
    End If

    'There are a few commands/Switches that you can change in the ShellStr
    'We use x command now to keep the folder stucture, replace it with e if you want only the files
    '-aoa Overwrite All existing files without prompt.
    '-aos Skip extracting of existing files.
    '-aou aUto rename extracting file (for example, name.txt will be renamed to name_1.txt).
    '-aot auto rename existing file (for example, name.txt will be renamed to name_1.txt).
    'Use -r if you also want to unzip the subfolders from the zip file
    'You can add -ppassword if you want to unzip a zip file with password (only 7zip files)
    'Change "*.*" to for example "*.txt" if you only want to unzip the txt files
    'Use "*.xl*" for all Excel files: xls, xlsx, xlsm, xlsb
    ShellStr = PathZipProgram & "7z.exe x -aoa -r" _
             & " " & Chr(34) & FileNameZip & Chr(34) _
             & " -o" & Chr(34) & NameUnZipFolder & Chr(34) & " " & "*.*"

    ShellAndWait ShellStr, vbHide

End Sub

