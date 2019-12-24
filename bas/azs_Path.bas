Attribute VB_Name = "azs_Path"
Option Explicit

Public Function pathDesktopAddress() As String
    pathDesktopAddress = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
End Function

Public Function pathWorkSheetAddress() As String
    pathWorkSheetAddress = ActiveWorkbook.Path & "\"
End Function

Public Function pathWorkbookFullName() As String
    pathWorkbookFullName = ActiveWorkbook.FullName
End Function
