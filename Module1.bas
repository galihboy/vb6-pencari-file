Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' GetDirectoryFromPathFilename: Return directory path contain filename
Public Function GetDirectoryFromPathFilename(path As String) As String
    Dim pos As Integer
    pos = InStrRev(path, "\")
    If pos > 0 Then
        GetDirectoryFromPathFilename = Left$(path, pos)
    Else
        GetDirectoryFromPathFilename = ""
    End If
End Function


