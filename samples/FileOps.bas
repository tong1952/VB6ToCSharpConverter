Attribute VB_Name = "FileOps"
Option Explicit

Public Sub ReadFile(ByVal path As String)
    Dim fileNum As Integer
    Dim line As String
    fileNum = FreeFile()
    On Error GoTo ErrH
    Open path For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        Debug.Print line
    Loop
    Close #fileNum
    Exit Sub
ErrH:
    MsgBox "Error reading file: " & Err.Description
    Close #fileNum
End Sub

Public Sub WriteFile(ByVal path As String, ByVal content As String)
    Dim fileNum As Integer
    fileNum = FreeFile()
    Open path For Output As #1
    Print #1, content
    Close #1
End Sub

Public Sub AppendFile(ByVal path As String, ByVal line As String)
    Open path For Append As #2
    Print #2, line
    Close #2
End Sub

Public Sub DeleteFile(ByVal path As String)
    On Error Resume Next
    Kill path
End Sub

Public Sub RenameFile(ByVal oldPath As String, ByVal newPath As String)
    Name oldPath As newPath
End Sub

Public Sub CreateFolder(ByVal path As String)
    MkDir path
End Sub

Public Sub DeleteFolder(ByVal path As String)
    RmDir path
End Sub

Public Sub CopyFile(ByVal src As String, ByVal dst As String)
    FileCopy src, dst
End Sub
