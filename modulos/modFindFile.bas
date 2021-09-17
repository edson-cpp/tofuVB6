Attribute VB_Name = "modFindFile"
Option Explicit

Public Function FindFile(ByVal Path As String, ByVal File As String) As String
    Dim DirName As String, LastDir As String
    If File = "" Then Exit Function
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    DirName = Dir(Path & "*.*", vbDirectory)
    Do While Not FileExist(Path & File)
        If DirName = "" Then Exit Do
        DoEvents
        If DirName <> "." And DirName <> ".." Then
            If (GetAttr(Path & DirName) And vbDirectory) = vbDirectory Then
                LastDir = DirName
                DirName = FindFile(Path & DirName & "\", File)
                If DirName <> "" Then
                    Path = DirName
                    Exit Do
                End If
                DirName = Dir(Path, vbDirectory)
                Do Until DirName = LastDir Or DirName = ""
                    DirName = Dir
                Loop
                If DirName = "" Then Exit Do
            End If
        End If
        DirName = Dir
    Loop
    If FileExist(Path & File) Then FindFile = Path
End Function

Public Function FileExist(Path As String) As Integer
    Dim Canal As Integer
    Canal = FreeFile
    On Error Resume Next
    Open Path For Input As Canal
    If Err = 0 Then
        FileExist = True
    Else
        FileExist = False
    End If
    Close Canal
End Function

