Attribute VB_Name = "modGetDir"
Option Explicit
Public Retorno As String
Public DefDir As String

Public Property Get GetDir(DefaultDir As String) As String
    DefDir = DefaultDir
    frmGetDir.Show vbModal
    GetDir = Retorno
End Property
