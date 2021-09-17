VERSION 5.00
Begin VB.Form frmGetDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selecione o Diretório"
   ClientHeight    =   4680
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4845
   Icon            =   "frmGetDir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnRede 
      Caption         =   "&Rede"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.DriveListBox drvGetDir 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   3495
   End
   Begin VB.DirListBox dirGetDir 
      Height          =   3240
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton btnCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label labDrive 
      AutoSize        =   -1  'True
      Caption         =   "Dri&ve:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   420
   End
   Begin VB.Label labDir 
      AutoSize        =   -1  'True
      Caption         =   "Diretório:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "frmGetDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function WNetConnectionDialog Lib "mpr.dll" _
(ByVal hWnd As Long, ByVal dwType As Long) As Long
Private Drive As String

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub btnCancelar_Click()
    modGetDir.Retorno = Empty
    Unload Me
End Sub

Private Sub btnOk_Click()
    modGetDir.Retorno = dirGetDir.Path
    Unload Me
End Sub

Private Sub btnRede_Click()
    Call WNetConnectionDialog(Me.hWnd, 1)
    drvGetDir.Refresh
End Sub

Private Sub dirGetDir_Change()
    labDir.Caption = dirGetDir.Path
End Sub

Private Sub drvGetDir_Change()
    On Error GoTo Erro
    dirGetDir.Path = drvGetDir.Drive
    On Error GoTo 0
    Exit Sub
Erro:
    If Err.Number = 68 Then
        MsgBox "Dispositivo indisponível", vbExclamation, "Falha no Dispositivo"
        drvGetDir.Drive = "C:"
    Else
        MsgBox "Não foi Possível Selecionar o Dispositivo." _
            & Chr(13) & "Erro # " & Str(Err.Number) & " foi gerado por " _
            & Err.Source & Chr(13) & Err.Description, vbCritical, _
            "Falha no Dispositivo"
        drvGetDir.Drive = "C:"
    End If
    On Error GoTo 0
End Sub

Private Sub Form_Load()
    drvGetDir.Drive = Mid(modGetDir.DefDir, 1, 2)
    Drive = drvGetDir.Drive
    If UCase(modGetDir.DefDir) = "C:" Then
        dirGetDir.Path = modGetDir.DefDir & "\"
    Else
        dirGetDir.Path = UCase(modGetDir.DefDir)
    End If
    labDir.Caption = dirGetDir.Path
End Sub

