VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAutorizacao 
   Caption         =   "Autorização"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2985
   Icon            =   "frmAutorizacao.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   102
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   199
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar stbSenha 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   1230
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4736
            Text            =   "Enter = Validar -|- Esc = Sair"
            TextSave        =   "Enter = Validar -|- Esc = Sair"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtSenha 
      Appearance      =   0  'Flat
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Image imgSenha 
      Height          =   960
      Left            =   120
      Picture         =   "frmAutorizacao.frx":0CCA
      Top             =   120
      Width           =   960
   End
   Begin VB.Label labSenha 
      AutoSize        =   -1  'True
      Caption         =   "Senha:"
      Height          =   195
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   510
   End
End
Attribute VB_Name = "frmAutorizacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub

Private Sub txtSenha_DblClick()
    If txtSenha.Text = Empty Then
        MsgBox "Por Favor Informe a Senha."
        Exit Sub
    End If
    Dim Cript As clsCrypt
    Set Cript = New clsCrypt
    If Cript.CriptSenha(txtSenha.Text) = _
        frmVale.Conn.Rs.Fields.Item("senha").Value Then
        Unload Me
        frmVale.Auth = True
    Else
        Beep
        MsgBox "Senha Inválida ", vbInformation, Me.Caption
    End If
End Sub

Private Sub txtSenha_GotFocus()
    txtSenha.SelStart = 0
    txtSenha.SelLength = 12
    txtSenha.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13 ' Enter
            txtSenha_DblClick
        Case 27 ' Esc
            Unload Me
    End Select
    KeyAscii = modGetKeyAscii.LetrasENumeros(KeyAscii)
End Sub

Private Sub txtSenha_LostFocus()
    txtSenha.BackColor = &H80000005 'Branco
End Sub
