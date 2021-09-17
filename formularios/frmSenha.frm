VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSenha 
   Caption         =   "Senha de Acesso"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2985
   Icon            =   "frmSenha.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   118
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   199
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtLogin 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar stbSenha 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   1470
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4763
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
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label labLogin 
      AutoSize        =   -1  'True
      Caption         =   "Login:"
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   435
   End
   Begin VB.Image imgSenha 
      Height          =   960
      Left            =   120
      Picture         =   "frmSenha.frx":0CCA
      Top             =   240
      Width           =   960
   End
   Begin VB.Label labSenha 
      AutoSize        =   -1  'True
      Caption         =   "Senha:"
      Height          =   195
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   510
   End
End
Attribute VB_Name = "frmSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LoginFail As Integer

Private Sub Form_Load()
    LoginFail = Empty
    Me.Caption = frmPrin.TituloSenha
    If frmPrin.CodigoUsuarioLogado = Empty Then
        txtLogin.Text = "admin"
    Else
        txtLogin.Text = frmPrin.LoginUsuarioLogado
    End If
End Sub

Private Sub txtLogin_GotFocus()
    txtLogin.SelStart = 0
    txtLogin.SelLength = 32
    txtLogin.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtLogin_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13 ' Enter
            txtSenha_DblClick
        Case 27 ' Esc
            Unload Me
    End Select
    KeyAscii = modGetKeyAscii.LetrasENumeros(KeyAscii)
End Sub

Private Sub txtLogin_LostFocus()
    txtLogin.BackColor = &H80000005
End Sub

Private Sub txtSenha_DblClick()
    If txtLogin.Text = Empty Then
        MsgBox "Por Favor Informe o Login."
        txtLogin.SetFocus
        Exit Sub
    End If
    If txtSenha.Text = Empty Then
        MsgBox "Por Favor Informe a Senha."
        txtSenha.SetFocus
        Exit Sub
    End If
    Dim Cript As clsCrypt
    Dim Conn As clsMyConnect
    Set Cript = New clsCrypt
    Set Conn = New clsMyConnect
    Call Conn.Connect
    Call Conn.Query("SELECT id, nome, login, senha, nivel " _
        & "FROM usuario WHERE login = '" & txtLogin.Text & "'")
    If Conn.Rs.RecordCount = Empty Then
        Beep
        MsgBox "Usuário Inválido", vbInformation, "Usuário Inválido"
        txtLogin.SetFocus
        Exit Sub
    End If
    If Cript.CriptSenha(txtSenha.Text) = _
        Conn.Rs.Fields.Item("senha").Value Then
        Unload Me
        frmPrin.Auth = True
        frmPrin.CodigoUsuarioLogado = Conn.Rs.Fields.Item("id").Value
        frmPrin.NomeUsuarioLogado = Conn.Rs.Fields.Item("nome").Value
        frmPrin.LoginUsuarioLogado = Conn.Rs.Fields.Item("login").Value
        frmPrin.NivelUsuarioLogado = Conn.Rs.Fields.Item("nivel").Value
    Else
        Beep
        LoginFail = LoginFail + 1
        MsgBox "Senha Inválida " & LoginFail, vbInformation, "Senha Inválida"
        If LoginFail = 3 Then
            Unload Me
        Else
            txtSenha.SetFocus
        End If
    End If
    Conn.Disconnect
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
