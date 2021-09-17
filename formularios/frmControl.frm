VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmControl 
   Caption         =   "Centro de Controle"
   ClientHeight    =   4095
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9090
   Icon            =   "frmControl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   273
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   606
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTipoVenda 
      Appearance      =   0  'Flat
      Caption         =   "Tipo de Venda"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   2760
      Width           =   4815
      Begin VB.OptionButton optMesa 
         Caption         =   "Venda por Mesa"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optCliente 
         Caption         =   "Venda por Cliente"
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox txtNomeFunc 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   3000
      TabIndex        =   5
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox txtCodFunc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1920
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtMesaFechada 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2640
      TabIndex        =   3
      Top             =   1920
      Width           =   450
   End
   Begin VB.TextBox txtTaxa 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1920
      TabIndex        =   2
      Top             =   1560
      Width           =   450
   End
   Begin VB.ComboBox cbbRecibo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Frame fraBackup 
      Appearance      =   0  'Flat
      Caption         =   "Pasta Padrão de Backup"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   4080
      TabIndex        =   14
      Top             =   840
      Width           =   4815
      Begin VB.CommandButton btnAbrir 
         DisabledPicture =   "frmControl.frx":0CCA
         Enabled         =   0   'False
         Height          =   300
         Left            =   4320
         Picture         =   "frmControl.frx":1054
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   840
         Width           =   340
      End
      Begin VB.TextBox txtBackup 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   3855
      End
      Begin VB.OptionButton optUnica 
         Caption         =   "Sempre a Mesma"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optUltimo 
         Caption         =   "Última Pasta Usada"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Timer tmrControl 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7080
      Top             =   3120
   End
   Begin VB.TextBox txtQtdeMesas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   0
      Top             =   840
      Width           =   450
   End
   Begin MSComctlLib.ImageList imlControl 
      Left            =   7560
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControl.frx":13DE
            Key             =   "salvar"
            Object.Tag             =   "salvar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControl.frx":20B8
            Key             =   "sair"
            Object.Tag             =   "sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbControl 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   12
      Top             =   3750
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13626
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1852
            MinWidth        =   1852
            Text            =   "F1 - Ajuda"
            TextSave        =   "F1 - Ajuda"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrControl 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   1111
      BandCount       =   1
      ImageList       =   "imlControl"
      _CBWidth        =   9090
      _CBHeight       =   630
      _Version        =   "6.7.9782"
      Child1          =   "tbrControl"
      MinHeight1      =   38
      Width1          =   200
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrControl 
         Height          =   570
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlControl"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "salvar"
               Object.ToolTipText     =   "Salvar"
               ImageKey        =   "salvar"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sair"
               Object.ToolTipText     =   "Sair"
               ImageKey        =   "sair"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label labGarcomBalcao 
      AutoSize        =   -1  'True
      Caption         =   "Garçom Venda Balcão:"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   2333
      Width           =   1650
   End
   Begin VB.Label labMinutos 
      AutoSize        =   -1  'True
      Caption         =   "minutos."
      Height          =   195
      Left            =   3120
      TabIndex        =   19
      Top             =   1980
      Width           =   585
   End
   Begin VB.Label labMesaFechada 
      AutoSize        =   -1  'True
      Caption         =   "Marcar mesas fechadas a mais de"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   1973
      Width           =   2415
   End
   Begin VB.Label labPercent 
      AutoSize        =   -1  'True
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2400
      TabIndex        =   17
      Top             =   1590
      Width           =   210
   End
   Begin VB.Label labTaxa 
      AutoSize        =   -1  'True
      Caption         =   "Taxa de Serviço:"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label labRecibo 
      AutoSize        =   -1  'True
      Caption         =   "Imprimir recibo para:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   1260
      Width           =   1410
   End
   Begin VB.Label labQtdeMesas 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde de Mesas:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   893
      Width           =   1125
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mniSalvar 
         Caption         =   "&Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mniSair 
         Caption         =   "Sai&r"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "&Ajuda"
      Begin VB.Menu mniConteudo 
         Caption         =   "&Conteúdo"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************
' Quando inserir um novo item só é preciso alte- *
' rar as subs mniSalvar_Click e CarregaValores   *
'*************************************************
Option Explicit
Public Conn As clsMyConnect
Dim LocalizarOQue As String
Public Localizar As Boolean 'Define o retorno da tela localizar

Private Sub btnAbrir_Click()
    Dim ret As String
    Conn.Rs.MoveFirst
    Conn.Rs.Find "campo = 'Backup'"
    ret = modGetDir.GetDir(Mid(Conn.Rs.Fields.Item("valor").Value, 2))
    If Not ret = Empty Then
        txtBackup.Text = ret
    End If
End Sub

Private Sub Form_Load()
    Set Conn = New clsMyConnect
    Call Conn.Connect
    If Conn.NumErro <> 0 Then GoTo SubFail
    Call Conn.Query("SELECT id, campo, valor FROM config")
    If Conn.NumErro <> 0 Then GoTo SubFail
    Call CarregaValores
    Exit Sub
SubFail:
    MsgBox "Não foi Possível Conectar-se com a Base de Dados." _
        & Chr(13) & "Erro # " & Str(Conn.NumErro) & " foi gerado por " _
        & Conn.SrcErro & Chr(13) & Conn.DescErro, vbCritical, "Falha de Conexão"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Conn.Disconnect
End Sub

Private Sub mniSair_Click()
    Unload Me
End Sub

Private Sub mniSalvar_Click()
    Dim Backup As String
    If optUltimo.Value = True Then
        Backup = "0" & Mid(Conn.GetValue("SELECT valor FROM config " _
            & "WHERE campo = 'Backup'"), 2)
    Else
        Backup = "1" & txtBackup.Text
    End If
    frmPrin.TipoVenda = IIf(optMesa.Value, "0", "1")
    frmPrin.mniMesas.Caption = IIf(optMesa.Value, "Controle de &Mesas", "Frente de &Vendas ao Cliente")
    frmPrin.tbrPrin.Buttons.Item("mesa").ToolTipText = frmPrin.mniMesas.Caption

    Call Conn.Query("UPDATE config SET valor = '" & txtQtdeMesas.Text & "' WHERE campo = 'QtdeMesas'")
    Call Conn.Query("UPDATE config SET valor = '" & Backup & "' WHERE campo = 'Backup'")
    Call Conn.Query("UPDATE config SET valor = '" & cbbRecibo.ListIndex & "' WHERE campo = 'Recibo'")
    Call Conn.Query("UPDATE config SET valor = '" & txtTaxa.Text & "' WHERE campo = 'Taxa'")
    Call Conn.Query("UPDATE config SET valor = '" & txtMesaFechada.Text & "' WHERE campo = 'MesaFechada'")
    Call Conn.Query("UPDATE config SET valor = '" & txtCodFunc.Text & "' WHERE campo = 'GarcomBalcao'")
    Call Conn.Query("UPDATE config SET valor = '" & frmPrin.TipoVenda & "' WHERE campo = 'TipoVenda'")
    ' Inserir novo item aqui
    
    If Conn.NumErro <> 0 Then
        MsgBox "Falha na Gravação dos Dados." _
            & Chr(13) & "Erro # " & Str(Conn.NumErro) & " foi gerado por " _
            & Conn.SrcErro & Chr(13) & Conn.DescErro, vbCritical, "Falha de Gravação"
    Else
        tmrControl.Enabled = True
        stbControl.Panels.Item(1).Text = "Registro Salvo com Êxito."
    End If
End Sub

Private Sub optUltimo_Click()
    btnAbrir.Enabled = False
    txtBackup.Enabled = False
End Sub

Private Sub optUnica_Click()
    btnAbrir.Enabled = True
    txtBackup.Enabled = True
    If txtBackup.Text = Empty Then txtBackup.Text = "C:"
End Sub

Private Sub tbrControl_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "salvar"
            mniSalvar_Click
        Case "sair"
            mniSair_Click
    End Select
End Sub

Private Sub CarregaValores()
    If Conn.Rs.BOF Or Conn.Rs.EOF Then GoTo SubFail
    '***********************************
    'Recupera Valor do Campo "QtdeMesas"
    '***********************************
    Conn.Rs.MoveFirst
    Conn.Rs.Find "campo = 'QtdeMesas'"
    If Conn.Rs.EOF Then GoTo SubFail
    txtQtdeMesas.Text = Conn.Rs.Fields.Item("valor").Value
    '***********************************
    'Recupera Valor do Campo "Backup"
    '***********************************
    Conn.Rs.MoveFirst
    Conn.Rs.Find "campo = 'Backup'"
    If Conn.Rs.EOF Then GoTo SubFail
    If Mid(Conn.Rs.Fields.Item("valor").Value, 1, 1) = "1" Then
        optUnica.Value = True
        txtBackup.Text = Mid(Conn.Rs.Fields.Item("valor").Value, 2)
    End If
    '***********************************
    'Recupera Valor do Campo "Recibo"
    '***********************************
    Conn.Rs.MoveFirst
    Conn.Rs.Find "campo = 'Recibo'"
    cbbRecibo.AddItem "Impressora"
    cbbRecibo.AddItem "Tela"
    cbbRecibo.AddItem "Arquivo"
    If Conn.Rs.EOF Then GoTo SubFail
    cbbRecibo.ListIndex = CInt(Conn.Rs.Fields.Item("valor").Value)
    '***********************************
    'Recupera Valor do Campo "Taxa"
    '***********************************
    Conn.Rs.MoveFirst
    Conn.Rs.Find "campo = 'Taxa'"
    If Conn.Rs.EOF Then GoTo SubFail
    txtTaxa.Text = Conn.Rs.Fields.Item("valor").Value
    '***********************************
    'Recupera Valor do Campo "MesaFechada"
    '***********************************
    Conn.Rs.MoveFirst
    Conn.Rs.Find "campo = 'MesaFechada'"
    If Conn.Rs.EOF Then GoTo SubFail
    txtMesaFechada.Text = Conn.Rs.Fields.Item("valor").Value
    '***********************************
    'Recupera Valor do Campo "GarcomBalcao"
    '***********************************
    Conn.Rs.MoveFirst
    Conn.Rs.Find "campo = 'GarcomBalcao'"
    If Conn.Rs.EOF Then GoTo SubFail
    txtCodFunc.Text = Conn.Rs.Fields.Item("valor").Value
    '***********************************
    'Recupera Valor do Campo "TipoVenda"
    '***********************************
    Conn.Rs.MoveFirst
    Conn.Rs.Find "campo = 'TipoVenda'"
    If Conn.Rs.EOF Then GoTo SubFail
    If Conn.Rs.Fields.Item("valor").Value = "1" Then
        optCliente.Value = True
    End If
    '***********************************
    'Recupera Valor do Campo "X"
    '***********************************
    'Inserir novo item aqui
    '***********************************
    
    Exit Sub
SubFail:
    MsgBox "Falha na Leitura dos Dados." _
        & Chr(13) & "Erro # " & Str(Conn.NumErro) & " foi gerado por " _
        & Conn.SrcErro & Chr(13) & Conn.DescErro, vbCritical, "Falha de Leitura"
End Sub

Private Sub tmrControl_Timer()
    tmrControl.Enabled = False
    stbControl.Panels.Item(1).Text = Empty
End Sub

Private Sub txtBackup_GotFocus()
    txtBackup.SelStart = 0
    txtBackup.SelLength = 255
    txtBackup.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtBackup_LostFocus()
    If txtBackup.Text = Empty Then txtBackup.Text = "C:"
    txtQtdeMesas.BackColor = &H80000005 'Branco
End Sub

Private Sub txtCodFunc_Change()
    If txtCodFunc.Text = Empty Then GoTo Limpa
    Call Conn.Query("SELECT nome FROM funcionario" _
        & " WHERE id = " & txtCodFunc.Text)
    If Conn.Rs.RecordCount = 0 Then
        txtNomeFunc.Text = "Registro Inexistente."
    Else
        txtNomeFunc.Text = Conn.Rs.Fields.Item("nome").Value
    End If
    Call Conn.Query("SELECT id, campo, valor FROM config")
    Exit Sub
Limpa:
    txtNomeFunc.Text = Empty
End Sub

Private Sub txtCodFunc_GotFocus()
    txtCodFunc.SelStart = 0
    txtCodFunc.SelLength = 11
    txtCodFunc.BackColor = &H80000018 'Amarelo
    stbControl.Panels.Item(1).Text = "Pressione F5 para Localizar"
    LocalizarOQue = "Funcionario"
End Sub

Private Sub txtCodFunc_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    KeyAscii = modGetKeyAscii.Numeros(KeyAscii)
End Sub

Private Sub txtCodFunc_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not KeyCode = vbKeyF5 Then Exit Sub
    Call Conn.Query("SELECT id, nome, cpf, situ, recvale, limite FROM funcionario")
    If Conn.Rs.RecordCount = 0 Then
        MsgBox "Não Há Dados Registrados", vbInformation, "Sem Registros"
        Call Conn.Query("SELECT id, campo, valor FROM config")
        Exit Sub
    End If
    Set frmPrin.frmPai = Me
    frmLocalizar.Show vbModal, Me
    If Localizar = True Then
        txtCodFunc.Text = Conn.Rs.Fields.Item("id").Value
    End If
    Call Conn.Query("SELECT id, campo, valor FROM config")
End Sub

Private Sub txtCodFunc_LostFocus()
    txtCodFunc.BackColor = &H80000005 'Branco
    stbControl.Panels.Item(1).Text = ""
    LocalizarOQue = "Control"
End Sub

Private Sub txtMesaFechada_GotFocus()
    txtMesaFechada.SelStart = 0
    txtMesaFechada.SelLength = 4
    txtMesaFechada.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtMesaFechada_LostFocus()
    txtMesaFechada.BackColor = &H80000005 'Branco
End Sub

Private Sub txtQtdeMesas_GotFocus()
    txtQtdeMesas.SelStart = 0
    txtQtdeMesas.SelLength = 4
    txtQtdeMesas.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtQtdeMesas_LostFocus()
    txtQtdeMesas.BackColor = &H80000005 'Branco
End Sub

Private Sub txtTaxa_GotFocus()
    txtTaxa.SelStart = 0
    txtTaxa.SelLength = 4
    txtTaxa.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtTaxa_LostFocus()
    txtTaxa.BackColor = &H80000005 'Branco
End Sub
