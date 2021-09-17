VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmPrin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tofu - Sistema de Restaurante"
   ClientHeight    =   2415
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7590
   Icon            =   "frmPrin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   506
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPrin 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   1560
   End
   Begin VB.PictureBox ticPrin 
      Height          =   480
      Left            =   840
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   3
      Top             =   1440
      Width           =   1200
   End
   Begin MSComctlLib.ImageList imlPrin 
      Left            =   120
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   -2147483643
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrin.frx":08CA
            Key             =   "control"
            Object.Tag             =   "control"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrin.frx":0F1A
            Key             =   "backup"
            Object.Tag             =   "backup"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrin.frx":143C
            Key             =   "funcionario"
            Object.Tag             =   "funcionario"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrin.frx":1A31
            Key             =   "produto"
            Object.Tag             =   "produto"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrin.frx":270B
            Key             =   "vale"
            Object.Tag             =   "vale"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrin.frx":2BE2
            Key             =   "saida"
            Object.Tag             =   "saida"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrin.frx":314E
            Key             =   "promocao"
            Object.Tag             =   "promocao"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrin.frx":3E28
            Key             =   "medalha"
            Object.Tag             =   "medalha"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrin.frx":4B02
            Key             =   "sair"
            Object.Tag             =   "sair"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrin.frx":57DC
            Key             =   "relatorio"
            Object.Tag             =   "relatorio"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrin.frx":64B6
            Key             =   "usuario"
            Object.Tag             =   "usuario"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrin.frx":7190
            Key             =   "mesa"
            Object.Tag             =   "mesa"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrin.frx":7E6A
            Key             =   "cliente"
            Object.Tag             =   "cliente"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrPrin 
      Align           =   1  'Align Top
      DragMode        =   1  'Automatic
      Height          =   630
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   1111
      BandCount       =   1
      ImageList       =   "imlPrin"
      _CBWidth        =   7590
      _CBHeight       =   630
      _Version        =   "6.0.8169"
      Child1          =   "tbrPrin"
      MinHeight1      =   38
      Width1          =   554
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrPrin 
         Height          =   570
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlPrin"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "mesa"
               Object.ToolTipText     =   "Controle de Mesas"
               ImageKey        =   "mesa"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "control"
               Object.ToolTipText     =   "Centro de Controle"
               ImageKey        =   "control"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "backup"
               Object.ToolTipText     =   "Cópia de Segurança"
               ImageKey        =   "backup"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "funcionario"
               Object.ToolTipText     =   "Cadastro de Funcionários"
               ImageKey        =   "funcionario"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "produto"
               Object.ToolTipText     =   "Cadastro de Produtos"
               ImageKey        =   "produto"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "usuario"
               Object.ToolTipText     =   "Cadastro de Usuários"
               ImageKey        =   "usuario"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "promocao"
               Object.ToolTipText     =   "Cadastro de Promoções"
               ImageKey        =   "promocao"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "medalha"
               Object.ToolTipText     =   "Cadastro de Bonificação"
               ImageKey        =   "medalha"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cliente"
               Object.ToolTipText     =   "Cadastro de Clientes"
               ImageKey        =   "cliente"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "vale"
               Object.ToolTipText     =   "Cadastro de Vales"
               ImageKey        =   "vale"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "saida"
               Object.ToolTipText     =   "Saída Fornecedor / Serviços"
               ImageKey        =   "saida"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sair"
               Object.ToolTipText     =   "Sair do Sistema"
               ImageKey        =   "sair"
            EndProperty
         EndProperty
         MouseIcon       =   "frmPrin.frx":8B44
      End
   End
   Begin MSComctlLib.StatusBar stbPrin 
      Align           =   2  'Align Bottom
      DragMode        =   1  'Automatic
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   2070
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   609
      SimpleText      =   "Teste"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10372
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1376
            MinWidth        =   794
            TextSave        =   "10:41 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1535
            MinWidth        =   794
            TextSave        =   "5/15/2019"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mniMesas 
         Caption         =   "Controle de &Mesas"
         Shortcut        =   ^M
      End
      Begin VB.Menu mniSepDoze 
         Caption         =   "-"
      End
      Begin VB.Menu mniAbrir 
         Caption         =   "&Abrir Período"
         Shortcut        =   ^A
      End
      Begin VB.Menu mniFechar 
         Caption         =   "&Fechar Período"
         Shortcut        =   ^F
      End
      Begin VB.Menu mniSepOnze 
         Caption         =   "-"
      End
      Begin VB.Menu mniControl 
         Caption         =   "&Centro de Controle"
         Shortcut        =   ^E
      End
      Begin VB.Menu mniBackup 
         Caption         =   "Cópia de &Segurança"
         Shortcut        =   ^S
      End
      Begin VB.Menu mniSepDois 
         Caption         =   "-"
      End
      Begin VB.Menu mniSair 
         Caption         =   "Sai&r do Sistema"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastro"
      Begin VB.Menu mniFuncionarios 
         Caption         =   "&Funcionários"
      End
      Begin VB.Menu mniProdutos 
         Caption         =   "&Produtos"
      End
      Begin VB.Menu mniUsuario 
         Caption         =   "&Usuários"
      End
      Begin VB.Menu mniFornecedor 
         Caption         =   "F&ornecedores"
      End
      Begin VB.Menu mniSepUm 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCaixa 
         Caption         =   "&Caixa"
         Begin VB.Menu mniVales 
            Caption         =   "&Vales"
         End
         Begin VB.Menu mniSaidaForn 
            Caption         =   "&Saída a Forn./Serviços"
         End
      End
      Begin VB.Menu mniSepTres 
         Caption         =   "-"
      End
      Begin VB.Menu mniClientes 
         Caption         =   "Clie&ntes"
      End
      Begin VB.Menu mniPromocoes 
         Caption         =   "Pro&moções"
      End
      Begin VB.Menu mniBonificacao 
         Caption         =   "&Bonificação"
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Begin VB.Menu mnuRelCaixa 
         Caption         =   "&Caixa"
         Begin VB.Menu mniEntrada 
            Caption         =   "&Entrada de Suprimento"
         End
         Begin VB.Menu mniSaida 
            Caption         =   "S&aída de Suprimento"
         End
         Begin VB.Menu mniSepCinco 
            Caption         =   "-"
         End
         Begin VB.Menu mniRepasse 
            Caption         =   "&Repasse p/ Tesouraria"
         End
         Begin VB.Menu mniSepSeis 
            Caption         =   "-"
         End
         Begin VB.Menu mniBalancete 
            Caption         =   "&Balancete"
         End
         Begin VB.Menu mniParcial 
            Caption         =   "&Fechamento Parcial"
         End
         Begin VB.Menu mniContab 
            Caption         =   "C&ontabilização"
         End
         Begin VB.Menu mniCaixa 
            Caption         =   "Fechamento de &Caixa"
         End
      End
      Begin VB.Menu mniSepDez 
         Caption         =   "-"
      End
      Begin VB.Menu mniRelFuncionarios 
         Caption         =   "&Funcionários"
      End
      Begin VB.Menu mniRelProdutos 
         Caption         =   "&Produtos"
      End
      Begin VB.Menu mniRelPromocoes 
         Caption         =   "Pr&omoções"
      End
      Begin VB.Menu mniSepSete 
         Caption         =   "-"
      End
      Begin VB.Menu mniRelCap 
         Caption         =   "Co&ntas a Pagar"
      End
      Begin VB.Menu mniSepOito 
         Caption         =   "-"
      End
      Begin VB.Menu mniRelVales 
         Caption         =   "&Vales"
      End
      Begin VB.Menu mniRelValesGeral 
         Caption         =   "V&ales Geral"
      End
      Begin VB.Menu mniSepNove 
         Caption         =   "-"
      End
      Begin VB.Menu mniRelSaida 
         Caption         =   "&Saída a Fornecedor / Serviços"
      End
      Begin VB.Menu mniRelFechamento 
         Caption         =   "F&echamento de Período"
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "Aj&uda"
      Begin VB.Menu mniConteudo 
         Caption         =   "&Conteúdo"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mniSobre 
         Caption         =   "&Sobre"
      End
   End
End
Attribute VB_Name = "frmPrin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sair As Boolean
Public Auth As Boolean
Public frmPai As Form
Public Conn As clsMyConnect
Public CodigoUsuarioLogado As Long
Public NomeUsuarioLogado As String
Public LoginUsuarioLogado As String
Public NivelUsuarioLogado As Integer
Public TituloSenha As String
Public TipoVenda As String
Private WithEvents Conex As ADODB.Connection
Attribute Conex.VB_VarHelpID = -1
Private WithEvents Rs As ADODB.Recordset
Attribute Rs.VB_VarHelpID = -1
Private NumErro As Long
Private DescErro As String
Private SrcErro As String
Private Declare Function WaitForSingleObject Lib "kernel32" _
   (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" _
   (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
Private Const INFINITE = -1&
Private Const SYNCHRONIZE = &H100000
Private iTask As Long
Private ret As Long
Private pHandle As Long
Private MySQLDir As String
Private MySQLCommand As String

Private Sub Form_Initialize()
    Set Conn = New clsMyConnect
    Call Conn.Connect
    If Conn.NumErro = -2147467259 Then 'Não foi possível conectar-se
        If Environ("OS") = "Windows_NT" Then
            iTask = Shell("NET START MYSQL", vbHide)
            pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
            ret = WaitForSingleObject(pHandle, INFINITE)
            ret = CloseHandle(pHandle)
        Else
            MySQLDir = modFindFile.FindFile("C:\", "winmysqladmin.exe")
            Call Shell(MySQLDir & "\mysqld.exe", vbHide)
            'Faz a aplicação esperar 10 segundos
            Dim newHour As Integer
            Dim newMinute As Integer
            Dim newSecond As Integer
            Dim waitTime As String
            newHour = Hour(Now())
            newMinute = Minute(Now())
            newSecond = Second(Now()) + 10
            waitTime = TimeSerial(newHour, newMinute, newSecond)
            While Not waitTime = Time
            Wend
            '************************************
        End If
        Call Conn.Connect
        If Conn.NumErro = -2147467259 Then 'Não foi possível conectar-se
            Call CriaDB
        End If
    ElseIf Conn.NumErro = 0 Then
        TipoVenda = Conn.GetValue("SELECT valor FROM config WHERE campo = 'TipoVenda'")
        If TipoVenda = "1" Then
            mniMesas.Caption = "Frente de &Vendas ao Cliente"
            tbrPrin.Buttons.Item("mesa").ToolTipText = "Frente de &Vendas ao Cliente"
        End If
        Call Conn.Disconnect
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then 'ESC
        If tmrPrin.Enabled = False Then
            Me.WindowState = vbMinimized
            mniMesas_Click
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major _
        & "." & App.Minor & "." & App.Revision
    ticPrin.ToolTipText = Me.Caption
    CodigoUsuarioLogado = Empty
    NomeUsuarioLogado = Empty
    LoginUsuarioLogado = Empty
    NivelUsuarioLogado = Empty
    TituloSenha = "Senha de Acesso"
    Sair = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Sair Then
        Cancel = 0
    Else
        Cancel = 1
        frmPrin.Hide
        'ticPrin.FlashEnabled = True
        MsgBox "Tofu continua ativo no System Tray" & Chr(13) & "Ao lado do relógio do sistema", vbInformation
        'ticPrin.FlashEnabled = False
    End If
End Sub

Private Sub mniAbrir_Click()
    Auth = False
    frmSenha.Show vbModal, Me
    If Auth Then
        frmAbrir.Show vbModeless
    End If
End Sub

Private Sub mniBackup_Click()
    Auth = False
    frmSenha.Show vbModal, Me
    If Auth Then
        frmBackup.Show vbModeless
    End If
End Sub

Private Sub mniBonificacao_Click()
    Auth = False
    frmSenha.Show vbModal, Me
    If Auth Then
        frmBonificacao.Show vbModeless
    End If
End Sub

Private Sub mniClientes_Click()
    Auth = False
    frmSenha.Show vbModal, Me
    If Auth Then
        frmCliente.Show vbModeless
    End If
End Sub

Private Sub mniControl_Click()
    Auth = False
    frmSenha.Show vbModal, Me
    If Auth Then
        frmControl.Show vbModeless
    End If
End Sub

Private Sub mniFechar_Click()
    Auth = False
    frmSenha.Show vbModal, Me
    If Auth Then
        frmFechar.Show vbModeless
    End If
End Sub

Private Sub mniFornecedor_Click()
    Auth = False
    frmSenha.Show vbModal, Me
    If Auth Then
        frmFornecedor.Show vbModeless
    End If
End Sub

Private Sub mniFuncionarios_Click()
    Auth = False
    frmSenha.Show vbModal, Me
    If Auth Then
        frmFuncionario.Show vbModeless
    End If
End Sub

Private Sub mniMesas_Click()
    If TipoVenda = "0" Then
        frmMesa.Show
    Else
        frmFrente.Show
    End If
    'frmPrin.WindowState = vbMinimized
End Sub

Private Sub mniProdutos_Click()
    Auth = False
    frmSenha.Show vbModal, Me
    If Auth Then
        frmProduto.Show vbModeless
    End If
End Sub

Private Sub mniPromocoes_Click()
    Call Conn.Connect
    Call Conn.Query("SELECT id FROM produto")
    If Conn.Rs.RecordCount = 0 Then
        MsgBox "Não Há Produtos Registrados", vbInformation, "Sem Registros"
        Call Conn.Disconnect
        Exit Sub
    End If
    Call Conn.Disconnect
    Auth = False
    frmSenha.Show vbModal, Me
    If Auth Then
        frmPromocao.Show vbModeless
    End If
End Sub

Private Sub mniSaidaForn_Click()
    Auth = False
    frmSenha.Show vbModal, Me
    If Auth Then
        frmSaida.Show vbModeless
    End If
End Sub

Private Sub mniSair_Click()
    Sair = True
    End
End Sub

Private Sub mniUsuario_Click()
    Auth = False
    frmSenha.Show vbModal, Me
    If Auth Then
        If frmPrin.NivelUsuarioLogado = 3 Then
            MsgBox "Acesso Negado", vbExclamation, "Área Restrita"
        Else
            frmUsuario.Show vbModeless
        End If
    End If
End Sub

Private Sub mniVales_Click()
    Auth = False
    frmSenha.Show vbModal, Me
    If Auth Then
        frmVale.Show vbModeless
    End If
End Sub

Private Sub tbrPrin_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "mesa"
            mniMesas_Click
        Case "control"
            mniControl_Click
        Case "backup"
            mniBackup_Click
        Case "funcionario"
            mniFuncionarios_Click
        Case "produto"
            mniProdutos_Click
        Case "usuario"
            mniUsuario_Click
        Case "cliente"
            mniClientes_Click
        Case "promocao"
            mniPromocoes_Click
        Case "medalha"
            mniBonificacao_Click
        Case "vale"
            mniVales_Click
        Case "saida"
            mniSaidaForn_Click
        Case "sair"
            mniSair_Click
    End Select
End Sub

Private Sub ticPrin_LeftButtonDoubleClick()
    frmPrin.Show
End Sub

Private Sub CriaDB()
    Dim Arq As Integer
    Dim lineFile As String
    Dim sql As String
    Set Conex = New ADODB.Connection
    Set Rs = New ADODB.Recordset
    mnuArquivo.Visible = False
    mnuCadastro.Visible = False
    mnuRelatorios.Visible = False
    mnuAjuda.Visible = False
    stbPrin.Visible = False
    cbrPrin.Visible = False
    tbrPrin.Visible = False
    On Error GoTo ConFail
    Conex.ConnectionTimeout = 60
    Conex.CommandTimeout = 400
    Conex.CursorLocation = adUseClient
    Conex.Open "DRIVER={MySQL ODBC 3.51 Driver}" _
        & ";SERVER=localhost;UID=root;DATABASE=mysql"
    On Error GoTo 0
    NumErro = Empty
    Rs.CursorType = adOpenStatic
    Rs.CursorLocation = adUseClient
    Rs.LockType = adLockPessimistic
    Rs.ActiveConnection = Conex
    
    Rs.Open "DROP DATABASE IF EXISTS TEST;"
    Rs.Open "DELETE FROM mysql.user WHERE User='' AND Host = '%';"
    Rs.Open "DELETE FROM mysql.db WHERE User='' AND Host = '%';"
    Rs.Open "DELETE FROM mysql.tables_priv WHERE User='' AND Host = '%';"
    Rs.Open "DELETE FROM mysql.columns_priv WHERE User='' AND Host = '%';"
    Rs.Open "DELETE FROM mysql.user WHERE User='' AND Host = 'localhost';"
    Rs.Open "DELETE FROM mysql.db WHERE User='' AND Host = 'localhost';"
    Rs.Open "DELETE FROM mysql.tables_priv WHERE User='' AND Host = 'localhost';"
    Rs.Open "DELETE FROM mysql.columns_priv WHERE User='' AND Host = 'localhost';"
    Rs.Open "DELETE FROM mysql.user WHERE User='root' AND Host = '%';"
    Rs.Open "DELETE FROM mysql.db WHERE User='root' AND Host = '%';"
    Rs.Open "DELETE FROM mysql.tables_priv WHERE User='root' AND Host = '%';"
    Rs.Open "DELETE FROM mysql.columns_priv WHERE User='root' AND Host = '%';"
    Rs.Open "REVOKE ALL PRIVILEGES ON *.* FROM 'root'@'localhost';"
    Rs.Open "REVOKE GRANT OPTION ON *.* FROM 'root'@'localhost';"
    Rs.Open "DELETE FROM mysql.db WHERE User='root' AND Host = 'localhost';"
    Rs.Open "DELETE FROM mysql.tables_priv WHERE User='root' AND Host = 'localhost';"
    Rs.Open "DELETE FROM mysql.columns_priv WHERE User='root' AND Host = 'localhost';"
    Rs.Open "FLUSH PRIVILEGES;"
    Rs.Open "GRANT ALL PRIVILEGES ON *.* TO 'root'@'localhost' WITH GRANT OPTION;"
    Rs.Open "UPDATE mysql.user SET password='5c1fb21a20d15f82' WHERE User = 'root';"
    Rs.Open "GRANT ALL PRIVILEGES ON `tofu`.* TO 'arj'@'localhost';"
    Rs.Open "GRANT ALL PRIVILEGES ON `tofu`.* TO 'arj'@'%';"
    Rs.Open "UPDATE mysql.user SET password='359a15221ff9a9b7' WHERE User = 'arj';"
    Rs.Open "CREATE DATABASE IF NOT EXISTS `tofu`;"
    Rs.Open "USE tofu;"
    lineFile = Empty
    Arq = Empty
    sql = Empty
    ' Disponibiliza o próximo número de arquivo disponível
    Arq = FreeFile
    ' Abre Arquivo Para Leitura
    Open App.Path & "\tofu.sql" For Input As #Arq
    ' Lê o Arquivo Linha por Linha até o Fim
    Do While Not EOF(1)
        ' Lê a Linha Corrente
        Line Input #Arq, lineFile
        If Trim(lineFile) = Empty And _
            Not Trim(sql) = Empty Then
            Rs.Open sql
            sql = Empty
        Else
            sql = sql & Trim(lineFile) & " "
        End If
    Loop
    Close #Arq
    Rs.Open "INSERT INTO USUARIO ( ID, NOME, LOGIN, SENHA, CPF, NIVEL ) VALUES ( NULL, 'EDSON GONÇALVES DE AGUIAR', 'root', '*D8&c%Pp', '', 0 );"
    Rs.Open "INSERT INTO USUARIO ( ID, NOME, LOGIN, SENHA, CPF, NIVEL ) VALUES ( NULL, 'ADMINISTRADOR', 'admin', '(C8.^', '', 1 );"
    Rs.Open "INSERT INTO CONFIG (ID, CAMPO, VALOR ) VALUES ( NULL, 'QtdeMesas', '10' );"
    Rs.Open "INSERT INTO CONFIG (ID, CAMPO, VALOR ) VALUES ( NULL, 'Backup', '0C:' );"
    Rs.Open "INSERT INTO CONFIG (ID, CAMPO, VALOR ) VALUES ( NULL, 'Recibo', '0' );"
    Rs.Open "INSERT INTO CONFIG (ID, CAMPO, VALOR ) VALUES ( NULL, 'Taxa', '10' );"
    Rs.Open "INSERT INTO CONFIG (ID, CAMPO, VALOR ) VALUES ( NULL, 'MesaFechada', '10' );"
    Rs.Open "INSERT INTO CONFIG (ID, CAMPO, VALOR ) VALUES ( NULL, 'GarcomBalcao', '1' );"
    Rs.Open "INSERT INTO CONFIG (ID, CAMPO, VALOR ) VALUES ( NULL, 'TipoVenda', '0' );"
    Kill App.Path & "\tofu.sql"
    
    Conex.Close
    iTask = Shell(MySQLDir & "\mysqladmin -u root shutdown", vbHide)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    ret = WaitForSingleObject(pHandle, INFINITE)
    ret = CloseHandle(pHandle)
    
    'Faz a aplicação esperar 10 segundos
    Dim newHour As Integer
    Dim newMinute As Integer
    Dim newSecond As Integer
    Dim waitTime As String
    newHour = Hour(Now())
    newMinute = Minute(Now())
    newSecond = Second(Now()) + 10
    waitTime = TimeSerial(newHour, newMinute, newSecond)
    While Not waitTime = Time
    Wend
    '************************************
    
    If Environ("OS") = "Windows_NT" Then
        iTask = Shell("NET START MYSQL", vbHide)
        pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
        ret = WaitForSingleObject(pHandle, INFINITE)
        ret = CloseHandle(pHandle)
    Else
        Call Shell(MySQLDir & "\mysqld.exe", vbHide)
        'Faz a aplicação esperar 10 segundos
        newHour = Hour(Now())
        newMinute = Minute(Now())
        newSecond = Second(Now()) + 10
        waitTime = TimeSerial(newHour, newMinute, newSecond)
        While Not waitTime = Time
        Wend
        '************************************
    End If
    
    mnuArquivo.Visible = True
    mnuCadastro.Visible = True
    mnuRelatorios.Visible = True
    mnuAjuda.Visible = True
    stbPrin.Visible = True
    cbrPrin.Visible = True
    tbrPrin.Visible = True
    tmrPrin.Enabled = False
    Exit Sub
ConFail:
    NumErro = Err.Number
    DescErro = Err.Description
    SrcErro = Err.Source
    MsgBox "Não foi Possível Conectar-se com a Base de Dados." _
        & Chr(13) & "Erro # " & Str(NumErro) & " foi gerado por " _
        & SrcErro & Chr(13) & DescErro, vbCritical, "Falha de Conexão"
    On Error GoTo 0
End Sub

