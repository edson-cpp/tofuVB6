VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmFrente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frente de Vendas ao Cliente"
   ClientHeight    =   5745
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9855
   Icon            =   "frmFrente.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   383
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   657
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlFreD 
      Left            =   4920
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFrente.frx":0CCA
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFrente.frx":19A4
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFrente.frx":267E
            Key             =   "salvar"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrFrente 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5520
      Top             =   4320
   End
   Begin MSComctlLib.ImageList imlFrente 
      Left            =   4320
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   -2147483643
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFrente.frx":3358
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFrente.frx":4032
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFrente.frx":4D0C
            Key             =   "salvar"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraProduto 
      Caption         =   "Produtos"
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   9615
      Begin VB.TextBox txtParcial 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtCodPro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtDescPro 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtQtdePro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtVlrPro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid mfgFrente 
         Height          =   1815
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3201
         _Version        =   393216
         FixedCols       =   0
         Enabled         =   0   'False
         ScrollBars      =   2
         Appearance      =   0
      End
      Begin VB.Label labParcial 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Parcial:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label labPreco 
         AutoSize        =   -1  'True
         Caption         =   "Preço"
         Height          =   195
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   420
      End
      Begin VB.Label labQtde 
         AutoSize        =   -1  'True
         Caption         =   "Qtde"
         Height          =   195
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   345
      End
      Begin VB.Label labDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   720
      End
      Begin VB.Label labCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Cliente"
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   9615
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   840
         MaxLength       =   10
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1440
         Width           =   4695
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         MaxLength       =   60
         TabIndex        =   16
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox txtFone 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   3000
         MaxLength       =   32
         TabIndex        =   15
         Top             =   1080
         Width           =   1335
      End
      Begin MSMask.MaskEdBox mebCpf 
         Height          =   300
         Left            =   840
         TabIndex        =   19
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   14
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin VB.Label labCodCli 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   420
         Width           =   540
      End
      Begin VB.Label labCpf 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CPF:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1140
         Width           =   345
      End
      Begin VB.Label labEmail 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   1500
         Width           =   465
      End
      Begin VB.Label labNome 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   780
         Width           =   555
      End
      Begin VB.Label labFone 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fone:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2520
         TabIndex        =   20
         Top             =   1140
         Width           =   405
      End
      Begin VB.Label labStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   840
         TabIndex        =   14
         Top             =   1800
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin MSComctlLib.StatusBar stbFrente 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   5400
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12488
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1852
            MinWidth        =   1852
            TextSave        =   "5/15/2019"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "10:54 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1852
            MinWidth        =   1852
            Text            =   "F1 - Ajuda"
            TextSave        =   "F1 - Ajuda"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrFrente 
      Align           =   1  'Align Top
      DragMode        =   1  'Automatic
      Height          =   630
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   1111
      BandCount       =   1
      _CBWidth        =   9855
      _CBHeight       =   630
      _Version        =   "6.0.8169"
      Child1          =   "tbrFrente"
      MinHeight1      =   38
      Width1          =   554
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrFrente 
         Height          =   570
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlFrente"
         DisabledImageList=   "imlFreD"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "salvar"
               Object.ToolTipText     =   "Salvar"
               ImageKey        =   "salvar"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cancelar"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "cancelar"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sair"
               Object.ToolTipText     =   "Sair"
               ImageKey        =   "sair"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mniSalvar 
         Caption         =   "&Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mniCancelar 
         Caption         =   "C&ancelar"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mniSepUm 
         Caption         =   "-"
      End
      Begin VB.Menu mniSair 
         Caption         =   "Sai&r"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "Aj&uda"
      Begin VB.Menu mniConteudo 
         Caption         =   "&Conteúdo"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmFrente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Conn As clsMyConnect
Dim LocalizarOQue As String
Public Localizar As Boolean 'Define o retorno da tela localizar
Dim Sair As Boolean
Dim TaxaServ As Single
Dim Total As Single

Private Sub Form_Load()
    Dim I As Integer
    Set Conn = New clsMyConnect
    Call Conn.Connect
    If Conn.NumErro <> 0 Then GoTo SubFail
    With mfgFrente
        .Cols = 4
        .Rows = 1
        .Clear
        .ColAlignment(3) = 1
        .ColWidth(0) = 975
        .ColWidth(1) = 3210
        .ColWidth(2) = 855
        .ColWidth(3) = 1095
        .TextArray(0) = "Código"
        .TextArray(1) = "Descrição"
        .TextArray(2) = "Qtde"
        .TextArray(3) = "Preço"
    End With
    Exit Sub
SubFail:
    MsgBox "Não foi Possível Conectar-se com a Base de Dados." _
        & Chr(13) & "Erro # " & Str(Conn.NumErro) & " foi gerado por " _
        & Conn.SrcErro & Chr(13) & Conn.DescErro, vbCritical, "Falha de Conexão"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Conn.Disconnect
End Sub

Private Sub mfgFrente_GotFocus()
    stbFrente.Panels.Item(1).Text = "Pressione Del para Subtrair 1, Shift+Del para Excluir o item, Esc para Cancelar"
End Sub

Private Sub mfgFrente_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And Shift = 1 Then
        GoTo Exclui
    ElseIf KeyCode = vbKeyDelete And Shift = 0 Then
        If mfgFrente.TextMatrix(mfgFrente.Row, 2) = 1 Then
            GoTo Exclui
        Else
            mfgFrente.TextMatrix(mfgFrente.Row, 2) = _
                CInt(mfgFrente.TextMatrix(mfgFrente.Row, 2)) - 1
        End If
        GoTo Finaliza
    End If
    Exit Sub
Exclui:
    If mfgFrente.Rows = 2 Then
        mfgFrente.Rows = 1
        mfgFrente.Enabled = False
        txtCodPro.SetFocus
    Else
        mfgFrente.RemoveItem mfgFrente.Row
    End If
Finaliza:
    mfgFrente.Tag = "Alterado"
    Call CalculaParcial
End Sub

Private Sub mfgFrente_LostFocus()
    stbFrente.Panels.Item(1).Text = "Pressione Esc para Cancelar"
End Sub

Private Sub mniCancelar_Click()
    'If lvwMesas.Tag = Empty Then
    '    Call HabCampos(False)
    '    stbFrente.Panels.Item(1).Text = "Pressione Enter para cancelar, Esc para Cancelar"
    '    Me.Tag = "Cancelar"
    '    Me.Caption = Me.Caption & " - Cancelando"
    'Else
    '    frmPrin.Auth = False
    '    frmPrin.TituloSenha = "Senha de Usuário"
    '    frmSenha.Show vbModal, Me
    '    frmPrin.TituloSenha = "Senha de Acesso"
    '    If Not frmPrin.Auth Then
    '        Exit Sub
    '    End If
    '    frmPrin.Auth = False
    '    frmPrin.TituloSenha = "Senha de Supervisor"
    '    frmSenha.Show vbModal, Me
    '    frmPrin.TituloSenha = "Senha de Acesso"
    '    If frmPrin.Auth And frmPrin.NivelUsuarioLogado <= 2 Then
    '        MsgBox "Cancelado com Êxito", vbInformation, "Concluído!"
    '    Else
    '        MsgBox "Necessário Autorização de Supervisor", _
    '            vbExclamation, "Autorização"
    '    End If
    'End If
End Sub

Private Sub mniSair_Click()
    Unload Me
End Sub

Private Sub mniSalvar_Click()
    Dim Frente As Long
    Dim I As Integer
    Dim dinh As String
    Dim cheq As String
    Dim cart As String
    Dim tick As String
    Dim desc As String
    If mfgFrente.Rows = 1 Then
        MsgBox "Por favor informe pelo menos" & Chr(13) & _
            "um item para efetuar a venda", vbExclamation, "Venda Vazia"
        Exit Sub
    End If
    dinh = frmFechaVenda.txtDinheiro.Text
    cheq = frmFechaVenda.txtCheque.Text
    cart = frmFechaVenda.txtCartao.Text
    tick = frmFechaVenda.txtTicket.Text
    desc = frmFechaVenda.txtDesconto.Text
    Call Conn.Query("INSERT INTO venda ( " _
        & "id, cid, dataVen, dinheiro, cheque, cartao, ticket, desconto ) " _
        & "VALUES ( null, " _
        & txtCodigo.Text & ", " _
        & Date & ", " _
        & Replace(CDbl(IIf(dinh = "", 0, dinh)), ",", ".") & ", " _
        & Replace(CDbl(IIf(cheq = "", 0, cheq)), ",", ".") & ", " _
        & Replace(CDbl(IIf(cart = "", 0, cart)), ",", ".") & ", " _
        & Replace(CDbl(IIf(tick = "", 0, tick)), ",", ".") & ", " _
        & Replace(CDbl(IIf(desc = "", 0, desc)), ",", ".") & " )")
    Frente = Conn.GetValue("SELECT MAX(id) FROM venda")
    For I = 1 To mfgFrente.Rows - 1
        Call Conn.Query("INSERT INTO vendaItem ( " _
            & "id, pid, vid, qtde, preco ) " _
            & "VALUES ( null, " _
            & txtCodigo.Text & ", " _
            & Frente & ", " _
            & txtQtdePro.Text & ", " _
            & txtVlrPro.Text & " )")
    Next
    txtParcial.Text = Empty
    txtCodPro.Text = Empty
    txtVlrPro.Text = Empty
    txtQtdePro.Text = Empty
    txtDescPro.Text = Empty
    mfgFrente.Enabled = False
    mfgFrente.Rows = 1
    stbFrente.Panels.Item(1).Text = ""
    Call HabCampos(True)
    'imgSalvar.Enabled = False
    'imgSalvar.Visible = False
    txtParcial.DataChanged = False
    txtCodPro.DataChanged = False
    txtQtdePro.DataChanged = False
    txtVlrPro.DataChanged = False
    txtDescPro.DataChanged = False
    mfgFrente.Tag = Empty
    txtCodigo.SetFocus
End Sub

Private Sub tbrFrente_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "consultar"
            'mniConsultar_Click
        Case "cancelar"
            mniCancelar_Click
        Case "transferir"
            'mniTransfer_Click
        Case "estornar"
            'mniEstornar_Click
        Case "fechar"
            'mniFechar_Click
        Case "pagar"
            'mniPagar_Click
        Case "sair"
            mniSair_Click
    End Select
End Sub

Private Sub tmrFrente_Timer()
    tmrFrente.Enabled = False
    stbFrente.Panels.Item(1).Text = Empty
End Sub

Public Sub Load()
    Select Case LocalizarOQue
        Case "Cliente"
            Call LoadCliente
        Case "Produto"
            Call LoadProduto
    End Select
End Sub

' cmpProc = Campo Procurado
Public Sub Campo(txt As String)
    Select Case LocalizarOQue
        Case "Cliente"
            Call CampoCliente(txt)
        Case "Produto"
            Call CampoProduto(txt)
    End Select
End Sub

Public Sub LoadCliente()
    With frmLocalizar.cbbProcura
        .AddItem ("Código")
        .AddItem ("Nome")
        .AddItem ("CPF")
        .Text = "Nome"
    End With
    With frmLocalizar.mfgLocalizar
        .Cols = 3
        .Rows = 1
        .Clear
        .ColWidth(0) = 1000
        .ColWidth(1) = 4000
        .ColWidth(2) = 2100
        .TextArray(0) = "Código"
        .TextArray(1) = "Nome"
        .TextArray(2) = "CPF"
        Conn.Rs.MoveFirst
        Do While Not Conn.Rs.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Conn.Rs.Fields.Item("id").Value
            .TextMatrix(.Rows - 1, 1) = Conn.Rs.Fields.Item("nome").Value
            .TextMatrix(.Rows - 1, 2) = Conn.Rs.Fields.Item("cpf").Value
            Conn.Rs.MoveNext
        Loop
    End With
    frmLocalizar.Caption = "Localização de Clientes"
    frmLocalizar.tabela = "cliente"
End Sub

Public Sub CampoCliente(txt As String)
    Select Case txt
        Case "Código"
            frmLocalizar.Campo = "id"
        Case "Nome"
            frmLocalizar.Campo = "nome"
        Case "CPF"
            frmLocalizar.Campo = "cpf"
    End Select
End Sub

Public Sub LoadProduto()
    With frmLocalizar.cbbProcura
        .AddItem ("Código")
        .AddItem ("Descrição")
        .Text = "Descrição"
    End With
    With frmLocalizar.mfgLocalizar
        .Cols = 4
        .Rows = 1
        .Clear
        .ColWidth(0) = 800
        .ColWidth(1) = 3200
        .ColWidth(2) = 2100
        .ColWidth(3) = 1447
        .TextArray(0) = "Código"
        .TextArray(1) = "Descrição"
        .TextArray(2) = "Preço"
        .TextArray(3) = "Tipo"
        Conn.Rs.MoveFirst
        Do While Not Conn.Rs.EOF
            Dim x As Integer
            Dim y As String
            x = Conn.Rs.Fields.Item("tipo").Value
            y = IIf(x = 1, "Comida", "Bebida")
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Conn.Rs.Fields.Item("id").Value
            .TextMatrix(.Rows - 1, 1) = Conn.Rs.Fields.Item("descricao").Value
            .TextMatrix(.Rows - 1, 2) = modMoeda.FmtMoeda(Conn.Rs.Fields.Item("preco").Value)
            .TextMatrix(.Rows - 1, 3) = y
            Conn.Rs.MoveNext
        Loop
        frmLocalizar.Caption = "Localização de Produtos"
        frmLocalizar.tabela = "produto"
    End With
End Sub

Public Sub CampoProduto(txt As String)
    Select Case txt
        Case "Código"
            frmLocalizar.Campo = "id"
        Case "Descrição"
            frmLocalizar.Campo = "descricao"
    End Select
End Sub

Private Sub txtCodigo_Change()
    If Not txtCodigo.Text = Empty Then
        Call Conn.Query("SELECT id, nome, cpf, fone, email FROM cliente")
        Conn.Rs.Find "id = " & txtCodigo.Text
        If Conn.Rs.BOF Or Conn.Rs.EOF Then
            tmrFrente.Enabled = True
            stbFrente.Panels.Item(1).Text = "Registro Inexistente."
            txtCodPro.Enabled = False
        Else
            txtNome.Text = Conn.Rs.Fields.Item("nome").Value
            mebCpf.Text = Conn.Rs.Fields.Item("cpf").Value
            txtFone.Text = Conn.Rs.Fields.Item("fone").Value
            txtEmail.Text = Conn.Rs.Fields.Item("email").Value
            txtCodPro.Enabled = True
            If txtCodigo.Text = "1" Then
                txtCodPro.SetFocus
            End If
            Exit Sub
        End If
    End If
    txtNome.Text = Empty
    mebCpf.Text = Empty
    txtFone.Text = Empty
    txtEmail.Text = Empty
End Sub

Private Sub txtCodigo_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not KeyCode = vbKeyF5 Then Exit Sub
    Call Conn.Query("SELECT id, nome, cpf, email FROM cliente")
    If Conn.Rs.RecordCount = 0 Then
        MsgBox "Não Há Dados Registrados", vbInformation, "Sem Registros"
        Exit Sub
    End If
    Set frmPrin.frmPai = Me
    frmLocalizar.Show vbModal, Me
    If Localizar = True Then
        txtCodigo.Text = Conn.Rs.Fields.Item("id").Value
    End If
End Sub

Private Sub txtCodPro_Change()
    If txtCodPro.Text = Empty Then GoTo Limpa
    Call Conn.Query("SELECT produto.descricao, produto.preco" _
        & ", promocao.preco as promo FROM produto" _
        & " LEFT JOIN promocao ON produto.id = promocao.pid" _
        & " AND promocao.dia = " & Weekday(Date, vbSunday) _
        & " WHERE produto.id = " & txtCodPro.Text)
    If Conn.Rs.RecordCount = 0 Then
        txtVlrPro.Text = Empty
        txtQtdePro.Text = Empty
        txtDescPro.Text = "Registro Inexistente."
        txtQtdePro.Enabled = False
        txtVlrPro.Enabled = False
    Else
        If IsNull(Conn.Rs.Fields.Item("promo").Value) Then
            Conn.Rs.Fields.Item("promo").Value = _
                Conn.Rs.Fields.Item("preco").Value
        End If
        txtDescPro.Text = Conn.Rs.Fields.Item("descricao").Value
        txtVlrPro.Text = modMoeda.FmtMoeda(Conn.Rs.Fields.Item("promo").Value)
        txtQtdePro.Text = 1
        txtQtdePro.Enabled = True
        txtVlrPro.Enabled = True
    End If
    Exit Sub
Limpa:
    txtDescPro.Text = Empty
    txtVlrPro.Text = Empty
    txtVlrPro.Enabled = False
    txtQtdePro.Text = Empty
    txtQtdePro.Enabled = False
End Sub

Private Sub txtCodPro_GotFocus()
    txtCodPro.SelStart = 0
    txtCodPro.SelLength = 11
    txtCodPro.BackColor = &H80000018 'Amarelo
    stbFrente.Panels.Item(1).Text = "Pressione F5 para Localizar, Enter para inserir, Esc para Cancelar"
    LocalizarOQue = "Produto"
End Sub

Private Sub txtCodPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim I As Integer
        If txtDescPro.Text = "Registro Inexistente." Then
            MsgBox "Registro Inexistente.", vbInformation
            Exit Sub
        End If
        If txtCodPro.Text = Empty Then
            Call mniSalvar_Click
            Exit Sub
        End If
        With mfgFrente
            For I = 0 To .Rows - 1
                If .TextMatrix(I, 0) = txtCodPro.Text _
                    And .TextMatrix(I, 3) = txtVlrPro.Text Then
                    .TextMatrix(I, 2) = CInt(.TextMatrix(I, 2)) + CInt(txtQtdePro.Text)
                    txtCodPro.Text = Empty
                    txtDescPro.Text = Empty
                    txtVlrPro.Text = Empty
                    txtQtdePro.Text = Empty
                    Call CalculaParcial
                    Exit Sub
                End If
            Next
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = txtCodPro.Text
            .TextMatrix(.Rows - 1, 1) = txtDescPro.Text
            .TextMatrix(.Rows - 1, 2) = txtQtdePro.Text
            .TextMatrix(.Rows - 1, 3) = modMoeda.FmtMoeda(txtVlrPro.Text)
            mfgFrente.Enabled = True
            txtCodPro.Text = Empty
            txtDescPro.Text = Empty
            txtVlrPro.Text = Empty
            txtQtdePro.Text = Empty
        End With
        txtCodPro.SelStart = 0
        txtCodPro.SelLength = 11
        Call CalculaParcial
        Exit Sub
    End If
    KeyAscii = modGetKeyAscii.Numeros(KeyAscii)
End Sub

Private Sub txtCodPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not KeyCode = vbKeyF5 Then Exit Sub
    Call Conn.Query("SELECT id, descricao, preco, tipo, divpessoa FROM produto")
    If Conn.Rs.RecordCount = 0 Then
        MsgBox "Não Há Dados Registrados", vbInformation, "Sem Registros"
        Exit Sub
    End If
    Set frmPrin.frmPai = Me
    frmLocalizar.Show vbModal, Me
    If Localizar = True Then
        txtCodPro.Text = Conn.Rs.Fields.Item("id").Value
    End If
End Sub

Private Sub txtCodPro_LostFocus()
    txtCodPro.BackColor = &H80000005 'Branco
    stbFrente.Panels.Item(1).Text = "Pressione Esc para Cancelar"
    LocalizarOQue = "Cliente"
End Sub

Private Sub txtParcial_GotFocus()
    txtParcial.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtParcial_LostFocus()
    txtParcial.BackColor = &H80000005 'Branco
End Sub

Private Sub HabCampos(situ As Boolean)
    txtCodPro.Enabled = situ
    labStatus.Visible = Not situ
    mniCancelar.Enabled = situ
    'mniConsultar.Enabled = situ
    mniConteudo.Enabled = situ
    'mniEstornar.Enabled = situ
    'mniFechar.Enabled = situ
    'mniMenu.Enabled = situ
    'mniPagar.Enabled = situ
    'mniPedido.Enabled = situ
    mniSair.Enabled = situ
    'mniTransfer.Enabled = situ
    tbrFrente.Buttons.Item("transferir").Enabled = situ
    tbrFrente.Buttons.Item("estornar").Enabled = situ
    tbrFrente.Buttons.Item("pagar").Enabled = situ
    tbrFrente.Buttons.Item("sair").Enabled = situ
    tbrFrente.Buttons.Item("consultar").Enabled = situ
    tbrFrente.Buttons.Item("fechar").Enabled = situ
    tbrFrente.Buttons.Item("cancelar").Enabled = situ
End Sub

Private Sub txtQtdePro_GotFocus()
    txtQtdePro.SelStart = 0
    txtQtdePro.SelLength = 10
    txtQtdePro.BackColor = &H80000018 'Amarelo
    stbFrente.Panels.Item(1).Text = "Pressione Enter para inserir, Esc para Cancelar"
End Sub

Private Sub txtQtdePro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call txtCodPro_KeyPress(KeyAscii)
        txtQtdePro.SelStart = 0
        txtQtdePro.SelLength = 10
        txtCodPro.SetFocus
    End If
    KeyAscii = modGetKeyAscii.Numeros(KeyAscii)
End Sub

Private Sub txtQtdePro_LostFocus()
    txtQtdePro.BackColor = &H80000005 'Branco
    stbFrente.Panels.Item(1).Text = "Pressione Esc para Cancelar"
End Sub

Private Sub txtVlrPro_GotFocus()
    txtVlrPro.SelStart = 0
    txtVlrPro.SelLength = 10
    txtVlrPro.BackColor = &H80000018 'Amarelo
    stbFrente.Panels.Item(1).Text = "Pressione Enter para inserir, Esc para Cancelar"
End Sub

Private Sub txtVlrPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtVlrPro.Text = modMoeda.FmtMoeda(txtVlrPro.Text)
        Call txtCodPro_KeyPress(KeyAscii)
        txtVlrPro.SelStart = 0
        txtVlrPro.SelLength = 10
        txtCodPro.SetFocus
    End If
    KeyAscii = modGetKeyAscii.Moeda(KeyAscii)
End Sub

Private Sub txtVlrPro_LostFocus()
    txtVlrPro.Text = modMoeda.FmtMoeda(txtVlrPro.Text)
    txtVlrPro.BackColor = &H80000005 'Branco
    stbFrente.Panels.Item(1).Text = "Pressione Esc para Cancelar"
End Sub

Private Sub CalculaParcial()
    Call Conn.Query("SELECT valor FROM config WHERE campo = 'Taxa';")
    Dim I As Integer
    Dim Valor As Integer
    TaxaServ = Empty
    For I = 1 To mfgFrente.Rows - 1
        Total = Total + CSng(mfgFrente.TextMatrix(I, 3)) * mfgFrente.TextMatrix(I, 2)
    Next
    Valor = Val(Conn.Rs.Fields.Item("valor").Value)
    txtParcial.Text = Format(Total * ((Valor / 100) + 1), "R$ ###,##0.00")
    Total = Format(Total, "R$ ###,##0.00")
    TaxaServ = (Total * ((Valor / 100) + 1)) - Total
    If txtParcial.Text = "R$ 0,00" Then txtParcial.Text = Empty
End Sub

Private Sub CarregaValores(QualSub As String)
    'txtCodFunc.Text = Conn.Rs.Fields.Item("fid").Value
    'Call Conn.Query("SELECT mesaItem.pid, mesaItem.qtde" _
    '    & ", mesaItem.preco, produto.descricao " _
    '    & "FROM mesaItem " _
    '    & "LEFT JOIN produto ON produto.id = mesaItem.pid " _
    '    & "WHERE mid = " & MesaId)
    mfgFrente.Rows = 1
    Do While Not Conn.Rs.EOF
        With mfgFrente
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Conn.Rs.Fields.Item("pid").Value
            .TextMatrix(.Rows - 1, 1) = Conn.Rs.Fields.Item("descricao").Value
            .TextMatrix(.Rows - 1, 2) = Conn.Rs.Fields.Item("qtde").Value
            .TextMatrix(.Rows - 1, 3) = Format(Conn.Rs.Fields.Item("preco").Value, "R$ ###,##0.00")
        End With
        Conn.Rs.MoveNext
    Loop
End Sub

Private Property Get Pagar(KeyAscii As Integer, ByRef SrcField As Object, ByRef DestField As Object)
    Pagar = KeyAscii
    If KeyAscii = vbKeyReturn Then
        If SrcField.Text = Empty Then
            'If CSng(Mid(stbMesas.Panels.Item(1).Text, 10))
            '    - CSng(IIf(txtDesconto.Text = Empty, _
            '    0, txtDesconto.Text)) = 0 Then
            If True Then
                Dim MesaId As Integer
                Dim Taxa As Integer
                Dim I As Integer
                Dim VendaId As Integer
                'MesaId = Conn.GetValue("SELECT id FROM mesa WHERE mesa = '" _
                '    & lvwMesas.SelectedItem.Text & "'")
                Taxa = Conn.GetValue("SELECT valor FROM config WHERE " _
                    & "campo = 'Taxa'")
                Call Conn.Query("DELETE FROM mesaItem WHERE mid = " & MesaId)
                Call Conn.Query("DELETE FROM mesa WHERE id = " & MesaId)
                'If txtSubDinheiro.Text = Empty Then txtSubDinheiro.Text = "0"
                'If txtSubCheque.Text = Empty Then txtSubCheque.Text = "0"
                'If txtSubCartao.Text = Empty Then txtSubCartao.Text = "0"
                'If txtSubTicket.Text = Empty Then txtSubTicket.Text = "0"
                'If txtDesconto.Text = Empty Then txtDesconto.Text = "0"
                'Call Conn.Query("INSERT INTO venda ( id, fid, dataVen" _
                '    & ", horaVen, taxa, mesa, dinheiro, cheque, cartao" _
                '    & ", ticket, desconto ) VALUES ( null" _
                '    & ", " & txtCodFunc.Text _
                '    & ", '" & Replace(Date, "/", "") & "'" _
                '    & ", '" & Replace(Time, ":", "") & "'" _
                '    & ", " & Taxa _
                '    & ", '" & lvwMesas.SelectedItem.Text & "'" _
                '    & ", " & Replace(CSng(txtSubDinheiro.Text), ",", ".") _
                '    & ", " & Replace(CSng(txtSubCheque.Text), ",", ".") _
                '    & ", " & Replace(CSng(txtSubCartao.Text), ",", ".") _
                '    & ", " & Replace(CSng(txtSubTicket.Text), ",", ".") _
                '    & ", " & Replace(CSng(txtDesconto.Text), ",", ".") & ")")
                VendaId = Conn.GetValue("SELECT MAX(id) FROM venda")
                'For I = 1 To mfgMesas.Rows - 1
                '    Call Conn.Query("INSERT INTO vendaItem ( " _
                '        & "id, pid, vid, qtde, preco ) " _
                '        & "VALUES ( null, " _
                '        & mfgMesas.TextMatrix(I, 0) & ", " _
                '        & VendaId & ", " _
                '        & mfgMesas.TextMatrix(I, 2) & ", " _
                '        & Replace(CDbl(mfgMesas.TextMatrix(I, 3)), ",", ".") _
                '        & " )")
                'Next
                'lvwMesas.SelectedItem.Icon = _
                '    imlListView.ListImages.Item("azul").Index
                'imgCancelar_Click
            End If
        ElseIf SrcField.Text = "-" Then
            DestField.Text = Empty
            SrcField.Text = Empty
            'stbMesas.Panels.Item(1).Text = "Restante: " _
            '    & Format(CSng(txtTotal.Text) - _
            '    CSng(IIf(txtSubDinheiro.Text = Empty, 0, txtSubDinheiro.Text)) - _
            '    CSng(IIf(txtSubCheque.Text = Empty, 0, txtSubCheque.Text)) - _
            '    CSng(IIf(txtSubCartao.Text = Empty, 0, txtSubCartao.Text)) - _
            '    CSng(IIf(txtSubTicket.Text = Empty, 0, txtSubTicket.Text)), _
            '    "R$ ###,##0.00")
        Else
            DestField.Text = _
                Format(CSng(SrcField.Text) + _
                CSng(IIf(DestField.Text = Empty, 0, _
                DestField.Text)), "R$ ###,##0.00")
            SrcField.Text = Empty
            'stbMesas.Panels.Item(1).Text = "Restante: " _
            '    & Format(CSng(txtTotal.Text) - _
            '    CSng(IIf(txtSubDinheiro.Text = Empty, 0, txtSubDinheiro.Text)) - _
            '    CSng(IIf(txtSubCheque.Text = Empty, 0, txtSubCheque.Text)) - _
            '    CSng(IIf(txtSubCartao.Text = Empty, 0, txtSubCartao.Text)) - _
            '    CSng(IIf(txtSubTicket.Text = Empty, 0, txtSubTicket.Text)), _
            '    "R$ ###,##0.00")
        End If
        Exit Property
    ElseIf KeyAscii = 45 Then ' -
        If Not SrcField.Text = Empty Then
            Pagar = Empty
        End If
        Exit Property
    End If
    Pagar = modGetKeyAscii.Moeda(KeyAscii)
End Property
