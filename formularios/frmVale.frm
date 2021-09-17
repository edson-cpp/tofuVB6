VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmVale 
   Caption         =   "Cadastro de Vales"
   ClientHeight    =   4590
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7230
   Icon            =   "frmVale.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   306
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   482
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtVlrPro 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   2400
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame fraValores 
      Caption         =   "Valores do Funcionário"
      Height          =   1095
      Left            =   5040
      TabIndex        =   21
      Top             =   3000
      Width           =   2055
      Begin VB.Label labLimite 
         AutoSize        =   -1  'True
         Caption         =   "Limite:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   570
      End
      Begin VB.Label labVales 
         AutoSize        =   -1  'True
         Caption         =   "Vales:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   540
      End
      Begin VB.Label labSaldo 
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   555
      End
      Begin VB.Label labValLimite 
         AutoSize        =   -1  'True
         Caption         =   "R$ 0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   24
         Top             =   360
         Width           =   690
      End
      Begin VB.Label labValVales 
         AutoSize        =   -1  'True
         Caption         =   "R$ 0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   23
         Top             =   600
         Width           =   690
      End
      Begin VB.Label labValSaldo 
         AutoSize        =   -1  'True
         Caption         =   "R$ 0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   22
         Top             =   840
         Width           =   690
      End
   End
   Begin VB.TextBox txtNomeFunc 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   2400
      TabIndex        =   4
      Top             =   1200
      Width           =   4695
   End
   Begin VB.ListBox lstProdutos 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1005
      Left            =   1800
      MultiSelect     =   2  'Extended
      TabIndex        =   8
      Top             =   1920
      Width           =   5295
   End
   Begin VB.TextBox txtDescPro 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   3600
      TabIndex        =   7
      Top             =   1560
      Width           =   3495
   End
   Begin MSMask.MaskEdBox mebEmissao 
      Height          =   300
      Left            =   3240
      TabIndex        =   1
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtCodPro 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   1200
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtTurno 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   5160
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   1200
      TabIndex        =   10
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtHistorico 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   780
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   3000
      Width           =   3615
   End
   Begin VB.TextBox txtCodFunc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Timer tmrVale 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   720
      Top             =   1920
   End
   Begin ComCtl3.CoolBar cbrVale 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   1111
      BandCount       =   1
      _CBWidth        =   7230
      _CBHeight       =   630
      _Version        =   "6.7.9782"
      Child1          =   "tbrVale"
      MinHeight1      =   38
      Width1          =   209
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrVale 
         Height          =   570
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlVale"
         DisabledImageList=   "imlDVale"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "novo"
               Object.ToolTipText     =   "Novo"
               ImageKey        =   "novo"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "editar"
               Object.ToolTipText     =   "Editar"
               ImageKey        =   "editar"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "localizar"
               Object.ToolTipText     =   "Localizar"
               ImageKey        =   "localizar"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "salvar"
               Object.ToolTipText     =   "Salvar"
               ImageKey        =   "salvar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cancelar"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "cancelar"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "excluir"
               Object.ToolTipText     =   "Excluir"
               ImageKey        =   "excluir"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sair"
               Object.ToolTipText     =   "Sair"
               ImageKey        =   "sair"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlDVale 
      Left            =   120
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":0CCA
            Key             =   "novo"
            Object.Tag             =   "novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":19A4
            Key             =   "editar"
            Object.Tag             =   "editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":267E
            Key             =   "salvar"
            Object.Tag             =   "salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":3358
            Key             =   "excluir"
            Object.Tag             =   "excluir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":4032
            Key             =   "sair"
            Object.Tag             =   "sair"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":4D0C
            Key             =   "cancelar"
            Object.Tag             =   "cancelar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":59E6
            Key             =   "localizar"
            Object.Tag             =   "localizar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":66C0
            Key             =   "inserir"
            Object.Tag             =   "inserir"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":739A
            Key             =   "remover"
            Object.Tag             =   "remover"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlVale 
      Left            =   120
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":8074
            Key             =   "novo"
            Object.Tag             =   "novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":8D4E
            Key             =   "editar"
            Object.Tag             =   "editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":9A28
            Key             =   "salvar"
            Object.Tag             =   "salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":A702
            Key             =   "excluir"
            Object.Tag             =   "excluir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":B3DC
            Key             =   "sair"
            Object.Tag             =   "sair"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":C0B6
            Key             =   "cancelar"
            Object.Tag             =   "cancelar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":CD90
            Key             =   "localizar"
            Object.Tag             =   "localizar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":D1E2
            Key             =   "inserir"
            Object.Tag             =   "inserir"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVale.frx":DEBC
            Key             =   "remover"
            Object.Tag             =   "remover"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbVale 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   11
      Top             =   4245
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10345
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1852
            MinWidth        =   1852
            Text            =   "F1 - Ajuda"
            TextSave        =   "F1 - Ajuda"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgRemover 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   480
      Left            =   1200
      ToolTipText     =   "Remover  Ctrl+R"
      Top             =   2445
      Width           =   480
   End
   Begin VB.Image imgInserir 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   480
      Left            =   1200
      ToolTipText     =   "Inserir   Ctrl+I"
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label labProdutos 
      AutoSize        =   -1  'True
      Caption         =   "Produtos:"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   1620
      Width           =   675
   End
   Begin VB.Label labTurno 
      AutoSize        =   -1  'True
      Caption         =   "Turno:"
      Height          =   195
      Left            =   4560
      TabIndex        =   19
      Top             =   900
      Width           =   465
   End
   Begin VB.Label labValor 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   3900
      Width           =   510
   End
   Begin VB.Label labHistorico 
      AutoSize        =   -1  'True
      Caption         =   "Histórico:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   3060
      Width           =   825
   End
   Begin VB.Label labFuncionario 
      AutoSize        =   -1  'True
      Caption         =   "Funcionário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1260
      Width           =   1065
   End
   Begin VB.Label labEmissao 
      AutoSize        =   -1  'True
      Caption         =   "Emissão:"
      Height          =   195
      Left            =   2520
      TabIndex        =   15
      Top             =   900
      Width           =   630
   End
   Begin VB.Label labCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Nº Docto:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   893
      Width           =   705
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mniNovo 
         Caption         =   "&Novo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mniEditar 
         Caption         =   "&Editar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mniLocalizar 
         Caption         =   "&Localizar"
         Shortcut        =   ^L
      End
      Begin VB.Menu mniSepUm 
         Caption         =   "-"
      End
      Begin VB.Menu mniSalvar 
         Caption         =   "&Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mniCancelar 
         Caption         =   "&Cancelar"
         Shortcut        =   ^U
      End
      Begin VB.Menu mniExcluir 
         Caption         =   "E&xcluir"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mniSepDois 
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
Attribute VB_Name = "frmVale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Conn As clsMyConnect
Dim Aux As clsMyConnect
Public Localizar As Boolean 'Define o retorno da tela localizar
Public Auth As Boolean
Dim LocalizarOQue As String
Dim Sair As Boolean
Dim botaoPressionado As String

Private Sub Form_Load()
    imgInserir.Picture = imlDVale.ListImages.Item("inserir").Picture
    imgRemover.Picture = imlDVale.ListImages.Item("remover").Picture
    Sair = True
    Localizar = False
    LocalizarOQue = "Vale"
    botaoPressionado = Empty
    mniEditar.Enabled = False
    mniSalvar.Enabled = False
    mniExcluir.Enabled = False
    mniCancelar.Enabled = False
    tbrVale.Buttons.Item("editar").Enabled = False
    tbrVale.Buttons.Item("salvar").Enabled = False
    tbrVale.Buttons.Item("excluir").Enabled = False
    tbrVale.Buttons.Item("cancelar").Enabled = False
    Set Conn = New clsMyConnect
    Set Aux = New clsMyConnect
    Call Conn.Connect
    If Conn.NumErro <> 0 Then GoTo SubFail
    Call Aux.Connect
    If Conn.NumErro <> 0 Then GoTo SubFail
    Call QueryMain
    If Conn.NumErro <> 0 Then GoTo SubFail
    Exit Sub
SubFail:
    MsgBox "Não foi Possível Conectar-se com a Base de Dados." _
        & Chr(13) & "Erro # " & Str(Conn.NumErro) & " foi gerado por " _
        & Conn.SrcErro & Chr(13) & Conn.DescErro, vbCritical, "Falha de Conexão"
End Sub

Private Sub imgInserir_Click()
    Dim PriHifen As Integer
    Dim SecHifen As Integer
    If CSng(txtVlrPro.Text) = Empty Then
        MsgBox "Valor do Produto Deve Ser Maior que Zero", vbInformation, _
            "Valor de Produto Inválido"
        Exit Sub
    End If
    PriHifen = InStr(1, lstProdutos.List(lstProdutos.ListIndex), "-") + 5
    SecHifen = InStrRev(lstProdutos.List(lstProdutos.ListIndex), "-") - PriHifen - 1
    Call lstProdutos.AddItem(txtCodPro.Text & " - " & txtVlrPro.Text _
        & " - " & txtDescPro.Text)
    If txtValor.Text = Empty Then
        txtValor.Text = Format(CSng(txtVlrPro.Text), "R$ ###,##0.00")
    Else
        txtValor.Text = Format(CSng(txtValor.Text) _
            + CSng(txtVlrPro.Text), "R$ ###,##0.00")
    End If
    txtCodPro.Text = Empty
    txtDescPro.Text = Empty
    txtVlrPro.Text = Empty
    lstProdutos.Enabled = True
    imgInserir.Enabled = False
    txtCodPro.SetFocus
End Sub

Private Sub imgInserir_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgInserir.BorderStyle = 1
End Sub

Private Sub imgInserir_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgInserir.BorderStyle = 0
End Sub

Private Sub imgRemover_Click()
    Dim I As Integer
    For I = lstProdutos.ListCount - 1 To 0 Step -1
        If lstProdutos.Selected(I) = True Then
            Dim PriHifen As Integer
            Dim SecHifen As Integer
            PriHifen = InStr(1, lstProdutos.List(lstProdutos.ListIndex), "-") + 5
            SecHifen = InStrRev(lstProdutos.List(lstProdutos.ListIndex), "-") - PriHifen - 1
            txtValor.Text = Format(CSng(txtValor.Text) - Mid(lstProdutos.List( _
                lstProdutos.ListIndex), PriHifen, SecHifen), "R$ ###,##0.00")
            Call lstProdutos.RemoveItem(I)
        End If
    Next
    imgRemover.Enabled = False
    imgRemover.Picture = imlDVale.ListImages.Item("remover").Picture
    If lstProdutos.ListCount = 0 Then
        lstProdutos.Enabled = False
    End If
    If CSng(txtValor.Text) = Empty Then
        txtValor.Text = Empty
    End If
End Sub

Private Sub imgRemover_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgRemover.BorderStyle = 1
End Sub

Private Sub imgRemover_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgRemover.BorderStyle = 0
End Sub

Private Sub lstProdutos_Click()
    imgRemover.Enabled = True
    imgRemover.Picture = imlVale.ListImages.Item("remover").Picture
End Sub

Private Sub lstProdutos_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If lstProdutos.ListIndex = -1 Then
            Call Time
            tmrVale.Interval = 3000
            tmrVale.Enabled = True
            stbVale.Panels.Item(1).Text = "Por Favor Selecione um Item da Lista"
            Exit Sub
        End If
        imgRemover_Click
    End If
End Sub

Private Sub tbrVale_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "novo"
            mniNovo_Click
        Case "editar"
            mniEditar_Click
        Case "localizar"
            mniLocalizar_Click
        Case "salvar"
            mniSalvar_Click
        Case "cancelar"
            mniCancelar_Click
        Case "excluir"
            mniExcluir_Click
        Case "sair"
            mniSair_Click
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = IIf(Sair, 0, 1)
    If Cancel = 1 Then
        MsgBox "Existe um Cadastro em Aberto." _
            & Chr(13) & "Por Favor Encerre Antes de Fechar."
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Conn.Disconnect
End Sub

Private Sub mebEmissao_GotFocus()
    mebEmissao.SelStart = 0
    mebEmissao.SelLength = 10
    mebEmissao.BackColor = &H80000018
End Sub

Private Sub mebEmissao_LostFocus()
    mebEmissao.BackColor = &H80000005
End Sub

Private Sub mniCancelar_Click()
    Call HabCampos(True)
    Call ZeraValores
    botaoPressionado = Empty
    txtCodigo.SetFocus
End Sub

Private Sub mniEditar_Click()
    Call HabCampos(False)
    botaoPressionado = "alterar"
    txtCodFunc.SetFocus
End Sub

Private Sub mniExcluir_Click()
    If MsgBox("Deseja Excluir o Registro Selecionado?", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Excluir Registro.") = vbNo Then
        Exit Sub
    End If
    Call HabCampos(True)
    Call Conn.Query("DELETE FROM vale WHERE id = " & txtCodigo.Text)
    If Conn.NumErro = -2147217871 Then 'Falha de chave estrangeira
        Call Time
        MsgBox "Vale possui relacionamentos" _
            & Chr(13) & "e não pode ser excluído." _
            , vbExclamation, "Falha de Exclusão"
    ElseIf Not Conn.NumErro = 0 Then
        Call Time
        MsgBox "Não foi Possível Excluir o Registro."
    Else
        Call Time
        tmrVale.Enabled = True
        stbVale.Panels.Item(1).Text = "Registro Excluído com Êxito."
        Call ZeraValores
        Call QueryMain
        If Conn.NumErro <> 0 Then
            MsgBox "Não foi Possível Reler os Dados." _
                & Chr(13) & Me.Caption & " Será Fechado."
            Unload Me
        End If
    End If
End Sub

Private Sub mniLocalizar_Click()
    If Conn.Rs.RecordCount = 0 Then
        MsgBox "Não Há Dados Registrados", vbInformation, "Sem Registros"
        Exit Sub
    End If
    Set frmPrin.frmPai = Me
    frmLocalizar.Show vbModal, Me
    If Localizar = True Then
        txtCodigo.Text = Conn.Rs.Fields.Item(0).Value
    End If
End Sub

Private Sub mniNovo_Click()
    Call HabCampos(False)
    mebEmissao.Text = Date
    txtTurno.Text = 1
    botaoPressionado = "novo"
    txtCodFunc.SetFocus
End Sub

Private Sub mniSair_Click()
    Unload Me
End Sub

Private Sub mniSalvar_Click()
    If txtCodFunc.Text = Empty Or _
        txtHistorico.Text = Empty Or _
        txtValor.Text = Empty Then
            MsgBox "Por Favor Preencha os Campos em Negrito", vbInformation, _
                "Preencher Campos"
            Exit Sub
    End If
    If Not txtCodPro.Text = Empty Then
        MsgBox "Há um Produto não Lançado", vbInformation, "Produto não Lançado"
        Exit Sub
    End If
    If CSng(txtValor.Text) + CSng(labValVales.Caption) > _
        CSng(labValLimite.Caption) Then
        MsgBox "Valor do Vale Excede o Limite do Funcionário", vbCritical, _
            "Limite Excedido"
        Exit Sub
    End If
    Auth = False
    Call Conn.Query("SELECT id, senha FROM funcionario" _
        & " WHERE id = " & txtCodFunc.Text)
    frmAutorizacao.Show vbModal, Me
    Call QueryMain
    If Not Auth Then Exit Sub
    Dim idReg As Long
    Dim myQuery As String
    idReg = 0
    If botaoPressionado = "novo" Then
        Call Conn.Query("INSERT INTO vale(id,fid) VALUES(null," _
            & txtCodFunc.Text & ")")
        If Conn.NumErro = -2147217871 Then
            MsgBox "Não foi Possível Inserir o Registro." _
                & Chr(13) & "Erro # " & Str(Conn.NumErro) _
                & " foi gerado por " & Conn.SrcErro & Chr(13) _
                & "Funcionário Inexistente", vbCritical, "Falha de Gravação"
            Conn.NumErro = 0
            txtCodFunc.SetFocus
            Exit Sub
        ElseIf Conn.NumErro <> 0 Then
            MsgBox "Não foi Possível Inserir o Registro." _
                & Chr(13) & "Erro # " & Str(Conn.NumErro) _
                & " foi gerado por " & Conn.SrcErro & Chr(13) _
                & Conn.DescErro, vbCritical, "Falha de Gravação"
            Exit Sub
        Else
            idReg = Conn.GetValue("SELECT MAX(id) FROM vale")
        End If
    Else
        idReg = CLng(txtCodigo.Text)
    End If
    mebEmissao.PromptInclude = True
    myQuery = "UPDATE vale SET" _
        & " fid = " & txtCodFunc.Text _
        & ", emissao = '" & modData.VbToMy(mebEmissao.Text) & "'" _
        & ", historico = '" & txtHistorico.Text & "'" _
        & ", valor = " & Replace(CDbl(txtValor.Text), ",", ".") _
        & ", turno = " & txtTurno.Text _
        & " WHERE id = " & idReg
    mebEmissao.PromptInclude = False
    Call Conn.Query(myQuery)
    If Conn.NumErro <> 0 Then
        MsgBox "Não foi Possível Gravar Registro." _
            & Chr(13) & "Erro # " & Str(Conn.NumErro) _
            & " foi gerado por " & Conn.SrcErro & Chr(13) _
            & Conn.DescErro, vbCritical, "Falha de Gravação"
        Exit Sub
    Else
        Call Time
        tmrVale.Enabled = True
        stbVale.Panels.Item(1).Text = "Registro Gravado com Êxito."
        Call QueryMain
        If Conn.NumErro <> 0 Then
            MsgBox "Não foi Possível Reler os Dados." _
                & Chr(13) & "Cadastro de Usuários Será Fechado."
        End If
    End If
    Call HabCampos(True)
    If botaoPressionado = "novo" Then
        Call ZeraValores
    End If
    botaoPressionado = Empty
    txtCodigo.SetFocus
End Sub

Private Sub tmrVale_Timer()
    Call Time
End Sub

Private Sub txtCodFunc_Change()
    If txtCodFunc.Text = Empty Then GoTo Limpa
    Call Aux.Query("SELECT nome, limite FROM funcionario" _
        & " WHERE id = " & txtCodFunc.Text)
    If Aux.Rs.RecordCount = 0 Then
        txtNomeFunc.Text = "Registro Inexistente."
        labValLimite.Caption = "R$ 0,00"
        labValSaldo.Caption = "R$ 0,00"
        labValVales.Caption = "R$ 0,00"
    Else
        txtNomeFunc.Text = Aux.Rs.Fields.Item("nome").Value
        labValLimite.Caption = Format(Aux.Rs.Fields.Item("limite").Value, _
            "R$ ###,##0.00")
        Call Aux.Query("SELECT SUM(valor) FROM vale WHERE fid = " _
            & txtCodFunc.Text)
        labValVales.Caption = Format(Aux.Rs.Fields.Item("SUM(valor)").Value, _
            "R$ ###,##0.00")
        'Previne erro em caso de a tabela de vales estar vazia
        If labValVales.Caption = Empty Then
            labValVales.Caption = "R$ 0,00"
        End If
        labValSaldo.Caption = Format(CSng(labValLimite.Caption) - _
            CSng(labValVales.Caption), "R$ ###,##0.00")
        If labValSaldo.Caption = Empty Then
            labValSaldo.Caption = "R$ 0,00"
        End If
    End If
    Exit Sub
Limpa:
    txtNomeFunc.Text = Empty
    labValLimite.Caption = "R$ 0,00"
    labValSaldo.Caption = "R$ 0,00"
    labValVales.Caption = "R$ 0,00"
End Sub

Private Sub txtCodFunc_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Numeros(KeyAscii)
End Sub

Private Sub txtCodFunc_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not KeyCode = vbKeyF5 Then Exit Sub
    Call Conn.Query("SELECT id, nome, cpf, situ, recvale, limite FROM funcionario")
    If Conn.Rs.RecordCount = 0 Then
        MsgBox "Não Há Dados Registrados", vbInformation, "Sem Registros"
        Call QueryMain
        Exit Sub
    End If
    Set frmPrin.frmPai = Me
    frmLocalizar.Show vbModal, Me
    If Localizar = True Then
        txtCodFunc.Text = Conn.Rs.Fields.Item("id").Value
    End If
    Call QueryMain
End Sub

Private Sub txtCodigo_Change()
    Call Localiza
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Numeros(KeyAscii)
End Sub

Private Sub Localiza()
    Dim Hab As Boolean
    Dim Vls As Single
    If Not txtCodigo.Text = Empty Then
        Call QueryMain
        Call Time
        Conn.Rs.Find "id = " & txtCodigo.Text
        If Conn.Rs.EOF Then
            tmrVale.Enabled = True
            stbVale.Panels.Item(1).Text = "Registro Inexistente."
        Else
            mebEmissao.Text = IIf(IsNull(Conn.Rs.Fields.Item("emissao").Value), "", Conn.Rs.Fields.Item("emissao").Value)
            txtTurno.Text = Conn.Rs.Fields.Item("turno").Value
            txtCodFunc.Text = Conn.Rs.Fields.Item("fid").Value
            txtNomeFunc.Text = IIf(IsNull(Conn.Rs.Fields.Item("nome").Value), Empty, Conn.Rs.Fields.Item("nome").Value)
            txtHistorico.Text = Conn.Rs.Fields.Item("historico").Value
            txtValor.Text = Conn.Rs.Fields.Item("valor").Value
            Call txtValor_LostFocus
            Call Aux.Query("SELECT valeItem.id, valeItem.pid, valeItem.vid" _
                & ", valeItem.vPro, produto.descricao FROM valeItem" _
                & " LEFT JOIN produto ON valeItem.pid = produto.id" _
                & " WHERE valeItem.vid = " & Conn.Rs.Fields.Item("id").Value)
            lstProdutos.Clear
            Vls = 0
            While Not Aux.Rs.EOF
                Vls = Vls + Aux.Rs.Fields.Item("vPro").Value
                lstProdutos.AddItem ( _
                    Aux.Rs.Fields.Item("pid").Value & " - " & _
                    Aux.Rs.Fields.Item("descricao").Value)
                Aux.Rs.MoveNext
            Wend
            labValLimite.Caption = Format(Conn.Rs.Fields.Item("limite").Value, "R$ ###,##0.00")
            labValVales.Caption = Format(Vls, "R$ ###,##0.00")
            labValSaldo.Caption = Format(CSng(Mid(labValLimite.Caption, 3)) _
                - CSng(Mid(labValVales.Caption, 3)), "R$ ###,##0.00")
            Hab = True
            GoTo habilita
        End If
    End If
    mebEmissao.Text = Empty
    txtTurno.Text = Empty
    txtCodFunc.Text = Empty
    txtNomeFunc.Text = Empty
    txtCodPro.Text = Empty
    txtDescPro.Text = Empty
    txtHistorico.Text = Empty
    txtValor.Text = Empty
    txtVlrPro.Text = Empty
    lstProdutos.Clear
    labValLimite.Caption = "R$ 0,00"
    labValVales.Caption = "R$ 0,00"
    labValSaldo.Caption = "R$ 0,00"
    Hab = False
habilita:
    With tbrVale.Buttons
        .Item("localizar").Enabled = Not Hab
        .Item("editar").Enabled = Hab
        .Item("novo").Enabled = Not Hab
        .Item("salvar").Enabled = Hab
        .Item("cancelar").Enabled = Hab
        .Item("excluir").Enabled = Hab
    End With
    mniLocalizar.Enabled = Not Hab
    mniEditar.Enabled = Hab
    mniSalvar.Enabled = Hab
    mniCancelar.Enabled = Hab
    mniNovo.Enabled = Not Hab
    mniExcluir.Enabled = Hab
End Sub

Private Sub txtCodigo_GotFocus()
    txtCodigo.SelStart = 0
    txtCodigo.SelLength = 10
    txtCodigo.BackColor = &H80000018
End Sub

Private Sub txtCodigo_LostFocus()
    txtCodigo.BackColor = &H80000005
End Sub

Private Sub txtCodPro_Change()
    If txtCodPro.Text = Empty Then GoTo Limpa
    Call Aux.Query("SELECT descricao, preco FROM produto" _
        & " WHERE id = " & txtCodPro.Text)
    If Aux.Rs.RecordCount = 0 Then
        txtVlrPro.Text = Empty
        txtDescPro.Text = "Registro Inexistente."
        txtVlrPro.Enabled = False
        imgInserir.Enabled = False
        imgInserir.Picture = imlDVale.ListImages.Item("inserir").Picture
    Else
        txtDescPro.Text = Aux.Rs.Fields.Item("descricao").Value
        txtVlrPro.Text = Aux.Rs.Fields.Item("preco").Value
        Call txtVlrPro_LostFocus
        txtVlrPro.Enabled = True
        imgInserir.Enabled = True
        imgInserir.Picture = imlVale.ListImages.Item("inserir").Picture
    End If
    Exit Sub
Limpa:
    txtDescPro.Text = Empty
    txtVlrPro.Text = Empty
    txtVlrPro.Enabled = False
    imgInserir.Enabled = False
    imgInserir.Picture = imlDVale.ListImages.Item("inserir").Picture
End Sub

Private Sub txtCodPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If imgInserir.Enabled = False Then Exit Sub
        imgInserir_Click
        Exit Sub
    End If
    KeyAscii = modGetKeyAscii.Numeros(KeyAscii)
End Sub

Private Sub txtCodPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not KeyCode = 116 Then Exit Sub
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
    Call QueryMain
End Sub

Private Sub txtHistorico_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Maiusculas(KeyAscii)
End Sub

Private Sub txtTurno_GotFocus()
    txtTurno.SelStart = 0
    txtTurno.SelLength = 1
    txtTurno.BackColor = &H80000018
End Sub

Private Sub txtTurno_LostFocus()
    txtTurno.BackColor = &H80000005
End Sub

Private Sub txtCodFunc_GotFocus()
    txtCodFunc.SelStart = 0
    txtCodFunc.SelLength = 10
    txtCodFunc.BackColor = &H80000018
    stbVale.Panels.Item(1).Text = "Pressione F5 para Localizar"
    LocalizarOQue = "Funcionario"
End Sub

Private Sub txtCodFunc_LostFocus()
    txtCodFunc.BackColor = &H80000005
    stbVale.Panels.Item(1).Text = ""
    LocalizarOQue = "Vale"
End Sub

Private Sub txtNomeFunc_GotFocus()
    txtNomeFunc.SelStart = 0
    txtNomeFunc.SelLength = 50
    txtNomeFunc.BackColor = &H80000018
End Sub

Private Sub txtNomeFunc_LostFocus()
    txtNomeFunc.BackColor = &H80000005
End Sub

Private Sub txtCodPro_GotFocus()
    txtCodPro.SelStart = 0
    txtCodPro.SelLength = 10
    txtCodPro.BackColor = &H80000018
    stbVale.Panels.Item(1).Text = "Pressione F5 para Localizar"
    LocalizarOQue = "Produto"
End Sub

Private Sub txtCodPro_LostFocus()
    txtCodPro.BackColor = &H80000005
    stbVale.Panels.Item(1).Text = ""
    LocalizarOQue = "Vale"
End Sub

Private Sub txtDescPro_GotFocus()
    txtDescPro.SelStart = 0
    txtDescPro.SelLength = 50
    txtDescPro.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtDescPro_LostFocus()
    txtDescPro.BackColor = &H80000005 'Branco
End Sub

Private Sub lstProdutos_GotFocus()
    lstProdutos.BackColor = &H80000018 'Amarelo
End Sub

Private Sub lstProdutos_LostFocus()
    lstProdutos.BackColor = &H80000005 'Branco
End Sub

Private Sub txtHistorico_GotFocus()
    txtHistorico.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtHistorico_LostFocus()
    txtHistorico.BackColor = &H80000005 'Branco
End Sub

Private Sub txtValor_GotFocus()
    txtValor.SelStart = 0
    txtValor.SelLength = 13
    txtValor.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Moeda(KeyAscii)
End Sub

Private Sub txtValor_LostFocus()
    txtValor.BackColor = &H80000005 'Branco
    txtValor.Text = modMoeda.FmtMoeda(txtValor.Text)
End Sub

Private Sub HabCampos(situ As Boolean)
    txtCodigo.Enabled = situ
    txtCodFunc.Enabled = Not situ
    txtCodPro.Enabled = Not situ
    txtHistorico.Enabled = Not situ
    txtValor.Enabled = Not situ
    With tbrVale.Buttons
        .Item("localizar").Enabled = situ
        .Item("sair").Enabled = situ
        .Item("novo").Enabled = situ
        .Item("salvar").Enabled = Not situ
        .Item("cancelar").Enabled = Not situ
    End With
    mniLocalizar.Enabled = situ
    mniSair.Enabled = situ
    mniSalvar.Enabled = Not situ
    mniCancelar.Enabled = Not situ
    mniNovo.Enabled = situ
    Sair = situ
End Sub

Private Sub ZeraValores()
    tbrVale.Buttons.Item("editar").Enabled = False
    tbrVale.Buttons.Item("excluir").Enabled = False
    mniEditar.Enabled = False
    mniExcluir.Enabled = False
    mebEmissao.Text = Empty
    txtTurno.Text = Empty
    txtCodFunc.Text = Empty
    txtNomeFunc.Text = Empty
    txtCodPro.Text = Empty
    txtDescPro.Text = Empty
    txtHistorico.Text = Empty
    txtValor.Text = Empty
    txtVlrPro.Text = Empty
    lstProdutos.Clear
    labValLimite.Caption = "R$ 0,00"
    labValVales.Caption = "R$ 0,00"
    labValSaldo.Caption = "R$ 0,00"
End Sub

Private Sub Time()
    stbVale.Panels.Item(1).Text = Empty
    tmrVale.Interval = 0
    tmrVale.Enabled = False
    tmrVale.Interval = 10000
End Sub

Private Sub txtVlrPro_GotFocus()
    txtVlrPro.SelStart = 0
    txtVlrPro.SelLength = 13
    txtVlrPro.BackColor = &H80000018
End Sub

Private Sub txtVlrPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtVlrPro_LostFocus
        imgInserir_Click
        Exit Sub
    End If
    KeyAscii = modGetKeyAscii.Moeda(KeyAscii)
End Sub

Private Sub txtVlrPro_LostFocus()
    txtVlrPro.BackColor = &H80000005 'Branco
    txtVlrPro.Text = modMoeda.FmtMoeda(txtVlrPro.Text)
End Sub

Public Sub Load()
    Select Case LocalizarOQue
        Case "Vale"
            Call LoadVale
        Case "Funcionario"
            Call LoadFuncionario
        Case "Produto"
            Call LoadProduto
    End Select
End Sub

' cmpProc = Campo Procurado
Public Sub Campo(txt As String)
    Select Case LocalizarOQue
        Case "Vale"
            Call CampoVale(txt)
        Case "Funcionario"
            Call CampoFuncionario(txt)
        Case "Produto"
            Call CampoProduto(txt)
    End Select
End Sub

Public Sub LoadVale()
    With frmLocalizar.cbbProcura
        .AddItem ("Nº Docto")
        .AddItem ("Emissão")
        .AddItem ("Funcionário")
        .Text = "Funcionário"
    End With
    With frmLocalizar.mfgLocalizar
        .Cols = 4
        .Rows = 1
        .Clear
        .ColWidth(0) = 800
        .ColWidth(1) = 3800
        .ColWidth(2) = 1500
        .ColWidth(3) = 1447
        .TextArray(0) = "Nº Docto"
        .TextArray(1) = "Funcionário"
        .TextArray(2) = "Emissão"
        .TextArray(3) = "Valor"
    End With
    ReDim colsWidth(0 To 4)
    Dim Rows As Integer: Rows = 1
    Dim fieldWidth As Integer
    Conn.Rs.MoveFirst
    Do While Not Conn.Rs.EOF
        frmLocalizar.mfgLocalizar.Rows = frmLocalizar.mfgLocalizar.Rows + 1
        frmLocalizar.mfgLocalizar.TextMatrix(Rows, 0) = Conn.Rs.Fields.Item("id").Value
        fieldWidth = TextWidth(Conn.Rs.Fields.Item("id").Value)
        If colsWidth(0) < fieldWidth Then
            colsWidth(0) = fieldWidth
        End If
        frmLocalizar.mfgLocalizar.TextMatrix(Rows, 1) = Conn.Rs.Fields.Item("nome").Value
        fieldWidth = TextWidth(Conn.Rs.Fields.Item("nome").Value)
        If colsWidth(1) < fieldWidth Then
            colsWidth(1) = fieldWidth
        End If
        frmLocalizar.mfgLocalizar.TextMatrix(Rows, 2) = Conn.Rs.Fields.Item("emissao").Value
        fieldWidth = TextWidth(Conn.Rs.Fields.Item("emissao").Value)
        If colsWidth(2) < fieldWidth Then
            colsWidth(2) = fieldWidth
        End If
        frmLocalizar.mfgLocalizar.TextMatrix(Rows, 3) = modMoeda.FmtMoeda(Conn.Rs.Fields.Item("valor").Value)
        fieldWidth = TextWidth(modMoeda.FmtMoeda(Conn.Rs.Fields.Item("valor").Value))
        If colsWidth(3) < fieldWidth Then
            colsWidth(3) = fieldWidth
        End If
        Conn.Rs.MoveNext
        Rows = Rows + 1
    Loop
    frmLocalizar.Caption = "Localização de Vales"
    frmLocalizar.tabela = "vale"
End Sub

Public Sub CampoVale(txt As String)
    Select Case txt
        Case "Nº Docto"
            frmLocalizar.Campo = "id"
        Case "Funcionário"
            frmLocalizar.Campo = "nome"
        Case "Emissão"
            frmLocalizar.Campo = "emissao"
        Case "Valor"
            frmLocalizar.Campo = "valor"
    End Select
End Sub

Public Sub LoadFuncionario()
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
    End With
    ReDim colsWidth(0 To 3)
    Dim Rows As Integer: Rows = 1
    Dim fieldWidth As Integer
    Conn.Rs.MoveFirst
    Do While Not Conn.Rs.EOF
        frmLocalizar.mfgLocalizar.Rows = frmLocalizar.mfgLocalizar.Rows + 1
        frmLocalizar.mfgLocalizar.TextMatrix(Rows, 0) = Conn.Rs.Fields.Item("id").Value
        fieldWidth = TextWidth(Conn.Rs.Fields.Item("id").Value)
        If colsWidth(0) < fieldWidth Then
            colsWidth(0) = fieldWidth
        End If
        frmLocalizar.mfgLocalizar.TextMatrix(Rows, 1) = Conn.Rs.Fields.Item("nome").Value
        fieldWidth = TextWidth(Conn.Rs.Fields.Item("nome").Value)
        If colsWidth(1) < fieldWidth Then
            colsWidth(1) = fieldWidth
        End If
        frmLocalizar.mfgLocalizar.TextMatrix(Rows, 2) = Conn.Rs.Fields.Item("cpf").Value
        fieldWidth = TextWidth(Conn.Rs.Fields.Item("cpf").Value)
        If colsWidth(2) < fieldWidth Then
            colsWidth(2) = fieldWidth
        End If
        Conn.Rs.MoveNext
        Rows = Rows + 1
    Loop
    frmLocalizar.Caption = "Localização de Funcionários"
    frmLocalizar.tabela = "funcionario"
End Sub

Public Sub CampoFuncionario(txt As String)
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
        ReDim colsWidth(0 To 4)
        Dim Rows As Integer: Rows = 1
        Dim fieldWidth As Integer
        Conn.Rs.MoveFirst
        Do While Not Conn.Rs.EOF
            .Rows = frmLocalizar.mfgLocalizar.Rows + 1
            .TextMatrix(Rows, 0) = Conn.Rs.Fields.Item("id").Value
            fieldWidth = TextWidth(Conn.Rs.Fields.Item("id").Value)
            If colsWidth(0) < fieldWidth Then
                colsWidth(0) = fieldWidth
            End If
            .TextMatrix(Rows, 1) = Conn.Rs.Fields.Item("descricao").Value
            fieldWidth = TextWidth(Conn.Rs.Fields.Item("descricao").Value)
            If colsWidth(1) < fieldWidth Then
                colsWidth(1) = fieldWidth
            End If
            .TextMatrix(Rows, 2) = modMoeda.FmtMoeda(Conn.Rs.Fields.Item("preco").Value)
            fieldWidth = TextWidth(modMoeda.FmtMoeda(Conn.Rs.Fields.Item("preco").Value))
            If colsWidth(2) < fieldWidth Then
                colsWidth(2) = fieldWidth
            End If
            Dim x As Integer
            Dim y As String
            x = Conn.Rs.Fields.Item("tipo").Value
            y = IIf(x = 0, "Comida", "Bebida")
            .TextMatrix(Rows, 3) = y
            fieldWidth = TextWidth(y)
            If colsWidth(3) < fieldWidth Then
                colsWidth(3) = fieldWidth
            End If
            Conn.Rs.MoveNext
            Rows = Rows + 1
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

Private Sub QueryMain()
    Call Conn.Query("SELECT vale.id, vale.fid, vale.emissao" _
        & ", vale.historico, vale.valor, vale.turno" _
        & ", funcionario.nome, funcionario.limite FROM vale" _
        & " LEFT JOIN funcionario ON vale.fid = funcionario.id")
End Sub
