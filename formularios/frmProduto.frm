VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmProduto 
   Caption         =   "Cadastro de Produtos"
   ClientHeight    =   3075
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6945
   Icon            =   "frmProduto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   205
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   463
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbbDivPessoa 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox cbbTipo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtDescricao 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   5295
   End
   Begin VB.TextBox txtPreco 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Timer tmrProduto 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3480
      Top             =   1680
   End
   Begin MSComctlLib.ImageList imlProD 
      Left            =   4800
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":0CCA
            Key             =   "novo"
            Object.Tag             =   "novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":19A4
            Key             =   "editar"
            Object.Tag             =   "editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":267E
            Key             =   "salvar"
            Object.Tag             =   "salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":3358
            Key             =   "excluir"
            Object.Tag             =   "excluir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":4032
            Key             =   "sair"
            Object.Tag             =   "sair"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":4D0C
            Key             =   "cancelar"
            Object.Tag             =   "cancelar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":59E6
            Key             =   "localizar"
            Object.Tag             =   "localizar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlProduto 
      Left            =   4200
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":66C0
            Key             =   "novo"
            Object.Tag             =   "novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":739A
            Key             =   "editar"
            Object.Tag             =   "editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":8074
            Key             =   "salvar"
            Object.Tag             =   "salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":8D4E
            Key             =   "excluir"
            Object.Tag             =   "excluir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":9A28
            Key             =   "sair"
            Object.Tag             =   "sair"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":A702
            Key             =   "cancelar"
            Object.Tag             =   "cancelar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":B3DC
            Key             =   "localizar"
            Object.Tag             =   "localizar"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrProduto 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   1111
      BandCount       =   1
      _CBWidth        =   6945
      _CBHeight       =   630
      _Version        =   "6.7.9782"
      Child1          =   "tbrProduto"
      MinHeight1      =   38
      Width1          =   14
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrProduto 
         Height          =   570
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlProduto"
         DisabledImageList=   "imlProD"
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
   Begin MSComctlLib.StatusBar stbProduto 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   2730
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9869
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1852
            MinWidth        =   1852
            Text            =   "F1 - Ajuda"
            TextSave        =   "F1 - Ajuda"
         EndProperty
      EndProperty
   End
   Begin VB.Label labCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   893
      Width           =   540
   End
   Begin VB.Label labDescricao 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
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
      TabIndex        =   9
      Top             =   1260
      Width           =   930
   End
   Begin VB.Label labPreco 
      AutoSize        =   -1  'True
      Caption         =   "Preço:"
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
      TabIndex        =   8
      Top             =   1620
      Width           =   570
   End
   Begin VB.Label labTipo 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
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
      TabIndex        =   7
      Top             =   1980
      Width           =   450
   End
   Begin VB.Label labDivPessoa 
      AutoSize        =   -1  'True
      Caption         =   "Divisão Pessoa:"
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
      TabIndex        =   6
      Top             =   2340
      Width           =   1380
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
Attribute VB_Name = "frmProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Conn As clsMyConnect
Public Localizar As Boolean
Dim Sair As Boolean
Dim botaoPressionado As String

Private Sub cbbDivPessoa_GotFocus()
    cbbDivPessoa.BackColor = &H80000018 'Amarelo
End Sub

Private Sub cbbDivPessoa_LostFocus()
    cbbDivPessoa.BackColor = &H80000005 'Branco
End Sub

Private Sub cbbTipo_GotFocus()
    cbbTipo.BackColor = &H80000018 'Amarelo
End Sub

Private Sub cbbTipo_LostFocus()
    cbbTipo.BackColor = &H80000005 'Branco
End Sub

Private Sub Form_Load()
    cbbTipo.AddItem "Bebida"
    cbbTipo.AddItem "Comida"
    cbbTipo.Text = "Comida"
    cbbDivPessoa.AddItem "Não"
    cbbDivPessoa.AddItem "Sim"
    cbbDivPessoa.Text = "Sim"
    Sair = True
    Localizar = False
    botaoPressionado = Empty
    mniEditar.Enabled = False
    mniSalvar.Enabled = False
    mniExcluir.Enabled = False
    mniCancelar.Enabled = False
    tbrProduto.Buttons.Item("editar").Enabled = False
    tbrProduto.Buttons.Item("salvar").Enabled = False
    tbrProduto.Buttons.Item("excluir").Enabled = False
    tbrProduto.Buttons.Item("cancelar").Enabled = False
    Set Conn = New clsMyConnect
    Call Conn.Connect
    If Conn.NumErro <> 0 Then GoTo SubFail
    Call Conn.Query("SELECT id, descricao, preco, tipo, divpessoa FROM produto")
    If Conn.NumErro <> 0 Then GoTo SubFail
    Exit Sub
SubFail:
    MsgBox "Não foi Possível Conectar-se com a Base de Dados." _
        & Chr(13) & "Erro # " & Str(Conn.NumErro) & " foi gerado por " _
        & Conn.SrcErro & Chr(13) & Conn.DescErro, vbCritical, "Falha de Conexão"
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

Private Sub mniCancelar_Click()
    Call HabCampos(True)
    Call ZeraValores
    botaoPressionado = Empty
    txtCodigo.SetFocus
End Sub

Private Sub mniEditar_Click()
    Call HabCampos(False)
    botaoPressionado = "alterar"
    txtDescricao.SetFocus
End Sub

Private Sub mniExcluir_Click()
    If MsgBox("Deseja Excluir o Registro Selecionado?", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Excluir Registro.") = vbNo Then
        Exit Sub
    End If
    Call HabCampos(True)
    Call Conn.Query("DELETE FROM produto WHERE id = " & txtCodigo.Text)
    If Conn.NumErro = -2147217871 Then 'Falha de chave estrangeira
        Call Time
        MsgBox "Produto possui relacionamentos" _
            & Chr(13) & "e não pode ser excluído." _
            , vbExclamation, "Falha de Exclusão"
    ElseIf Not Conn.NumErro = 0 Then
        Call Time
        MsgBox "Não foi Possível Excluir o Registro."
    Else
        Call Time
        tmrProduto.Enabled = True
        stbProduto.Panels.Item(1).Text = "Registro Excluído com Êxito."
        Call ZeraValores
        Call Conn.Query("SELECT id, descricao, preco, tipo, divpessoa FROM produto")
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
    botaoPressionado = "novo"
    txtDescricao.SetFocus
End Sub

Private Sub mniSair_Click()
    Unload Me
End Sub

Private Sub mniSalvar_Click()
    If txtDescricao.Text = Empty Or _
        txtPreco.Text = Empty Then
            MsgBox "Por Favor Preencha os Campos em Negrito", _
                vbInformation, "Informação!"
            Exit Sub
    End If
    Dim idReg As Long
    Dim myQuery As String
    idReg = 0
    If botaoPressionado = "novo" Then
        Call Conn.Query("INSERT INTO produto(id) VALUES(null)")
        If Conn.NumErro <> 0 Then
            MsgBox "Não foi Possível Inserir o Registro.", vbCritical, "Falha de Gravação"
            Exit Sub
        Else
            idReg = Conn.GetValue("SELECT MAX(id) FROM produto")
        End If
    Else
        'Converte Caracter para Tipo Long
        idReg = CLng(txtCodigo.Text)
    End If
    myQuery = "UPDATE produto SET" _
        & " descricao = '" & txtDescricao.Text & "'" _
        & ", tipo = " & cbbTipo.ListIndex _
        & ", divpessoa = " & cbbDivPessoa.ListIndex _
        & ", preco = " & Replace(CDbl(txtPreco.Text), ",", ".") _
        & " WHERE id = " & idReg
    Call Conn.Query(myQuery)
    If Conn.NumErro <> 0 Then
        MsgBox "Não foi Possível Gravar Registro.", _
            vbCritical, "Falha de Gravação"
        Exit Sub
    Else
        Call Time
        tmrProduto.Enabled = True
        stbProduto.Panels.Item(1).Text = "Registro Gravado com Êxito."
        Call Conn.Query("SELECT id, descricao, preco, tipo, divpessoa FROM produto")
        If Conn.NumErro <> 0 Then
            MsgBox "Não foi Possível Reler os Dados." _
                & Chr(13) & Me.Caption & " Será Fechado."
        End If
    End If
    Call HabCampos(True)
    If botaoPressionado = "novo" Then
        Call ZeraValores
    End If
    botaoPressionado = Empty
    txtCodigo.SetFocus
End Sub

Private Sub tbrProduto_ButtonClick(ByVal Button As MSComctlLib.Button)
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

Private Sub tmrProduto_Timer()
    Call Time
End Sub

Private Sub txtCodigo_Change()
    Call Localiza
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Numeros(KeyAscii)
End Sub

Private Sub Localiza()
    Dim Hab As Boolean
    If Not txtCodigo.Text = Empty Then
        Call Conn.Query("SELECT id, descricao, preco, tipo, divpessoa FROM produto")
        Call Time
        Conn.Rs.Find "id = " & txtCodigo.Text
        If Conn.Rs.BOF Or Conn.Rs.EOF Then
            tmrProduto.Enabled = True
            stbProduto.Panels.Item(1).Text = "Registro Inexistente."
        Else
            txtDescricao.Text = Conn.Rs.Fields.Item("descricao").Value
            cbbTipo.ListIndex = Conn.Rs.Fields.Item("tipo").Value
            cbbDivPessoa.ListIndex = Conn.Rs.Fields.Item("divpessoa").Value
            txtPreco.Text = Conn.Rs.Fields.Item("preco").Value
            Call txtPreco_LostFocus
            Hab = True
            GoTo habilita
        End If
    End If
    txtDescricao.Text = Empty
    cbbTipo.ListIndex = 1
    cbbDivPessoa.ListIndex = 1
    txtPreco.Text = Empty
    Hab = False
habilita:
    With tbrProduto.Buttons
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
    txtCodigo.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtCodigo_LostFocus()
    txtCodigo.BackColor = &H80000005 'Branco
End Sub

Private Sub txtPreco_GotFocus()
    txtPreco.SelStart = 0
    txtPreco.SelLength = 13
    txtPreco.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtPreco_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Moeda(KeyAscii)
End Sub

Private Sub txtPreco_LostFocus()
    txtPreco.BackColor = &H80000005 'Branco
    txtPreco.Text = modMoeda.FmtMoeda(txtPreco.Text)
End Sub

Private Sub txtDescricao_GotFocus()
    txtDescricao.SelStart = 0
    txtDescricao.SelLength = 50
    txtDescricao.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Maiusculas(KeyAscii)
End Sub

Private Sub txtDescricao_LostFocus()
    txtDescricao.BackColor = &H80000005 'Branco
End Sub

Private Sub HabCampos(situ As Boolean)
    txtCodigo.Enabled = situ
    txtDescricao.Enabled = Not situ
    cbbTipo.Enabled = Not situ
    cbbDivPessoa.Enabled = Not situ
    txtPreco.Enabled = Not situ
    With tbrProduto.Buttons
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
    tbrProduto.Buttons.Item("editar").Enabled = False
    tbrProduto.Buttons.Item("excluir").Enabled = False
    mniEditar.Enabled = False
    mniExcluir.Enabled = False
    txtCodigo.Text = Empty
    txtDescricao.Text = Empty
    txtPreco.Text = Empty
    cbbTipo.ListIndex = 1
    cbbDivPessoa.ListIndex = 1
End Sub

Private Sub Time()
    stbProduto.Panels.Item(1).Text = Empty
    tmrProduto.Interval = 0
    tmrProduto.Enabled = False
    tmrProduto.Interval = 10000
End Sub

Public Sub Load()
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
        .ColWidth(1) = 3800
        .ColWidth(2) = 1500
        .ColWidth(3) = 1447
        .TextArray(0) = "Código"
        .TextArray(1) = "Descrição"
        .TextArray(2) = "Preço"
        .TextArray(3) = "Tipo"
    End With
    ReDim colsWidth(0 To 4)
    Dim Rows As Integer: Rows = 1
    Dim fieldWidth As Integer
    Conn.Rs.MoveFirst
    Do While Not Conn.Rs.EOF
        frmLocalizar.mfgLocalizar.Rows = frmLocalizar.mfgLocalizar.Rows + 1
        frmLocalizar.mfgLocalizar.TextMatrix(Rows, 0) = _
            Conn.Rs.Fields.Item("id").Value
        fieldWidth = TextWidth(Conn.Rs.Fields.Item("id").Value)
        If colsWidth(0) < fieldWidth Then
            colsWidth(0) = fieldWidth
        End If
        frmLocalizar.mfgLocalizar.TextMatrix(Rows, 1) = _
            Conn.Rs.Fields.Item("descricao").Value
        fieldWidth = TextWidth(Conn.Rs.Fields.Item("descricao").Value)
        If colsWidth(1) < fieldWidth Then
            colsWidth(1) = fieldWidth
        End If
        frmLocalizar.mfgLocalizar.TextMatrix(Rows, 2) = _
            modMoeda.FmtMoeda(Conn.Rs.Fields.Item("preco").Value)
        fieldWidth = TextWidth(modMoeda.FmtMoeda(Conn.Rs.Fields.Item("preco").Value))
        If colsWidth(2) < fieldWidth Then
            colsWidth(2) = fieldWidth
        End If
        Dim x As Integer
        Dim y As String
        x = Conn.Rs.Fields.Item("tipo").Value
        y = IIf(x = 0, "Comida", "Bebida")
        frmLocalizar.mfgLocalizar.TextMatrix(Rows, 3) = y
        fieldWidth = TextWidth(y)
        If colsWidth(3) < fieldWidth Then
            colsWidth(3) = fieldWidth
        End If
        Conn.Rs.MoveNext
        Rows = Rows + 1
    Loop
    frmLocalizar.Caption = "Localização de Produtos"
    frmLocalizar.tabela = "produto"
End Sub

' cmpProc = Campo Procurado
Public Sub Campo(txt As String)
    Select Case txt
        Case "Código"
            frmLocalizar.Campo = "id"
        Case "Descrição"
            frmLocalizar.Campo = "descricao"
    End Select
End Sub
