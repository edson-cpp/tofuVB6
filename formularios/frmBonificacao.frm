VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmBonificacao 
   Caption         =   "Cadastro de Bonificação"
   ClientHeight    =   3150
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6765
   Icon            =   "frmBonificacao.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   210
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   451
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   1200
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtHistorico 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   780
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox txtCodFunc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtNomeFunc 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Timer tmrBonificacao 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   4920
      Top             =   1800
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
   Begin MSComctlLib.ImageList imlBonifD 
      Left            =   6000
      Top             =   1680
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
            Picture         =   "frmBonificacao.frx":0CCA
            Key             =   "novo"
            Object.Tag             =   "novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBonificacao.frx":19A4
            Key             =   "editar"
            Object.Tag             =   "editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBonificacao.frx":267E
            Key             =   "salvar"
            Object.Tag             =   "salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBonificacao.frx":3358
            Key             =   "excluir"
            Object.Tag             =   "excluir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBonificacao.frx":4032
            Key             =   "sair"
            Object.Tag             =   "sair"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBonificacao.frx":4D0C
            Key             =   "cancelar"
            Object.Tag             =   "cancelar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBonificacao.frx":59E6
            Key             =   "localizar"
            Object.Tag             =   "localizar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlBonificacao 
      Left            =   5400
      Top             =   1680
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
            Picture         =   "frmBonificacao.frx":66C0
            Key             =   "novo"
            Object.Tag             =   "novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBonificacao.frx":739A
            Key             =   "editar"
            Object.Tag             =   "editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBonificacao.frx":8074
            Key             =   "salvar"
            Object.Tag             =   "salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBonificacao.frx":8D4E
            Key             =   "excluir"
            Object.Tag             =   "excluir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBonificacao.frx":9A28
            Key             =   "sair"
            Object.Tag             =   "sair"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBonificacao.frx":A702
            Key             =   "cancelar"
            Object.Tag             =   "cancelar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBonificacao.frx":B3DC
            Key             =   "localizar"
            Object.Tag             =   "localizar"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrBonificacao 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   1111
      BandCount       =   1
      _CBWidth        =   6765
      _CBHeight       =   630
      _Version        =   "6.7.9782"
      Child1          =   "tbrBonificacao"
      MinHeight1      =   38
      Width1          =   14
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrBonificacao 
         Height          =   570
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlBonificacao"
         DisabledImageList=   "imlBonifD"
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
   Begin MSComctlLib.StatusBar stbBonificacao 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   8
      Top             =   2805
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9525
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1852
            MinWidth        =   1852
            Text            =   "F1 - Ajuda"
            TextSave        =   "F1 - Ajuda"
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox mebEmissao 
      Height          =   300
      Left            =   3120
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
      TabIndex        =   13
      Top             =   2460
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
      TabIndex        =   12
      Top             =   1620
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
      TabIndex        =   11
      Top             =   1260
      Width           =   1065
   End
   Begin VB.Label labEmissao 
      AutoSize        =   -1  'True
      Caption         =   "Emissão:"
      Height          =   195
      Left            =   2400
      TabIndex        =   10
      Top             =   900
      Width           =   630
   End
   Begin VB.Label labCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Nº Docto:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   900
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
Attribute VB_Name = "frmBonificacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Conn As clsMyConnect
Private Aux As clsMyConnect
Public Localizar As Boolean
Dim LocalizarOQue As String
Dim Sair As Boolean
Dim botaoPressionado As String

Private Sub Form_Load()
    Sair = True
    Localizar = False
    LocalizarOQue = "Bonificação"
    botaoPressionado = Empty
    mniEditar.Enabled = False
    mniSalvar.Enabled = False
    mniExcluir.Enabled = False
    mniCancelar.Enabled = False
    tbrBonificacao.Buttons.Item("editar").Enabled = False
    tbrBonificacao.Buttons.Item("salvar").Enabled = False
    tbrBonificacao.Buttons.Item("excluir").Enabled = False
    tbrBonificacao.Buttons.Item("cancelar").Enabled = False
    Set Conn = New clsMyConnect
    Call Conn.Connect
    If Conn.NumErro <> 0 Then GoTo SubFail
    Set Aux = New clsMyConnect
    Call Aux.Connect
    If Aux.NumErro <> 0 Then GoTo SubFail
    Call QueryMain
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
End Sub

Private Sub mniExcluir_Click()
    If MsgBox("Deseja Excluir o Registro Selecionado?", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Excluir Registro.") = vbNo Then
        Exit Sub
    End If
    Call HabCampos(True)
    Call Aux.Query("DELETE FROM bonificacao WHERE id = " & txtCodigo.Text)
    If Aux.NumErro = -2147217871 Then 'Falha de chave estrangeira
        Call Time
        MsgBox "Bonificação possui relacionamentos" _
            & Chr(13) & "e não pode ser excluído." _
            , vbExclamation, "Falha de Exclusão"
    ElseIf Not Aux.NumErro = 0 Then
        Call Time
        MsgBox "Não foi Possível Excluir o Registro."
    Else
        Call Time
        tmrBonificacao.Enabled = True
        stbBonificacao.Panels.Item(1).Text = "Registro Excluído com Êxito."
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
    botaoPressionado = "novo"
End Sub

Private Sub mniSair_Click()
    Unload Me
End Sub

Private Sub mniSalvar_Click()
    If txtCodFunc.Text = Empty Or _
        txtHistorico.Text = Empty Or _
        txtValor.Text = Empty Then
            MsgBox "Por Favor Preencha os Campos em Negrito", _
                vbInformation, "Informação!"
            Exit Sub
    End If
    Dim idReg As Long
    Dim myQuery As String
    idReg = 0
    If botaoPressionado = "novo" Then
        Call Aux.Query("INSERT INTO bonificacao ( id, fid" _
            & " ) VALUES ( null, " & txtCodFunc.Text & ")")
        If Aux.NumErro <> 0 Then
            MsgBox "Não foi Possível Inserir o Registro.", vbCritical, "Falha de Gravação"
            Exit Sub
        Else
            idReg = Aux.GetValue("SELECT MAX(id) FROM bonificacao")
        End If
    Else
        idReg = CLng(txtCodigo.Text)
    End If
    mebEmissao.PromptInclude = True
    myQuery = "UPDATE bonificacao SET" _
        & " fid = " & txtCodFunc.Text _
        & ", emissao = '" & modData.VbToMy(mebEmissao.Text) & "'" _
        & ", historico = '" & txtHistorico.Text & "'" _
        & ", valor = " & Replace(CDbl(txtValor.Text), ",", ".") _
        & ", turno = 0 WHERE id = " & idReg
    mebEmissao.PromptInclude = False
    Call Aux.Query(myQuery)
    If Aux.NumErro <> 0 Then
        MsgBox "Não foi Possível Gravar Registro.", _
            vbCritical, "Falha de Gravação"
        Exit Sub
    Else
        Call Time
        tmrBonificacao.Enabled = True
        stbBonificacao.Panels.Item(1).Text = "Registro Gravado com Êxito."
        Call QueryMain
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

Private Sub tbrBonificacao_ButtonClick(ByVal Button As MSComctlLib.Button)
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

Private Sub tmrBonificacao_Timer()
    Call Time
End Sub

Private Sub txtCodFunc_Change()
    If txtCodFunc.Text = Empty Then GoTo Limpa
    Call Aux.Query("SELECT nome FROM funcionario" _
        & " WHERE id = " & txtCodFunc.Text)
    If Aux.Rs.RecordCount = 0 Then
        txtNomeFunc.Text = "Registro Inexistente."
    Else
        txtNomeFunc.Text = Aux.Rs.Fields.Item("nome").Value
    End If
    Exit Sub
Limpa:
    txtNomeFunc.Text = Empty
End Sub

Private Sub txtCodFunc_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Numeros(KeyAscii)
End Sub

Private Sub txtCodFunc_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not KeyCode = vbKeyF5 Then Exit Sub
    Call Conn.Query("SELECT id, nome, cpf, situ, recvale" _
        & ", limite FROM funcionario")
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

Private Sub txtCodFunc_GotFocus()
    txtCodFunc.SelStart = 0
    txtCodFunc.SelLength = 10
    txtCodFunc.BackColor = &H80000018
    stbBonificacao.Panels.Item(1).Text = "Pressione F5 para Localizar"
    LocalizarOQue = "Funcionário"
End Sub

Private Sub txtCodFunc_LostFocus()
    txtCodFunc.BackColor = &H80000005
    stbBonificacao.Panels.Item(1).Text = ""
    LocalizarOQue = "Bonificação"
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
        Call QueryMain
        Call Time
        Conn.Rs.Find "id = " & txtCodigo.Text
        If Conn.Rs.EOF Then
            tmrBonificacao.Enabled = True
            stbBonificacao.Panels.Item(1).Text = "Registro Inexistente."
        Else
            mebEmissao.Text = IIf(IsNull(Conn.Rs.Fields.Item("emissao").Value), "", Conn.Rs.Fields.Item("emissao").Value)
            txtCodFunc.Text = Conn.Rs.Fields.Item("fid").Value
            txtNomeFunc.Text = Conn.Rs.Fields.Item("nome").Value
            txtHistorico.Text = Conn.Rs.Fields.Item("historico").Value
            txtValor.Text = modMoeda.FmtMoeda(Conn.Rs.Fields.Item("valor").Value)
            Hab = True
            GoTo habilita
        End If
    End If
    mebEmissao.Text = Empty
    txtCodFunc.Text = Empty
    txtNomeFunc.Text = Empty
    txtHistorico.Text = Empty
    txtValor.Text = Empty
    Hab = False
habilita:
    With tbrBonificacao.Buttons
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

Private Sub txtHistorico_GotFocus()
    txtHistorico.BackColor = &H80000018
End Sub

Private Sub txtHistorico_LostFocus()
    txtHistorico.BackColor = &H80000005
End Sub

Private Sub txtHistorico_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Maiusculas(KeyAscii)
End Sub

Private Sub txtNomeFunc_GotFocus()
    txtNomeFunc.SelStart = 0
    txtNomeFunc.SelLength = 50
    txtNomeFunc.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtNomeFunc_LostFocus()
    txtNomeFunc.BackColor = &H80000005 'Branco
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
    mebEmissao.Text = IIf(situ, Empty, Date)
    txtCodigo.Enabled = situ
    txtCodFunc.Enabled = Not situ
    txtHistorico.Enabled = Not situ
    txtValor.Enabled = Not situ
    With tbrBonificacao.Buttons
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
    tbrBonificacao.Buttons.Item("editar").Enabled = False
    tbrBonificacao.Buttons.Item("excluir").Enabled = False
    mniEditar.Enabled = False
    mniExcluir.Enabled = False
    mebEmissao.Text = Empty
    txtCodFunc.Text = Empty
    txtNomeFunc.Text = Empty
    txtHistorico.Text = Empty
    txtValor.Text = Empty
End Sub

Private Sub Time()
    stbBonificacao.Panels.Item(1).Text = Empty
    tmrBonificacao.Interval = 0
    tmrBonificacao.Enabled = False
    tmrBonificacao.Interval = 10000
End Sub

Private Sub QueryMain()
    Call Conn.Query("SELECT bonificacao.id, bonificacao.fid" _
        & ", bonificacao.emissao, bonificacao.historico" _
        & ", bonificacao.valor, bonificacao.turno" _
        & ", funcionario.nome FROM bonificacao" _
        & " LEFT JOIN funcionario ON bonificacao.fid = funcionario.id")
End Sub

Public Sub Load()
    Select Case LocalizarOQue
        Case "Bonificação"
            Call LoadBonificacao
        Case "Funcionário"
            Call LoadFuncionario
    End Select
End Sub

' cmpProc = Campo Procurado
Public Sub Campo(txt As String)
    Select Case LocalizarOQue
        Case "Bonificação"
            Call CampoBonificacao(txt)
        Case "Funcionário"
            Call CampoFuncionario(txt)
    End Select
End Sub

Public Sub LoadBonificacao()
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
    frmLocalizar.Caption = "Localização de Bonificação"
    frmLocalizar.tabela = "bonificacao"
End Sub

Public Sub CampoBonificacao(txt As String)
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

