VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmFuncionario 
   Caption         =   "Cadastro de Funcionários"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6735
   Icon            =   "frmFuncionario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbbRecVale 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox cbbSituacao 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtSenha 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtConfirmSenha 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Timer tmrFuncionario 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   4680
      Top             =   2520
   End
   Begin VB.TextBox txtLimite 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   5295
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin MSComctlLib.ImageList imlFuncD 
      Left            =   5760
      Top             =   2400
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
            Picture         =   "frmFuncionario.frx":0CCA
            Key             =   "novo"
            Object.Tag             =   "novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFuncionario.frx":19A4
            Key             =   "editar"
            Object.Tag             =   "editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFuncionario.frx":267E
            Key             =   "salvar"
            Object.Tag             =   "salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFuncionario.frx":3358
            Key             =   "excluir"
            Object.Tag             =   "excluir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFuncionario.frx":4032
            Key             =   "sair"
            Object.Tag             =   "sair"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFuncionario.frx":4D0C
            Key             =   "cancelar"
            Object.Tag             =   "cancelar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFuncionario.frx":59E6
            Key             =   "localizar"
            Object.Tag             =   "localizar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlFuncionario 
      Left            =   5160
      Top             =   2400
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
            Picture         =   "frmFuncionario.frx":66C0
            Key             =   "novo"
            Object.Tag             =   "novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFuncionario.frx":739A
            Key             =   "editar"
            Object.Tag             =   "editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFuncionario.frx":8074
            Key             =   "salvar"
            Object.Tag             =   "salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFuncionario.frx":8D4E
            Key             =   "excluir"
            Object.Tag             =   "excluir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFuncionario.frx":9A28
            Key             =   "sair"
            Object.Tag             =   "sair"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFuncionario.frx":A702
            Key             =   "cancelar"
            Object.Tag             =   "cancelar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFuncionario.frx":B3DC
            Key             =   "localizar"
            Object.Tag             =   "localizar"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrFuncionario 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1111
      BandCount       =   1
      _CBWidth        =   6735
      _CBHeight       =   630
      _Version        =   "6.7.9782"
      Child1          =   "tbrFuncionario"
      MinHeight1      =   38
      Width1          =   209
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrFuncionario 
         Height          =   570
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlFuncionario"
         DisabledImageList=   "imlFuncD"
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
   Begin MSComctlLib.StatusBar stbFuncionario 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   6
      Top             =   3135
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9499
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1852
            MinWidth        =   1852
            Text            =   "F1 - Ajuda"
            TextSave        =   "F1 - Ajuda"
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox mebCpf 
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
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
   Begin VB.Label labSenha 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
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
      Left            =   120
      TabIndex        =   16
      Top             =   2340
      Width           =   615
   End
   Begin VB.Label labConfirmSenha 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar:"
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
      Left            =   120
      TabIndex        =   15
      Top             =   2700
      Width           =   870
   End
   Begin VB.Label labRecVale 
      AutoSize        =   -1  'True
      Caption         =   "Recebe Vale:"
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
      Left            =   3000
      TabIndex        =   14
      Top             =   1980
      Width           =   1170
   End
   Begin VB.Label labSituacao 
      AutoSize        =   -1  'True
      Caption         =   "Situação:"
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
      Left            =   3000
      TabIndex        =   13
      Top             =   1620
      Width           =   825
   End
   Begin VB.Label labCpf 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CPF:"
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
      Left            =   120
      TabIndex        =   12
      Top             =   1620
      Width           =   420
   End
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
      TabIndex        =   11
      Top             =   1980
      Width           =   570
   End
   Begin VB.Label labNome 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1260
      Width           =   555
   End
   Begin VB.Label labCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   893
      Width           =   540
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
Attribute VB_Name = "frmFuncionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Conn As clsMyConnect
Public Localizar As Boolean
Dim Sair As Boolean
Dim botaoPressionado As String
Dim Cript As clsCrypt

Private Sub cbbRecVale_GotFocus()
    cbbRecVale.BackColor = &H80000018 'Amarelo
End Sub

Private Sub cbbRecVale_LostFocus()
    cbbRecVale.BackColor = &H80000005 'Branco
End Sub

Private Sub cbbSituacao_GotFocus()
    cbbSituacao.BackColor = &H80000018 'Amarelo
End Sub

Private Sub cbbSituacao_LostFocus()
    cbbSituacao.BackColor = &H80000005 'Branco
End Sub

Private Sub Form_Load()
    cbbSituacao.AddItem "Demitido"
    cbbSituacao.AddItem "Admitido"
    cbbSituacao.Text = "Admitido"
    cbbRecVale.AddItem "Não"
    cbbRecVale.AddItem "Sim"
    cbbRecVale.Text = "Sim"
    Sair = True
    Localizar = False
    botaoPressionado = Empty
    mniEditar.Enabled = False
    mniSalvar.Enabled = False
    mniExcluir.Enabled = False
    mniCancelar.Enabled = False
    tbrFuncionario.Buttons.Item("editar").Enabled = False
    tbrFuncionario.Buttons.Item("salvar").Enabled = False
    tbrFuncionario.Buttons.Item("excluir").Enabled = False
    tbrFuncionario.Buttons.Item("cancelar").Enabled = False
    Set Conn = New clsMyConnect
    Set Cript = New clsCrypt
    Call Conn.Connect
    If Conn.NumErro <> 0 Then GoTo SubFail
    Call Conn.Query("SELECT id, nome, cpf, situ, recvale, limite, senha FROM funcionario")
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

Private Sub mebCpf_GotFocus()
    mebCpf.SelStart = 0
    mebCpf.SelLength = 14
    mebCpf.BackColor = &H80000018 'Amarela
End Sub

Private Sub mebCpf_LostFocus()
    mebCpf.BackColor = &H80000005 'Branca
End Sub

Private Sub mebCpf_Validate(Cancel As Boolean)
    If mebCpf.Text = Empty Then Exit Sub
    Cancel = Not modCheckCNPJCPF.CheckCPF(mebCpf.Text)
    If Cancel Then
        MsgBox "Número de CPF Inválido", vbInformation, "CPF Inválido"
    End If
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
    txtNome.SetFocus
End Sub

Private Sub mniExcluir_Click()
    If MsgBox("Deseja Excluir o Registro Selecionado?", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Excluir Registro.") = vbNo Then
        Exit Sub
    End If
    Call HabCampos(True)
    Call Conn.Query("DELETE FROM funcionario WHERE id = " & txtCodigo.Text)
    If Conn.NumErro = -2147217871 Then 'Falha de chave estrangeira
        Call Time
        MsgBox "Funcionário possui relacionamentos" _
            & Chr(13) & "e não pode ser excluído." _
            , vbExclamation, "Falha de Exclusão"
    ElseIf Not Conn.NumErro = 0 Then
        Call Time
        MsgBox "Não foi Possível Excluir o Registro.", vbExclamation, "Falha de Exclusão"
    Else
        Call Time
        tmrFuncionario.Enabled = True
        stbFuncionario.Panels.Item(1).Text = "Registro Excluído com Êxito."
        Call ZeraValores
        Call Conn.Query("SELECT id, nome, situ, recvale, limite, senha FROM funcionario")
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
    txtNome.SetFocus
End Sub

Private Sub mniSair_Click()
    Unload Me
End Sub

Private Sub mniSalvar_Click()
    If txtNome.Text = Empty Or _
        mebCpf.Text = Empty Or _
        txtLimite.Text = Empty Or _
        txtSenha.Text = Empty Or _
        txtConfirmSenha.Text = Empty Then
            MsgBox "Por Favor Preencha os Campos em Negrito", _
                vbInformation, "Informação!"
            Exit Sub
    End If
    If Not txtSenha.Text = txtConfirmSenha.Text Then
        MsgBox "Os Campos 'Senha' e 'Confirmar Senha' devem ser iguais"
        Exit Sub
    End If
    If Len(txtSenha.Text) < 6 Then
        MsgBox "A Senha Deve ter Comprimento Mínimo de 6 Caracteres"
        Exit Sub
    End If
    Dim idReg As Long
    Dim myQuery As String
    idReg = 0
    If botaoPressionado = "novo" Then
        If mebCpf.Text = Conn.GetValue( _
            "SELECT cpf FROM funcionario WHERE cpf = '" & mebCpf.Text & "'") Then
            MsgBox "Funcionário: " _
                & Conn.GetValue("SELECT nome FROM funcionario WHERE cpf = '" _
                & mebCpf.Text & "'") & Chr(13) & "CPF: " _
                & Conn.GetValue("SELECT cpf FROM funcionario WHERE cpf = '" _
                & mebCpf.Text & "'") & Chr(13) _
                & " já Cadastrado.", vbInformation, "Funcionário Cadastrado"
            Call Conn.Query("SELECT id, nome, cpf, situ, recvale, limite, senha FROM funcionario")
            If Conn.NumErro <> 0 Then
                MsgBox "Não foi Possível Reler os Dados." _
                    & Chr(13) & Me.Caption & " Será Fechado."
            End If
            Exit Sub
        End If
        Call Conn.Query("INSERT INTO funcionario(id) VALUES(null)")
        If Conn.NumErro <> 0 Then
            MsgBox "Não foi Possível Inserir o Registro.", vbCritical, "Falha de Gravação"
            Exit Sub
        Else
            idReg = Conn.GetValue("SELECT MAX(id) FROM funcionario")
        End If
    Else
        'Converte Caracter para Tipo Long
        idReg = CLng(txtCodigo.Text)
    End If
    myQuery = "UPDATE funcionario SET" _
        & " nome = '" & txtNome.Text & "'" _
        & ", cpf = '" & mebCpf.Text & "'" _
        & ", situ = " & cbbSituacao.ListIndex _
        & ", recvale = " & cbbRecVale.ListIndex _
        & ", limite = " & Replace(CDbl(txtLimite.Text), ",", ".") _
        & ", senha = '" & Cript.CriptSenha(txtSenha.Text) & "'" _
        & " WHERE id = " & idReg
    Call Conn.Query(myQuery)
    If Conn.NumErro <> 0 Then
        MsgBox "Não foi Possível Gravar Registro.", _
            vbCritical, "Falha de Gravação"
        Exit Sub
    Else
        Call Time
        tmrFuncionario.Enabled = True
        stbFuncionario.Panels.Item(1).Text = "Registro Gravado com Êxito."
        Call Conn.Query("SELECT id, nome, cpf, situ, recvale, limite, senha FROM funcionario")
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

Private Sub tbrFuncionario_ButtonClick(ByVal Button As MSComctlLib.Button)
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

Private Sub tmrFuncionario_Timer()
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
        Call Conn.Query("SELECT id, nome, cpf, situ, recvale, limite, senha FROM funcionario")
        Call Time
        Conn.Rs.Find "id = " & txtCodigo.Text
        If Conn.Rs.BOF Or Conn.Rs.EOF Then
            tmrFuncionario.Enabled = True
            stbFuncionario.Panels.Item(1).Text = "Registro Inexistente."
        Else
            txtNome.Text = Conn.Rs.Fields.Item("nome").Value
            mebCpf.Text = Conn.Rs.Fields.Item("cpf").Value
            cbbSituacao.ListIndex = Conn.Rs.Fields.Item("situ").Value
            cbbRecVale.ListIndex = Conn.Rs.Fields.Item("recvale").Value
            txtLimite.Text = Conn.Rs.Fields.Item("limite").Value
            txtSenha.Text = Cript.DeCriptSenha(Conn.Rs.Fields.Item("senha").Value)
            txtConfirmSenha.Text = Cript.DeCriptSenha(Conn.Rs.Fields.Item("senha").Value)
            Call txtLimite_LostFocus
            Hab = True
            GoTo habilita
        End If
    End If
    txtNome.Text = Empty
    mebCpf.Text = Empty
    cbbSituacao.ListIndex = 1
    cbbRecVale.ListIndex = 1
    txtLimite.Text = Empty
    txtSenha.Text = Empty
    txtConfirmSenha.Text = Empty
    Hab = False
habilita:
    With tbrFuncionario.Buttons
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

Private Sub txtConfirmSenha_GotFocus()
    txtConfirmSenha.SelStart = 0
    txtConfirmSenha.SelLength = 12
    txtConfirmSenha.BackColor = &H80000018
End Sub

Private Sub txtConfirmSenha_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.LetrasENumeros(KeyAscii)
End Sub

Private Sub txtConfirmSenha_LostFocus()
    txtConfirmSenha.BackColor = &H80000005
End Sub

Private Sub txtLimite_GotFocus()
    txtLimite.SelStart = 0
    txtLimite.SelLength = 13
    txtLimite.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtLimite_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Moeda(KeyAscii)
End Sub

Private Sub txtLimite_LostFocus()
    txtLimite.BackColor = &H80000005 'Branco
    txtLimite.Text = modMoeda.FmtMoeda(txtLimite.Text)
End Sub

Private Sub txtNome_GotFocus()
    txtNome.SelStart = 0
    txtNome.SelLength = 50
    txtNome.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Maiusculas(KeyAscii)
End Sub

Private Sub txtNome_LostFocus()
    txtNome.BackColor = &H80000005 'Branco
End Sub

Private Sub HabCampos(situ As Boolean)
    txtCodigo.Enabled = situ
    txtNome.Enabled = Not situ
    mebCpf.Enabled = Not situ
    cbbSituacao.Enabled = Not situ
    cbbRecVale.Enabled = Not situ
    txtLimite.Enabled = Not situ
    txtSenha.Enabled = Not situ
    txtConfirmSenha.Enabled = Not situ
    With tbrFuncionario.Buttons
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
    tbrFuncionario.Buttons.Item("editar").Enabled = False
    tbrFuncionario.Buttons.Item("excluir").Enabled = False
    mniEditar.Enabled = False
    mniExcluir.Enabled = False
    txtCodigo.Text = Empty
    txtNome.Text = Empty
    mebCpf.Text = Empty
    txtLimite.Text = Empty
    txtSenha.Text = Empty
    txtConfirmSenha.Text = Empty
    cbbSituacao.ListIndex = 1
    cbbRecVale.ListIndex = 1
End Sub

Private Sub Time()
    stbFuncionario.Panels.Item(1).Text = Empty
    tmrFuncionario.Interval = 0
    tmrFuncionario.Enabled = False
    tmrFuncionario.Interval = 10000
End Sub

Public Sub Load()
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

' cmpProc = Campo Procurado
Public Sub Campo(txt As String)
    Select Case txt
        Case "Código"
            frmLocalizar.Campo = "id"
        Case "Nome"
            frmLocalizar.Campo = "nome"
        Case "CPF"
            frmLocalizar.Campo = "cpf"
    End Select
End Sub

Private Sub txtSenha_GotFocus()
    txtSenha.SelStart = 0
    txtSenha.SelLength = 12
    txtSenha.BackColor = &H80000018
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.LetrasENumeros(KeyAscii)
End Sub

Private Sub txtSenha_LostFocus()
    txtSenha.BackColor = &H80000005
End Sub
