VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmCliente 
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7125
   Icon            =   "frmCliente.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   475
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFone 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   1320
      MaxLength       =   32
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   1320
      MaxLength       =   60
      TabIndex        =   1
      Top             =   1200
      Width           =   5535
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2280
      Width           =   4695
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Timer tmrCliente 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6120
      Top             =   1800
   End
   Begin MSComctlLib.ImageList imlCliD 
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
            Picture         =   "frmCliente.frx":0CCA
            Key             =   "novo"
            Object.Tag             =   "novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":19A4
            Key             =   "editar"
            Object.Tag             =   "editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":267E
            Key             =   "salvar"
            Object.Tag             =   "salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":3358
            Key             =   "excluir"
            Object.Tag             =   "excluir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":4032
            Key             =   "sair"
            Object.Tag             =   "sair"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":4D0C
            Key             =   "cancelar"
            Object.Tag             =   "cancelar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":59E6
            Key             =   "localizar"
            Object.Tag             =   "localizar"
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
   Begin ComCtl3.CoolBar cbrCliente 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   1111
      BandCount       =   1
      _CBWidth        =   7125
      _CBHeight       =   630
      _Version        =   "6.0.8169"
      Child1          =   "tbrCliente"
      MinHeight1      =   38
      Width1          =   14
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrCliente 
         Height          =   570
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlCliente"
         DisabledImageList=   "imlCliD"
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
   Begin MSComctlLib.StatusBar stbCliente 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   7
      Top             =   2745
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10160
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1852
            MinWidth        =   1852
            Text            =   "F1 - Ajuda"
            TextSave        =   "F1 - Ajuda"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlCliente 
      Left            =   4800
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
            Picture         =   "frmCliente.frx":66C0
            Key             =   "novo"
            Object.Tag             =   "novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":739A
            Key             =   "editar"
            Object.Tag             =   "editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":8074
            Key             =   "salvar"
            Object.Tag             =   "salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":8D4E
            Key             =   "excluir"
            Object.Tag             =   "excluir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":9A28
            Key             =   "sair"
            Object.Tag             =   "sair"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":A702
            Key             =   "cancelar"
            Object.Tag             =   "cancelar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":B3DC
            Key             =   "localizar"
            Object.Tag             =   "localizar"
         EndProperty
      EndProperty
   End
   Begin VB.Label labFone 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1973
      Width           =   405
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
      Left            =   120
      TabIndex        =   11
      Top             =   1260
      Width           =   555
   End
   Begin VB.Label labEmail 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2333
      Width           =   465
   End
   Begin VB.Label labCpf 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CPF:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1620
      Width           =   345
   End
   Begin VB.Label labCodigo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "C�digo:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   900
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
         Caption         =   "&Conte�do"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Conn As clsMyConnect
Public Localizar As Boolean 'Define o retorno da tela localizar
Dim Sair As Boolean
Dim botaoPressionado As String

Private Sub Form_Load()
    Sair = True
    Localizar = False
    botaoPressionado = Empty
    mniEditar.Enabled = False
    mniSalvar.Enabled = False
    mniExcluir.Enabled = False
    mniCancelar.Enabled = False
    tbrCliente.Buttons.Item("editar").Enabled = False
    tbrCliente.Buttons.Item("salvar").Enabled = False
    tbrCliente.Buttons.Item("excluir").Enabled = False
    tbrCliente.Buttons.Item("cancelar").Enabled = False
    Set Conn = New clsMyConnect
    Call Conn.Connect
    If Conn.NumErro <> 0 Then GoTo SubFail
    Call QueryMain
    If Conn.NumErro <> 0 Then GoTo SubFail
    Exit Sub
SubFail:
    MsgBox "N�o foi Poss�vel Conectar-se com a Base de Dados." _
        & Chr(13) & "Erro # " & Str(Conn.NumErro) & " foi gerado por " _
        & Conn.SrcErro & Chr(13) & Conn.DescErro, vbCritical, "Falha de Conex�o"
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
    mebCpf.SelLength = 18
    mebCpf.BackColor = &H80000018 'Amarelo
End Sub

Private Sub mebCpf_LostFocus()
    mebCpf.BackColor = &H80000005 'Branco
End Sub

Private Sub mebCpf_Validate(Cancel As Boolean)
    If mebCpf.Text = Empty Then Exit Sub
    Cancel = Not modCheckCNPJCPF.CheckCPF(mebCpf.Text)
    If Cancel Then
        MsgBox "N�mero de CPF Inv�lido", vbInformation, "CPF Inv�lido"
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
    Call Conn.Query("DELETE FROM cliente WHERE id = " & txtCodigo.Text)
    If Conn.NumErro = -2147217871 Then 'Falha de chave estrangeira
        Call Time
        MsgBox "Cliente possui relacionamentos" _
            & Chr(13) & "e n�o pode ser exclu�do." _
            , vbExclamation, "Falha de Exclus�o"
    ElseIf Not Conn.NumErro = 0 Then
        Call Time
        MsgBox "N�o foi Poss�vel Excluir o Registro."
    Else
        Call Time
        tmrCliente.Enabled = True
        stbCliente.Panels.Item(1).Text = "Registro Exclu�do com �xito."
        Call ZeraValores
        Call QueryMain
        If Conn.NumErro <> 0 Then
            MsgBox "N�o foi Poss�vel Reler os Dados." _
                & Chr(13) & Me.Caption & " Ser� Fechado."
            Unload Me
        End If
    End If
End Sub

Private Sub mniLocalizar_Click()
    If Conn.Rs.RecordCount = 0 Then
        MsgBox "N�o H� Dados Registrados", vbInformation, "Sem Registros"
        Exit Sub
    End If
    Set frmPrin.frmPai = Me
    frmLocalizar.Show vbModal, Me
    If Localizar = True Then
        txtCodigo.Text = Conn.Rs.Fields.Item("id").Value
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
    If txtNome.Text = Empty Then
            MsgBox "Por Favor Preencha os Campos em Negrito"
            Exit Sub
    End If
    Dim idReg As Long
    Dim myQuery As String
    idReg = 0
    If botaoPressionado = "novo" Then
        Call Conn.Query("INSERT INTO cliente(id) VALUES(null)")
        If Conn.NumErro <> 0 Then
            MsgBox "N�o foi Poss�vel Inserir o Registro.", _
                vbCritical, "Falha de Grava��o"
            Exit Sub
        Else
            idReg = Conn.GetValue("SELECT MAX(id) FROM cliente")
        End If
    Else
        idReg = CLng(txtCodigo.Text)
    End If
    myQuery = "UPDATE cliente SET" _
        & " nome = '" & txtNome.Text & "'" _
        & ", cpf = '" & mebCpf.Text & "'" _
        & ", fone = '" & txtFone.Text & "'" _
        & ", email = '" & txtEmail.Text & "'" _
        & " WHERE id = " & idReg
    Call Conn.Query(myQuery)
    If Conn.NumErro <> 0 Then
        MsgBox "N�o foi Poss�vel Gravar Registro.", _
            vbCritical, "Falha de Grava��o"
        Exit Sub
    Else
        Call Time
        tmrCliente.Enabled = True
        stbCliente.Panels.Item(1).Text = "Registro Gravado com �xito."
        Call QueryMain
        If Conn.NumErro <> 0 Then
            MsgBox "N�o foi Poss�vel Reler os Dados." _
                & Chr(13) & Me.Caption & " Ser� Fechado."
        End If
    End If
    Call HabCampos(True)
    If botaoPressionado = "novo" Then
        Call ZeraValores
    End If
    botaoPressionado = Empty
    txtCodigo.SetFocus
End Sub

Private Sub tbrCliente_ButtonClick(ByVal Button As MSComctlLib.Button)
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

Private Sub tmrCliente_Timer()
    Call Time
End Sub

Private Sub txtCodigo_Change()
    Dim Hab As Boolean
    If Not txtCodigo.Text = Empty Then
        Call QueryMain
        Call Time
        Conn.Rs.Find "id = " & txtCodigo.Text
        If Conn.Rs.BOF Or Conn.Rs.EOF Then
            tmrCliente.Enabled = True
            stbCliente.Panels.Item(1).Text = "Registro Inexistente."
        Else
            txtNome.Text = Conn.Rs.Fields.Item("nome").Value
            mebCpf.Text = Conn.Rs.Fields.Item("cpf").Value
            txtFone.Text = Conn.Rs.Fields.Item("fone").Value
            txtEmail.Text = Conn.Rs.Fields.Item("email").Value
            Hab = True
            GoTo habilita
        End If
    End If
    txtNome.Text = Empty
    mebCpf.Text = Empty
    txtFone.Text = Empty
    txtEmail.Text = Empty
    Hab = False
habilita:
    With tbrCliente.Buttons
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

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Numeros(KeyAscii)
End Sub

Private Sub txtCodigo_GotFocus()
    txtCodigo.SelStart = 0
    txtCodigo.SelLength = 10
    txtCodigo.BackColor = &H80000018
End Sub

Private Sub txtCodigo_LostFocus()
    txtCodigo.BackColor = &H80000005
End Sub

Private Sub txtEmail_GotFocus()
    txtEmail.SelStart = 0
    txtEmail.SelLength = 12
    txtEmail.BackColor = &H80000018
End Sub

Private Sub txtEmail_LostFocus()
    txtEmail.BackColor = &H80000005
End Sub

Private Sub txtFone_GotFocus()
    txtFone.SelStart = 0
    txtFone.SelLength = 32
    txtFone.BackColor = &H80000018
End Sub

Private Sub txtFone_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.LetrasENumeros(KeyAscii)
End Sub

Private Sub txtFone_LostFocus()
    txtFone.BackColor = &H80000005
    txtFone.Text = modFone.fone(txtFone.Text)
End Sub

Private Sub txtNome_GotFocus()
    txtNome.SelStart = 0
    txtNome.SelLength = 50
    txtNome.BackColor = &H80000018
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Maiusculas(KeyAscii)
End Sub

Private Sub txtNome_LostFocus()
    txtNome.BackColor = &H80000005
End Sub

Private Sub HabCampos(situ As Boolean)
    txtCodigo.Enabled = situ
    txtNome.Enabled = Not situ
    mebCpf.Enabled = Not situ
    txtFone.Enabled = Not situ
    txtEmail.Enabled = Not situ
    With tbrCliente.Buttons
        .Item("salvar").Enabled = Not situ
        .Item("localizar").Enabled = situ
        .Item("sair").Enabled = situ
        .Item("novo").Enabled = situ
        .Item("cancelar").Enabled = Not situ
    End With
    mniSalvar.Enabled = Not situ
    mniLocalizar.Enabled = situ
    mniSair.Enabled = situ
    mniCancelar.Enabled = Not situ
    mniNovo.Enabled = situ
    Sair = situ
End Sub

Private Sub ZeraValores()
    tbrCliente.Buttons.Item("editar").Enabled = False
    tbrCliente.Buttons.Item("excluir").Enabled = False
    mniEditar.Enabled = False
    mniExcluir.Enabled = False
    txtCodigo.Text = Empty
    txtNome.Text = Empty
    mebCpf.Text = Empty
    txtFone.Text = Empty
    txtEmail.Text = Empty
End Sub

Private Sub Time()
    stbCliente.Panels.Item(1).Text = Empty
    tmrCliente.Interval = 0
    tmrCliente.Enabled = False
    tmrCliente.Interval = 10000
End Sub

Public Sub Load()
    With frmLocalizar.cbbProcura
        .AddItem ("C�digo")
        .AddItem ("Nome")
        .AddItem ("CPF")
        .Text = "Nome"
    End With
    With frmLocalizar.mfgLocalizar
        .Cols = 5
        .Rows = 1
        .Clear
        .ColWidth(0) = 1000
        .ColWidth(1) = 4000
        .ColWidth(2) = 2100
        .ColWidth(3) = 1200
        .ColWidth(4) = 4000
        .TextArray(0) = "C�digo"
        .TextArray(1) = "Nome"
        .TextArray(2) = "CPF"
        .TextArray(3) = "Telefone"
        .TextArray(4) = "E-mail"
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
        frmLocalizar.mfgLocalizar.TextMatrix(Rows, 2) = Conn.Rs.Fields.Item("cpf").Value
        fieldWidth = TextWidth(Conn.Rs.Fields.Item("cpf").Value)
        If colsWidth(2) < fieldWidth Then
            colsWidth(2) = fieldWidth
        End If
        frmLocalizar.mfgLocalizar.TextMatrix(Rows, 3) = Conn.Rs.Fields.Item("fone").Value
        fieldWidth = TextWidth(Conn.Rs.Fields.Item("fone").Value)
        If colsWidth(3) < fieldWidth Then
            colsWidth(3) = fieldWidth
        End If
        frmLocalizar.mfgLocalizar.TextMatrix(Rows, 4) = Conn.Rs.Fields.Item("email").Value
        fieldWidth = TextWidth(Conn.Rs.Fields.Item("email").Value)
        If colsWidth(4) < fieldWidth Then
            colsWidth(4) = fieldWidth
        End If
        Conn.Rs.MoveNext
        Rows = Rows + 1
    Loop
    frmLocalizar.Caption = "Localiza��o de Cliente"
    frmLocalizar.tabela = "cliente"
End Sub

' cmpProc = Campo Procurado
Public Sub Campo(txt As String)
    Select Case txt
        Case "C�digo"
            frmLocalizar.Campo = "id"
        Case "Nome"
            frmLocalizar.Campo = "nome"
        Case "CPF"
            frmLocalizar.Campo = "cpf"
    End Select
End Sub

Private Sub txtContato_LostFocus()
    'txtContato.BackColor = &H80000005
End Sub

Private Sub QueryMain()
    Call Conn.Query("SELECT id, nome, cpf, fone, email FROM cliente")
End Sub
