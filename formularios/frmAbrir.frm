VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmAbrir 
   Caption         =   "Abrir Período"
   ClientHeight    =   2655
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmAbrir.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrAbrir 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3120
      Top             =   1440
   End
   Begin MSMask.MaskEdBox mebHoraFin 
      Height          =   300
      Left            =   1680
      TabIndex        =   3
      Top             =   1800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   5
      Format          =   "hh:mm"
      Mask            =   "99:99"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mebHoraIni 
      Height          =   300
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   5
      Format          =   "hh:mm"
      Mask            =   "99:99"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mebDataFin 
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mebDataIni 
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin ComCtl3.CoolBar cbrAbrir 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1111
      BandCount       =   1
      _CBWidth        =   4680
      _CBHeight       =   630
      _Version        =   "6.7.9782"
      Child1          =   "tbrAbrir"
      MinHeight1      =   38
      Width1          =   209
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrAbrir 
         Height          =   570
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlAbrir"
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
   Begin MSComctlLib.StatusBar stbAbrir 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   2310
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5874
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1852
            MinWidth        =   1852
            Text            =   "F1 - Ajuda"
            TextSave        =   "F1 - Ajuda"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlAbrir 
      Left            =   4080
      Top             =   720
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
            Picture         =   "frmAbrir.frx":038A
            Key             =   "salvar"
            Object.Tag             =   "salvar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbrir.frx":1064
            Key             =   "sair"
            Object.Tag             =   "sair"
         EndProperty
      EndProperty
   End
   Begin VB.Label labFim 
      AutoSize        =   -1  'True
      Caption         =   "Data e Hora Final"
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
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label labInicio 
      AutoSize        =   -1  'True
      Caption         =   "Data e Hora Inicial"
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
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1620
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
      Caption         =   "Aj&uda"
      Begin VB.Menu mniConteudo 
         Caption         =   "&Conteúdo"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmAbrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conn As clsMyConnect

Private Sub Form_Load()
    Dim PerId As Long
    Set Conn = New clsMyConnect
    Call Conn.Connect
    If Conn.NumErro <> 0 Then GoTo SubFail
    PerId = Conn.GetValue("SELECT MAX(id) FROM periodo")
    If Conn.NumErro <> 0 Then GoTo SubFail
    Call Conn.Query("SELECT dataIni, dataFin, horaIni, " _
        & "horaFin, situ FROM periodo WHERE id = " _
        & IIf(PerId = Empty, 1, PerId))
    If Conn.NumErro <> 0 Then GoTo SubFail
    If Conn.Rs.RecordCount = 0 Then GoTo Zerar
    If Conn.Rs.Fields.Item("situ").Value = 1 Then
        mebDataIni.Enabled = False
        mebHoraIni.Enabled = False
        mebDataIni.Text = Conn.Rs.Fields.Item("dataIni").Value
        mebDataFin.Text = Conn.Rs.Fields.Item("dataFin").Value
        mebHoraIni.Text = Conn.Rs.Fields.Item("horaIni").Value
        mebHoraFin.Text = Conn.Rs.Fields.Item("horaFin").Value
        Exit Sub
    End If
Zerar:
    mebDataIni.Text = Date
    mebDataFin.Text = Date + 1
    mebHoraIni.Text = Time
    mebHoraFin.Text = "0200"
    Exit Sub
SubFail:
    MsgBox "Não foi Possível Conectar-se com a Base de Dados." _
        & Chr(13) & "Erro # " & Str(Conn.NumErro) & " foi gerado por " _
        & Conn.SrcErro & Chr(13) & Conn.DescErro, vbCritical, "Falha de Conexão"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Conn.Disconnect
End Sub

Private Sub mebDataFin_GotFocus()
    mebDataFin.BackColor = &H80000018 'Amarelo
    mebDataFin.Tag = mebDataIni.Text
    mebDataFin.SelStart = 0
    mebDataFin.SelLength = 10
End Sub

Private Sub mebDataFin_LostFocus()
    mebDataFin.BackColor = &H80000005 'Branco
    If mebDataFin.Tag = mebDataFin.Text Then Exit Sub
    mebDataFin.Text = modData.FmtData(mebDataFin.Text)
End Sub

Private Sub mebDataIni_GotFocus()
    mebDataIni.Tag = mebDataIni.Text
    mebDataIni.SelStart = 0
    mebDataIni.SelLength = 10
    mebDataIni.BackColor = &H80000018 'Amarelo
End Sub

Private Sub mebDataIni_LostFocus()
    mebDataIni.BackColor = &H80000005 'Branco
    If mebDataIni.Tag = mebDataIni.Text Then Exit Sub
    mebDataIni.Text = modData.FmtData(mebDataIni.Text)
End Sub

Private Sub mebHoraFin_GotFocus()
    mebHoraFin.BackColor = &H80000018 'Amarelo
    mebHoraFin.Tag = mebDataIni.Text
    mebHoraFin.SelStart = 0
    mebHoraFin.SelLength = 5
End Sub

Private Sub mebHoraFin_LostFocus()
    mebHoraFin.BackColor = &H80000005 'Branco
    If mebHoraFin.Tag = mebHoraFin.Text Then Exit Sub
    mebHoraFin.Text = modData.FmtHora(mebHoraFin.Text)
End Sub

Private Sub mebHoraIni_GotFocus()
    mebHoraIni.BackColor = &H80000018 'Amarelo
    mebHoraIni.Tag = mebDataIni.Text
    mebHoraIni.SelStart = 0
    mebHoraIni.SelLength = 5
End Sub

Private Sub mebHoraIni_LostFocus()
    mebHoraIni.BackColor = &H80000005 'Branco
    If mebHoraIni.Tag = mebHoraIni.Text Then Exit Sub
    mebHoraIni.Text = modData.FmtHora(mebHoraIni.Text)
End Sub

Private Sub mniSair_Click()
    Unload Me
End Sub

Private Sub mniSalvar_Click()
    If mebDataIni.Text = Empty Or mebHoraIni.Text = Empty Or _
        mebDataFin.Text = Empty Or mebHoraFin.Text = Empty Then
        MsgBox "Por Favor Preencha Todos os Campos.", _
            vbInformation, "Preencher Campos"
        Exit Sub
    End If
    If mebDataIni.Enabled Then
        Beep
        If MsgBox("ATENÇÃO: Depois de Aberto o Período, a Hora e" _
            & Chr(13) & "Data Iniciais Não Poderão Ser Alteradas." _
            & Chr(13) & "Deseja Continuar?", vbQuestion + vbYesNo _
            + vbDefaultButton2, "Gravar Período?") = vbNo Then
            Exit Sub
        End If
        Call Conn.Query("INSERT INTO periodo ( id, dataIni, dataFin, " _
            & "horaIni, horaFin, situ, dataAbe, horaAbe ) " _
            & "VALUES (null, '" & mebDataIni.Text & "', '" & mebDataFin.Text _
            & "', '" & mebHoraIni.Text & "', '" & mebHoraFin.Text _
            & "', 1, '" & Mid(Date, 1, 2) & Mid(Date, 4, 2) & Mid(Date, 7, 4) _
            & "', '" & Mid(Time, 1, 2) & Mid(Time, 4, 2) & "')")
        If Conn.NumErro <> 0 Then GoTo SubFail
    Else
        Call Conn.Query("UPDATE periodo SET dataFin = '" _
            & mebDataFin.Text & "', horaFin = '" & mebHoraFin.Text & "' " _
            & "WHERE id = " & Conn.GetValue("SELECT MAX(id) FROM periodo"))
        If Conn.NumErro <> 0 Then GoTo SubFail
    End If
    Call Form_Load
    tmrAbrir.Enabled = True
    stbAbrir.Panels.Item(1).Text = "Registro Salvo com Êxito."
    Exit Sub
SubFail:
    MsgBox "Não foi Possível Gravar os Dados." _
        & Chr(13) & "Erro # " & Str(Conn.NumErro) & " foi gerado por " _
        & Conn.SrcErro & Chr(13) & Conn.DescErro, vbCritical, "Falha de Gravação"
End Sub

Private Sub tbrAbrir_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "salvar"
            mniSalvar_Click
        Case "sair"
            mniSair_Click
    End Select
End Sub

Private Sub tmrAbrir_Timer()
    tmrAbrir.Enabled = False
    stbAbrir.Panels.Item(1).Text = Empty
End Sub
