VERSION 5.00
Begin VB.Form frmBackup 
   Caption         =   "Utilitário de Segurança"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   Icon            =   "frmBackup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBackup 
      Caption         =   "Cópia de Segurança"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   5775
      Begin VB.OptionButton optRestaurar 
         Caption         =   "R&estaurar Backup"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton optCriar 
         Caption         =   "Criar &Backup"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox txtBackup 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   4935
      End
      Begin VB.CommandButton btnAbrir 
         Height          =   300
         Left            =   5280
         Picture         =   "frmBackup.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   340
      End
      Begin VB.Label labBackup 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   540
      End
   End
   Begin VB.CommandButton btnBackup 
      Caption         =   "&Criar"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton btnFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "frmBackup.frx":1054
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Dim Table As clsMyConnect
Dim Field As clsMyConnect
Dim Data As clsMyConnect
Dim Arq As Integer

Private Sub btnAbrir_Click()
    Dim ret As String
    If optCriar.Value = True Then
        ret = modGetDir.GetDir(Mid(Table.GetValue( _
            "SELECT valor FROM config WHERE campo = 'Backup';"), 2))
    Else
        'ret = Application.GetOpenFilename( _
        '    "Arquivo de Backup do Tofu (*.arj),*.arj", 1, _
        '    "Localizar Arquivo de Backup")
    End If
    If Not ret = "False" And Not ret = Empty Then
        txtBackup.Text = ret
    End If
End Sub

Private Sub btnBackup_Click()
    Call HabCampos(False)
    If txtBackup.Text = Empty Then
        MsgBox "Por Favor Informe o Local de " & labBackup.Caption, _
            vbInformation, labBackup.Caption
    Else
        frmBackup.MousePointer = 11
        If optCriar.Value = True Then
            Call Criar
        ElseIf optRestaurar.Value = True Then
            Call Restaurar
        End If
        frmBackup.MousePointer = 0
    End If
    Call HabCampos(True)
End Sub

Private Sub btnFechar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set Table = New clsMyConnect
    Set Field = New clsMyConnect
    Set Data = New clsMyConnect
    Call Table.Connect
    Call Field.Connect
    Call Data.Connect
    txtBackup.Text = Mid(Data.GetValue("SELECT valor FROM config " _
        & "WHERE campo = 'Backup'"), 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Table.Disconnect
    Field.Disconnect
    Data.Disconnect
End Sub

Private Sub optCriar_Click()
    btnBackup.Caption = "&Criar"
    labBackup.Caption = "Destino"
End Sub

Private Sub optRestaurar_Click()
    btnBackup.Caption = "&Restaurar"
    labBackup.Caption = "Origem"
End Sub

Private Sub Criar()
    ' Disponibiliza o próximo número de arquivo disponível
    Arq = FreeFile
    ' Abre Arquivo Para Gravação
    On Error GoTo Erro
    Open App.Path & "\Backup.sql" For Output As #Arq
    GoTo Continua
Erro:
    If Err.Number = 76 Then 'Não consegue criar o arquivo
        MsgBox "O Caminho de " & labBackup.Caption & " é Inválido", _
            vbInformation, labBackup.Caption & " Inválido"
    ElseIf Err.Number = 75 Then 'Diretório "c:\TmpBackup" já existe
        RmDir "c:\TmpBackup"
        GoTo Finalize
    Else
        MsgBox "Não foi Possível Gravar o Arquivo." _
            & Chr(13) & "Erro # " & Str(Err.Number) _
            & " foi gerado por " & Err.Source & Chr(13) _
            & Err.Description, vbCritical, "Falha de Conexão"
        ChDir App.Path
        Kill App.Path & "\Backup.sql"
        If Not Dir("c:\TmpTofu\*.*") = Empty Then
            Kill "c:\TmpTofu\*.*"
        End If
        RmDir "c:\TmpTofu"
    End If
    On Error GoTo 0
    Exit Sub
Continua:
    Call Table.Query("show tables")
    While Not Table.Rs.EOF
        Call Field.Query("describe " _
            & Table.Rs.Fields.Item("Tables_in_tofu").Value)
        Dim sql As String
        sql = "SELECT id"
        Field.Rs.MoveNext
        While Not Field.Rs.EOF
            sql = sql & ", " & Field.Rs.Fields.Item("Field").Value
            Field.Rs.MoveNext
        Wend
        sql = sql & " FROM " & Table.Rs.Fields.Item("Tables_in_tofu").Value
        Call Data.Query(sql)
        While Not Data.Rs.EOF
            sql = "INSERT INTO " & Table.Rs.Fields.Item("Tables_in_tofu").Value
            sql = sql & " ( id"
            Field.Rs.MoveFirst
            Field.Rs.MoveNext
            While Not Field.Rs.EOF
                sql = sql & ", " & Field.Rs.Fields.Item("Field").Value
                Field.Rs.MoveNext
            Wend
            sql = sql & " ) VALUES ( "
            Field.Rs.MoveFirst
            Field.Rs.MoveNext
            Dim I As Integer
            Dim Text As Integer
            Dim Char As Integer
            sql = sql & Data.Rs.Fields.Item("id").Value
            For I = 1 To Data.Rs.Fields.Count - 1
                sql = sql & ", "
                Text = InStr(1, Field.Rs.Fields.Item("Type").Value, "text")
                Char = InStr(1, Field.Rs.Fields.Item("Type").Value, "char")
                If Text = 0 And Char = 0 Then
                    sql = sql & Replace(Data.Rs.Fields.Item(I).Value, ",", ".")
                Else
                    sql = sql & "'" & Data.Rs.Fields.Item(I).Value & "'"
                End If
                Field.Rs.MoveNext
            Next
            sql = sql & " );"
            Print #Arq, sql
            Data.Rs.MoveNext
        Wend
        Table.Rs.MoveNext
    Wend
    Close #Arq
Finalize:
    MkDir "c:\TmpBackup"
    FileCopy App.Path & "\Backup.sql", "c:\TmpBackup\Backup.sql"
    FileCopy App.Path & "\arj.exe", "c:\TmpBackup\arj.exe"
    Kill App.Path & "\Backup.sql"
    
    ChDir "c:\TmpBackup"
    iTask = Shell("arj a -va Backup.arj " _
        & "Backup.sql", vbHide)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    ret = WaitForSingleObject(pHandle, INFINITE)
    ret = CloseHandle(pHandle)
    ChDir App.Path
    
    Dim x As String
    x = Table.GetValue("SELECT valor FROM config WHERE campo = 'Backup';")
    If Mid(x, 1, 1) = "0" Then
        Call Table.Query("UPDATE config SET valor = '0" _
            & txtBackup.Text & "' WHERE campo = 'Backup'")
    End If
    x = txtBackup.Text & "\Backup" _
        & Mid(Date, 7, 4) & Mid(Date, 4, 2) & Mid(Date, 1, 2) _
        & "-" & Replace(Time, ":", "") & ".arj"
    FileCopy "c:\TmpBackup\Backup.arj", x
    Kill "c:\TmpBackup\Backup.arj"
    Kill "c:\TmpBackup\Backup.sql"
    Kill "c:\TmpBackup\arj.exe"
    RmDir "c:\TmpBackup"
    MsgBox "Backup foi Efetuado com Êxito em:" & Chr(13) _
        & x, vbInformation, "Concluído!"
    On Error GoTo 0
End Sub

Private Sub Restaurar()
    If Dir(txtBackup.Text) = Empty Then GoTo Invalido
    If MsgBox(Space(27) & "ATENÇÃO!" & Chr(13) _
        & "Esta Operação Fará com que Todos os Registros" & Chr(13) _
        & "Atualmente Cadastrados no Banco de Dados Sejam" & Chr(13) _
        & "Apagados Para Que Possam Ser Gravados os Dados Restaurados." _
        & Chr(13) & "Este Procedimento Pode Acarretar em Perda de Dados." _
        & Chr(13) & "Use-o Somente em Casos Extremos." & Chr(13) _
        & "Se tiver certeza deste procedimento, certifique-se de " & Chr(13) _
        & "que todos os terminais estejam desconectados e clique em 'Sim'", vbCritical + vbYesNo _
        + vbDefaultButton2, "Confirmar Operação.") = vbNo Then Exit Sub
    Call Table.Query("show tables")
    While Not Table.Rs.EOF
        Data.Query ("DELETE FROM " _
            & Table.Rs.Fields.Item("Tables_in_pizza").Value)
        Table.Rs.MoveNext
    Wend
    
    MkDir "c:\TmpBackup"
    FileCopy txtBackup.Text, "c:\TmpBackup\Backup.arj"
    FileCopy App.Path & "\arj.exe", "c:\TmpBackup\arj.exe"
    ChDir "c:\TmpBackup"
    iTask = Shell("arj x -va Backup.arj ", vbHide)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    ret = WaitForSingleObject(pHandle, INFINITE)
    ret = CloseHandle(pHandle)
    Dim lineFile As String
    ' Disponibiliza o próximo número de arquivo disponível
    Arq = FreeFile
    ' Abre Arquivo Para Leitura
    Open "Backup.sql" For Input As #Arq
    ' Lê o Arquivo Linha por Linha até o Fim
    Do While Not EOF(1)
        ' Lê a Linha Corrente
        Line Input #Arq, lineFile
        Call Data.Query(lineFile)
    Loop
    Close #Arq
    ChDir App.Path
    Kill "c:\TmpBackup\Backup.arj"
    Kill "c:\TmpBackup\Backup.sql"
    Kill "c:\TmpBackup\arj.exe"
    RmDir "c:\TmpBackup"
    MsgBox "Restauração foi Efetuada com Êxito", vbInformation, "Concluído!"
    Exit Sub
Invalido:
    MsgBox "O Caminho de " & labBackup.Caption & " é Inválido", _
        vbInformation, labBackup.Caption & " Inválido"
End Sub

Private Sub HabCampos(Bool As Boolean)
    optCriar.Enabled = Bool
    optRestaurar.Enabled = Bool
    txtBackup.Enabled = Bool
    btnAbrir.Enabled = Bool
    btnBackup.Enabled = Bool
    btnFechar.Enabled = Bool
End Sub

