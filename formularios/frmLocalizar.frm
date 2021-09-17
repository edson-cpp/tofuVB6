VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLocalizar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Localizar"
   ClientHeight    =   4365
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7815
   Icon            =   "frmLocalizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   291
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   521
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid mfgLocalizar 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4895
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   -2147483648
      BackColorSel    =   -2147483624
      ForeColorSel    =   -2147483630
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Frame fraProcura 
      Caption         =   "Localização"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7575
      Begin VB.ComboBox cbbProcura 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1560
         TabIndex        =   0
         Top             =   720
         Width           =   3735
      End
      Begin VB.CommandButton btnOk 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   6240
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton btnCancel 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   6240
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label labValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor Procurado:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   773
         Width           =   1185
      End
      Begin VB.Label labProcura 
         AutoSize        =   -1  'True
         Caption         =   "Campo de Procura:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   413
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frmLocalizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dtCatcher As clsMyConnect 'Declara classe para realizar consultas
Public Campo As String 'Campo Procurado
Public tabela As String 'Nome da tabela donde virão os dados
Public frmPai As Form 'Formulário que originou a chamada

Private Sub btnCancel_Click()
    frmPai.Localizar = False
    Unload Me
End Sub

Private Sub btnOk_Click()
    frmPai.Conn.Rs.AbsolutePosition = mfgLocalizar.Row
    frmPai.Localizar = True
    Unload Me
End Sub

Private Sub cbbProcura_LostFocus()
    Call frmPai.Campo(cbbProcura.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dtCatcher.Disconnect
End Sub

Private Sub mfgLocalizar_DblClick()
    btnOk_Click
End Sub

Private Sub Form_Load()
    Dim I As Integer
    Set dtCatcher = New clsMyConnect
    Set frmPai = frmPrin.frmPai
    Call frmPai.Load
    Call frmPai.Campo(cbbProcura.Text)
    dtCatcher.Connect
    frmPai.Conn.Rs.MoveFirst
    For I = 0 To mfgLocalizar.Cols - 1
        mfgLocalizar.ColSel = I
        mfgLocalizar.CellBackColor = &H80000018 'Amarelo
    Next
End Sub

Private Sub mfgLocalizar_EnterCell()
    mfgLocalizar.CellBackColor = &H80000018 'Amarelo
End Sub

Private Sub mfgLocalizar_LeaveCell()
    mfgLocalizar.CellBackColor = &H80000005 'Branco
End Sub

Private Sub txtValor_Change()
    Dim I As Integer
    For I = 0 To mfgLocalizar.Cols - 1
        mfgLocalizar.ColSel = I
        mfgLocalizar.CellBackColor = &H80000005
    Next
    frmPai.Conn.Rs.MoveFirst
    On Error GoTo erro
    frmPai.Conn.Rs.Find Campo & " = '" & dtCatcher.GetValue("SELECT " & Campo & " FROM " & tabela & " WHERE " & Campo & " LIKE '" & Trim(txtValor.Text) & "%'") & "'"
    GoTo ok
erro:
        On Error GoTo 0
        ' Registro não localizado
        If Err.Number = -2147352571 Then
            frmPai.Conn.Rs.Move -1, 1
            Exit Sub
        Else
            ' Erro desconhecido
            MsgBox "Erro em tempo de execução." _
                & Chr(13) & Chr(13) _
                & "Código:      " & Err.Number & Chr(13) _
                & "Descrição:  " & Err.Description _
                & Chr(13) & Chr(13) _
                & "Por favor entre em contato com o suporte técnico."
        End If
ok:
    If IsEmpty(dtCatcher.GetValue("SELECT " & Campo & " FROM " & tabela & " WHERE " & Campo & " LIKE '" & Trim(txtValor.Text) & "%'")) Then
        frmPai.Conn.Rs.MoveFirst
    End If
    mfgLocalizar.Row = frmPai.Conn.Rs.AbsolutePosition
    For I = 0 To mfgLocalizar.Cols - 1
        mfgLocalizar.ColSel = I
        mfgLocalizar.CellBackColor = &H80000018
    Next
    mfgLocalizar.Refresh
    On Error GoTo 0
End Sub

Private Sub txtValor_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then 'Seta Abaixo
        mfgLocalizar.SetFocus
    End If
End Sub
