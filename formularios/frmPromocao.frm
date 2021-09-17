VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPromocao 
   Caption         =   "Cadastro de Promoções"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   Icon            =   "frmPromocao.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   323
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   594
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraProcura 
      Caption         =   "Localização"
      Height          =   1215
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   8655
      Begin VB.CommandButton btnSair 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "Sai&r"
         Height          =   375
         Left            =   7200
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton btnSalvar 
         Appearance      =   0  'Flat
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7200
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1560
         TabIndex        =   19
         Top             =   720
         Width           =   3735
      End
      Begin VB.ComboBox cbbProcura 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label labProcura 
         AutoSize        =   -1  'True
         Caption         =   "Campo de Procura:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   413
         Width           =   1365
      End
      Begin VB.Label labValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor Procurado:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   773
         Width           =   1185
      End
   End
   Begin VB.TextBox txtSab 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7200
      TabIndex        =   14
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtSex 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7200
      TabIndex        =   13
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtQui 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7200
      TabIndex        =   12
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtQua 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7200
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtTer 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7200
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtSeg 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7200
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtDom 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7200
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid mfgProdutos 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5741
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   -2147483624
      ForeColorSel    =   -2147483630
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Label labInfo2 
      AutoSize        =   -1  'True
      Caption         =   "Deixe em branco para Valor padrão."
      Height          =   195
      Left            =   6240
      TabIndex        =   16
      Top             =   1800
      Width           =   2550
   End
   Begin VB.Label labInfo 
      AutoSize        =   -1  'True
      Caption         =   "Informe o Valor da Promoção do dia."
      Height          =   195
      Left            =   6240
      TabIndex        =   15
      Top             =   1560
      Width           =   2580
   End
   Begin VB.Label labSab 
      AutoSize        =   -1  'True
      Caption         =   "Sábado"
      Height          =   195
      Left            =   6240
      TabIndex        =   7
      Top             =   4380
      Width           =   555
   End
   Begin VB.Label labSex 
      AutoSize        =   -1  'True
      Caption         =   "Sexta"
      Height          =   195
      Left            =   6240
      TabIndex        =   6
      Top             =   4020
      Width           =   405
   End
   Begin VB.Label labQui 
      AutoSize        =   -1  'True
      Caption         =   "Quinta"
      Height          =   195
      Left            =   6240
      TabIndex        =   5
      Top             =   3660
      Width           =   465
   End
   Begin VB.Label labQua 
      AutoSize        =   -1  'True
      Caption         =   "Quarta"
      Height          =   195
      Left            =   6240
      TabIndex        =   4
      Top             =   3300
      Width           =   480
   End
   Begin VB.Label labTer 
      AutoSize        =   -1  'True
      Caption         =   "Terça"
      Height          =   195
      Left            =   6240
      TabIndex        =   3
      Top             =   2940
      Width           =   420
   End
   Begin VB.Label labSeg 
      AutoSize        =   -1  'True
      Caption         =   "Segunda"
      Height          =   195
      Left            =   6240
      TabIndex        =   2
      Top             =   2580
      Width           =   645
   End
   Begin VB.Label labDom 
      AutoSize        =   -1  'True
      Caption         =   "Domingo"
      Height          =   195
      Left            =   6240
      TabIndex        =   1
      Top             =   2220
      Width           =   630
   End
End
Attribute VB_Name = "frmPromocao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dtCatcher As clsMyConnect
Dim Conn As clsMyConnect
Dim Aux As clsMyConnect
Dim Campo As String

Private Sub btnSair_Click()
    If btnSair.Caption = "Sai&r" Then
        Unload Me
    Else
        mfgProdutos_EnterCell
    End If
End Sub

Private Sub btnSalvar_Click()
    If txtDom.DataChanged Then
        Call GravaPromo(txtDom.Text, txtDom.Tag, "1")
    End If
    If txtSeg.DataChanged Then
        Call GravaPromo(txtSeg.Text, txtSeg.Tag, "2")
    End If
    If txtTer.DataChanged Then
        Call GravaPromo(txtTer.Text, txtTer.Tag, "3")
    End If
    If txtQua.DataChanged Then
        Call GravaPromo(txtQua.Text, txtQua.Tag, "4")
    End If
    If txtQui.DataChanged Then
        Call GravaPromo(txtQui.Text, txtQui.Tag, "5")
    End If
    If txtSex.DataChanged Then
        Call GravaPromo(txtSex.Text, txtSex.Tag, "6")
    End If
    If txtSab.DataChanged Then
        Call GravaPromo(txtSab.Text, txtSab.Tag, "7")
    End If
    Call Cancelar
    MsgBox "Registro Salvo com Êxito", vbInformation, "Concluído!"
End Sub

Private Sub cbbProcura_Click()
    Select Case cbbProcura.ListIndex
        Case 0
            Campo = "id"
        Case 1
            Campo = "descricao"
    End Select
End Sub

Private Sub Form_Load()
    Set Conn = New clsMyConnect
    Set dtCatcher = New clsMyConnect
    Set Aux = New clsMyConnect
    Conn.Connect
    dtCatcher.Connect
    Aux.Connect
    Call Conn.Query("SELECT id, descricao, preco, tipo, divpessoa FROM produto")
    With cbbProcura
        .AddItem ("Código")
        .AddItem ("Descrição")
        .Text = "Descrição"
    End With
    With mfgProdutos
        .Cols = 4
        .Rows = 1
        .Clear
        .ColWidth(0) = 800
        .ColWidth(1) = 3250
        .ColWidth(2) = 1000
        .ColWidth(3) = 800
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
        mfgProdutos.Rows = mfgProdutos.Rows + 1
        mfgProdutos.TextMatrix(Rows, 0) = Conn.Rs.Fields.Item("id").Value
        fieldWidth = TextWidth(Conn.Rs.Fields.Item("id").Value)
        If colsWidth(0) < fieldWidth Then
            colsWidth(0) = fieldWidth
        End If
        mfgProdutos.TextMatrix(Rows, 1) = Conn.Rs.Fields.Item("descricao").Value
        fieldWidth = TextWidth(Conn.Rs.Fields.Item("descricao").Value)
        If colsWidth(1) < fieldWidth Then
            colsWidth(1) = fieldWidth
        End If
        mfgProdutos.TextMatrix(Rows, 2) = modMoeda.FmtMoeda(Conn.Rs.Fields.Item("preco").Value)
        fieldWidth = TextWidth(modMoeda.FmtMoeda(Conn.Rs.Fields.Item("preco").Value))
        If colsWidth(2) < fieldWidth Then
            colsWidth(2) = fieldWidth
        End If
        Dim x As Integer
        Dim y As String
        x = Conn.Rs.Fields.Item("tipo").Value
        y = IIf(x = 0, "Comida", "Bebida")
        mfgProdutos.TextMatrix(Rows, 3) = y
        fieldWidth = TextWidth(y)
        If colsWidth(3) < fieldWidth Then
            colsWidth(3) = fieldWidth
        End If
        Conn.Rs.MoveNext
        Rows = Rows + 1
    Loop
    mfgProdutos_EnterCell
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dtCatcher.Disconnect
    Conn.Disconnect
    Aux.Disconnect
End Sub

Private Sub mfgProdutos_EnterCell()
    mfgProdutos.Tag = "EnterCell"
    Conn.Rs.AbsolutePosition = mfgProdutos.Row
    mfgProdutos.CellBackColor = &H80000018 'Amarelo
    txtDom.Text = Empty
    txtSeg.Text = Empty
    txtTer.Text = Empty
    txtQua.Text = Empty
    txtQui.Text = Empty
    txtSex.Text = Empty
    txtSab.Text = Empty
    Call dtCatcher.Query("SELECT id, pid, dia, preco FROM promocao WHERE pid = " & Conn.Rs.Fields.Item("id").Value & " ORDER BY dia")
    While Not dtCatcher.Rs.EOF
        Select Case dtCatcher.Rs.Fields.Item("dia").Value
            Case 1
                txtDom.Text = modMoeda.FmtMoeda( _
                    dtCatcher.Rs.Fields.Item("preco").Value)
            Case 2
                txtSeg.Text = modMoeda.FmtMoeda( _
                    dtCatcher.Rs.Fields.Item("preco").Value)
            Case 3
                txtTer.Text = modMoeda.FmtMoeda( _
                    dtCatcher.Rs.Fields.Item("preco").Value)
            Case 4
                txtQua.Text = modMoeda.FmtMoeda( _
                    dtCatcher.Rs.Fields.Item("preco").Value)
            Case 5
                txtQui.Text = modMoeda.FmtMoeda( _
                    dtCatcher.Rs.Fields.Item("preco").Value)
            Case 6
                txtSex.Text = modMoeda.FmtMoeda( _
                    dtCatcher.Rs.Fields.Item("preco").Value)
            Case 7
                txtSab.Text = modMoeda.FmtMoeda( _
                    dtCatcher.Rs.Fields.Item("preco").Value)
        End Select
        dtCatcher.Rs.MoveNext
    Wend
    Call Cancelar
    mfgProdutos.Tag = ""
End Sub

Private Sub mfgProdutos_LeaveCell()
    mfgProdutos.CellBackColor = &H80000005 'Branco
    If Not btnSair.Caption = "&Cancelar" Then Exit Sub
    If MsgBox("Os preços dos itens foram alterados." & _
        Chr(13) & "Deseja salvar as alterações?", vbExclamation _
        + vbYesNo + vbDefaultButton1, "Salvar Alterações?") = _
        vbNo Then Exit Sub
        btnSalvar_Click
End Sub

Private Sub txtDom_Change()
    If mfgProdutos.Tag = "EnterCell" Then Exit Sub
    btnSair.Caption = "&Cancelar"
    btnSalvar.Enabled = True
    If txtDom.Tag = txtDom.Text Then
        txtDom.DataChanged = False
        Call Alterado
    End If
End Sub

Private Sub txtDom_GotFocus()
    txtDom.SelStart = 0
    txtDom.SelLength = 13
    txtDom.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtDom_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Moeda(KeyAscii)
End Sub

Private Sub txtDom_LostFocus()
    txtDom.BackColor = &H80000005 'Branco
    txtDom.Text = modMoeda.FmtMoeda(txtDom.Text)
End Sub

Private Sub txtQua_Change()
    If mfgProdutos.Tag = "EnterCell" Then Exit Sub
    btnSair.Caption = "&Cancelar"
    btnSalvar.Enabled = True
    If txtQua.Tag = txtQua.Text Then
        txtQua.DataChanged = False
        Call Alterado
    End If
End Sub

Private Sub txtQua_GotFocus()
    txtQua.SelStart = 0
    txtQua.SelLength = 13
    txtQua.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtQua_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Moeda(KeyAscii)
End Sub

Private Sub txtQua_LostFocus()
    txtQua.BackColor = &H80000005 'Branco
    txtQua.Text = modMoeda.FmtMoeda(txtQua.Text)
End Sub

Private Sub txtQui_Change()
    If mfgProdutos.Tag = "EnterCell" Then Exit Sub
    btnSair.Caption = "&Cancelar"
    btnSalvar.Enabled = True
    If txtQui.Tag = txtQui.Text Then
        txtQui.DataChanged = False
        Call Alterado
    End If
End Sub

Private Sub txtQui_GotFocus()
    txtQui.SelStart = 0
    txtQui.SelLength = 13
    txtQui.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtQui_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Moeda(KeyAscii)
End Sub

Private Sub txtQui_LostFocus()
    txtQui.BackColor = &H80000005 'Branco
    txtQui.Text = modMoeda.FmtMoeda(txtQui.Text)
End Sub

Private Sub txtSab_Change()
    If mfgProdutos.Tag = "EnterCell" Then Exit Sub
    btnSair.Caption = "&Cancelar"
    btnSalvar.Enabled = True
    If txtSab.Tag = txtSab.Text Then
        txtSab.DataChanged = False
        Call Alterado
    End If
End Sub

Private Sub txtSab_GotFocus()
    txtSab.SelStart = 0
    txtSab.SelLength = 13
    txtSab.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtSab_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Moeda(KeyAscii)
End Sub

Private Sub txtSab_LostFocus()
    txtSab.BackColor = &H80000005 'Branco
    txtSab.Text = modMoeda.FmtMoeda(txtSab.Text)
End Sub

Private Sub txtSeg_Change()
    If mfgProdutos.Tag = "EnterCell" Then Exit Sub
    btnSair.Caption = "&Cancelar"
    btnSalvar.Enabled = True
    If txtSeg.Tag = txtSeg.Text Then
        txtSeg.DataChanged = False
        Call Alterado
    End If
End Sub

Private Sub txtSeg_GotFocus()
    txtSeg.SelStart = 0
    txtSeg.SelLength = 13
    txtSeg.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtSeg_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Moeda(KeyAscii)
End Sub

Private Sub txtSeg_LostFocus()
    txtSeg.BackColor = &H80000005 'Branco
    txtSeg.Text = modMoeda.FmtMoeda(txtSeg.Text)
End Sub

Private Sub txtSex_Change()
    If mfgProdutos.Tag = "EnterCell" Then Exit Sub
    btnSair.Caption = "&Cancelar"
    btnSalvar.Enabled = True
    If txtSex.Tag = txtSex.Text Then
        txtSex.DataChanged = False
        Call Alterado
    End If
End Sub

Private Sub txtSex_GotFocus()
    txtSex.SelStart = 0
    txtSex.SelLength = 13
    txtSex.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtSex_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Moeda(KeyAscii)
End Sub

Private Sub txtSex_LostFocus()
    txtSex.BackColor = &H80000005 'Branco
    txtSex.Text = modMoeda.FmtMoeda(txtSex.Text)
End Sub

Private Sub txtTer_Change()
    If mfgProdutos.Tag = "EnterCell" Then Exit Sub
    btnSair.Caption = "&Cancelar"
    btnSalvar.Enabled = True
    If txtTer.Tag = txtTer.Text Then
        txtTer.DataChanged = False
        Call Alterado
    End If
End Sub

Private Sub txtTer_GotFocus()
    txtTer.SelStart = 0
    txtTer.SelLength = 13
    txtTer.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtTer_KeyPress(KeyAscii As Integer)
    KeyAscii = modGetKeyAscii.Moeda(KeyAscii)
End Sub

Private Sub txtTer_LostFocus()
    txtTer.BackColor = &H80000005 'Branco
    txtTer.Text = modMoeda.FmtMoeda(txtTer.Text)
End Sub

Private Sub txtValor_Change()
    Dim I As Integer
    For I = 0 To mfgProdutos.Cols - 1
        mfgProdutos.ColSel = I
        mfgProdutos.CellBackColor = &H80000005 'Branco
    Next
    Conn.Rs.MoveFirst
    On Error GoTo Erro
    Conn.Rs.Find Campo & " = '" & dtCatcher.GetValue("SELECT " _
        & Campo & " FROM produto WHERE " & Campo & " LIKE '" _
        & Trim(txtValor.Text) & "%'") & "'"
    If Conn.Rs.EOF Then
        Conn.Rs.MoveFirst
    End If
    mfgProdutos.Row = Conn.Rs.AbsolutePosition
    For I = 0 To mfgProdutos.Cols - 1
        mfgProdutos.ColSel = I
        mfgProdutos.CellBackColor = &H80000018 'Amarelo
    Next
    mfgProdutos.Refresh
    On Error GoTo 0
    Exit Sub
Erro:
    On Error GoTo 0
    ' Registro não localizado
    If Err.Number = -2147352571 Then
        Conn.Rs.Move -1, 1
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
End Sub

Private Sub txtValor_GotFocus()
    txtValor.BackColor = &H80000018 'Amarelo
End Sub

Private Sub txtValor_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then 'Seta Abaixo
        mfgProdutos.SetFocus
    End If
End Sub

Private Sub txtValor_LostFocus()
    txtValor.BackColor = &H80000005 'Branco
End Sub

Private Sub cbbProcura_GotFocus()
    cbbProcura.BackColor = &H80000018 'Amarelo
End Sub

Private Sub cbbProcura_LostFocus()
    cbbProcura.BackColor = &H80000005 'Branco
End Sub

Private Sub Cancelar()
    txtDom.Tag = txtDom.Text
    txtSeg.Tag = txtSeg.Text
    txtTer.Tag = txtTer.Text
    txtQua.Tag = txtQua.Text
    txtQui.Tag = txtQui.Text
    txtSex.Tag = txtSex.Text
    txtSab.Tag = txtSab.Text
    txtDom.DataChanged = False
    txtSeg.DataChanged = False
    txtTer.DataChanged = False
    txtQua.DataChanged = False
    txtQui.DataChanged = False
    txtSex.DataChanged = False
    txtSab.DataChanged = False
    btnSair.Caption = "Sai&r"
    btnSalvar.Enabled = False
End Sub

Private Sub GravaPromo(FieldText As String, FieldTag As String, FieldNumber As String)
    If FieldText = Empty And Not FieldTag = Empty Then
        Aux.Query ("DELETE FROM promocao WHERE pid = " _
            & Conn.Rs.Fields.Item("id").Value & " AND dia = " & FieldNumber)
    ElseIf Not FieldText = Empty And Not FieldTag = Empty Then
        Aux.Query ("UPDATE promocao SET preco = " _
            & Replace(CDbl(FieldText), ",", ".") & " WHERE pid = " _
            & Conn.Rs.Fields.Item("id").Value & " AND dia = " & FieldNumber)
    ElseIf Not FieldText = Empty And FieldTag = Empty Then
        Aux.Query ("INSERT INTO promocao ( id, pid, dia, preco " _
            & ") VALUES ( null, " & Conn.Rs.Fields.Item("id").Value _
            & ", " & FieldNumber & ", " _
            & Replace(CSng(FieldText), ",", ".")) & " )"
    End If
End Sub

Private Sub Alterado()
    If txtDom.DataChanged Then Exit Sub
    If txtSeg.DataChanged Then Exit Sub
    If txtTer.DataChanged Then Exit Sub
    If txtQua.DataChanged Then Exit Sub
    If txtQui.DataChanged Then Exit Sub
    If txtSex.DataChanged Then Exit Sub
    If txtSab.DataChanged Then Exit Sub
    btnSair.Caption = "Sai&r"
    btnSalvar.Enabled = False
End Sub
