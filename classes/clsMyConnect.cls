VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Conex As ADODB.Connection
Attribute Conex.VB_VarHelpID = -1
Public WithEvents Rs As ADODB.Recordset
Attribute Rs.VB_VarHelpID = -1
Public NumErro As Long
Public DescErro As String
Public SrcErro As String

Public Sub Connect()
    Dim sqlConn As String
    Dim myHost As String
    Dim myUid As String
    Dim myPwd As String
    Dim myBase As String
    Dim myDriver As String
    Dim Cript As clsCrypt
    Set Conex = New ADODB.Connection
    Set Cript = New clsCrypt
    NumErro = Empty
    DescErro = Empty
    SrcErro = Empty
    On Error GoTo ConFail
    myHost = modFileReader.Conf("tofu_mysql_host_connector")
    myUid = modFileReader.Conf("tofu_mysql_uid_connector")
    myPwd = Cript.DeCriptSenha(modFileReader.Conf("tofu_mysql_pwd_connector"))
    myBase = modFileReader.Conf("tofu_mysql_base_connector")
    myDriver = modFileReader.Conf("tofu_mysql_driver_connector")
    Conex.ConnectionTimeout = 60
    Conex.CommandTimeout = 400
    sqlConn = "DRIVER=" & myDriver _
        & ";SERVER=" & myHost _
        & ";UID=" & myUid _
        & ";PWD=" & myPwd _
        & ";DATABASE=" & myBase
    Conex.CursorLocation = adUseClient
    Conex.Open sqlConn
    On Error GoTo 0
    NumErro = Empty
    Exit Sub
ConFail:
    NumErro = Err.Number
    DescErro = Err.Description
    SrcErro = Err.Source
    On Error GoTo 0
End Sub

Public Sub Disconnect()
    On Error GoTo DiscFail
    Conex.Close
    On Error GoTo 0
    NumErro = Empty
    Exit Sub
DiscFail:
    NumErro = Err.Number
    DescErro = Err.Description
    SrcErro = Err.Source
    On Error GoTo 0
End Sub

Public Property Get GetValue(sql As String)
    Call Query(sql)
    On Error GoTo ValueFail
    If Rs.RecordCount = 0 Then
        GetValue = Empty
    Else
        If IsNull(Rs.Fields.Item(0).Value) Then
            GetValue = Empty
        Else
            GetValue = Rs.Fields.Item(0).Value
        End If
    End If
    On Error GoTo 0
    NumErro = Empty
    Exit Sub
ValueFail:
    NumErro = Err.Number
    DescErro = Err.Description
    SrcErro = Err.Source
    On Error GoTo 0
End Property

Public Property Get GetBlobText(sql As String)
    Call Query(sql)
    Dim myStream As ADODB.Stream
    Set myStream = New ADODB.Stream
    On Error GoTo BlobFail
    If Rs.RecordCount = 0 Then
        GetBlobText = Empty
    Else
        myStream.Open
        myStream.WriteText Rs.Fields.Item(0).Value
        GetBlobText = myStream.ReadText
    End If
    myStream.Close
    On Error GoTo 0
    NumErro = Empty
    Exit Sub
BlobFail:
    NumErro = Err.Number
    DescErro = Err.Description
    SrcErro = Err.Source
    On Error GoTo 0
End Property

Public Sub Query(sql As String)
    If NumErro = -2147217871 Then 'Falha de chave estrangeira
        NumErro = Empty
        DescErro = Empty
        SrcErro = Empty
    ElseIf Not NumErro = 0 Then
        MsgBox "Conexão com os Dados Está com Falha" _
            & Chr(13) & "Por Favor Reinicie o Aplicativo"
        Exit Sub
    End If
    sql = Replace(sql, "\", "\\")
    sql = Replace(sql, "\\'", "\'")
    On Error GoTo QueryFail
    Set Rs = New ADODB.Recordset
    Set Rs = CreateObject("ADODB.Recordset")
    Rs.CursorType = adOpenStatic
    Rs.CursorLocation = adUseClient
    Rs.LockType = adLockPessimistic
    Rs.Open sql, Conex
    On Error GoTo 0
    NumErro = Empty
    Exit Sub
QueryFail:
    NumErro = Err.Number
    DescErro = Err.Description
    SrcErro = Err.Source
    On Error GoTo 0
End Sub

