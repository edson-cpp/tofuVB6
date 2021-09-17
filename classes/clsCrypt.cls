VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function CriptSenha(Psenha As String) As Variant
    Dim v_sqlerrm As String
    Dim SenhaCript As String
    Dim var1 As String
    Const MIN_ASC = 32
    Const MAX_ASC = 126
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    
    chave = 2001 ''qualquer nº para montar o algorítimo da criptografia
    Dim offset As Long
    Dim str_len As Integer
    Dim i As Integer
    Dim ch As Integer
        
    to_text = ""
    offset = NumericPassword(chave)
    Rnd -1
    Randomize offset
    str_len = Len(Psenha)
    For i = 1 To str_len
        ch = Asc(Mid$(Psenha, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch + offset) Mod NUM_ASC)
            ch = ch + MIN_ASC
            to_text = to_text & Chr$(ch)
        End If
    Next i
    
    CriptSenha = to_text
End Function

Public Function DeCriptSenha(Psenha As String) As Variant

Dim v_sqlerrm As String
Dim SenhaCript As String

Dim var1 As String

Const MIN_ASC = 32  ' Space.
Const MAX_ASC = 126 ' ~.
Const NUM_ASC = MAX_ASC - MIN_ASC + 1

chave = 2001 ''qualquer nº para montar o algorítimo da criptografia
Dim offset As Long
Dim str_len As Integer
Dim i As Integer
Dim ch As Integer
 
to_text = ""
offset = NumericPassword(chave)
Rnd -1
Randomize offset
str_len = Len(Psenha)
For i = 1 To str_len
    ch = Asc(Mid$(Psenha, i, 1))
    If ch >= MIN_ASC And ch <= MAX_ASC Then
        ch = ch - MIN_ASC
        offset = Int((NUM_ASC + 1) * Rnd)
        ch = ((ch - offset) Mod NUM_ASC)
        If ch < 0 Then ch = ch + NUM_ASC
        ch = ch + MIN_ASC
        to_text = to_text & Chr$(ch)
    End If
Next i

DeCriptSenha = to_text
    
End Function

Private Function NumericPassword(ByVal password As String) As Long
    Dim Value As Long
    Dim ch As Long
    Dim shift1 As Long
    Dim shift2 As Long
    Dim i As Integer
    Dim str_len As Integer

    str_len = Len(password)
    For i = 1 To str_len
        ' Adiciona a próxima letra
        ch = Asc(Mid$(password, i, 1))
        Value = Value Xor (ch * 2 ^ shift1)
        Value = Value Xor (ch * 2 ^ shift2)

        ' Change the shift offsets.
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    NumericPassword = Value
End Function

