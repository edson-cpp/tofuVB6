Attribute VB_Name = "modGetKeyAscii"
Option Explicit

Public Property Get Minusculas(KeyAscii As Integer)
    Select Case KeyAscii
        Case 65 To 90, 192 To 196, 199 To 207, _
            210 To 214, 217 To 220
            Minusculas = KeyAscii + 32
        Case Else
            Minusculas = KeyAscii
    End Select
End Property

Public Property Get Maiusculas(KeyAscii As Integer)
    Select Case KeyAscii
        Case 97 To 122, 224 To 228, 231 To 239, _
            242 To 246, 249 To 252
            Maiusculas = KeyAscii - 32
        Case Else
            Maiusculas = KeyAscii
    End Select
End Property

Public Property Get Numeros(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, vbKeyBack, vbKeyReturn
            Numeros = KeyAscii
        Case Else
            Beep
            Numeros = Empty
    End Select
End Property

Public Property Get Letras(KeyAscii As Integer)
    Select Case KeyAscii
        Case 65 To 90, 192 To 196, 199 To 207, _
            210 To 214, 217 To 220, _
            97 To 122, 224 To 228, 231 To 239, _
            242 To 246, 249 To 252
            Letras = KeyAscii
        Case Else
            Beep
            Letras = Empty
    End Select
End Property

Public Property Get Moeda(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, 44, vbKeyBack, vbKeyReturn
            Moeda = KeyAscii
        Case 46
            Moeda = 44
        Case Else
            Beep
            Moeda = Empty
    End Select
End Property

Public Property Get LetrasENumeros(KeyAscii As Integer)
    Select Case KeyAscii
        Case 65 To 90, 97 To 122, 48 To 57, vbKeyBack, vbKeyReturn
            LetrasENumeros = KeyAscii
        Case Else
            Beep
            LetrasENumeros = Empty
    End Select
End Property
