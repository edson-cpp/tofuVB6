Attribute VB_Name = "modMoeda"
Option Explicit

Public Property Get FmtMoeda(FieldValue As String) As String
    If Val(Replace(FieldValue, ".", "")) > 999999.99 Then
        MsgBox "Valor máximo R$ 999.999,99"
        FmtMoeda = Empty
    Else
        FmtMoeda = Format(FieldValue, "R$ ###,##0.00")
    End If
End Property

