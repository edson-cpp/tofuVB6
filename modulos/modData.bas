Attribute VB_Name = "modData"
Option Explicit
Public Property Get FmtData(meb As String) As String
    If meb = Empty Then Exit Property
    Dim Data As String
    Dim dia As Integer
    Dim mes As Integer
    Dim ano As Integer
    Dim fevereiro As Integer
    Select Case Len(meb)
        Case 2
            meb = meb & Month(Date) & Year(Date)
        Case 4
            meb = meb & Year(Date)
        Case 6
            ano = Val(Mid(Year(Date), 3, 2))
            meb = Mid(meb, 1, 4) & _
                IIf(Mid(meb, 5, 2) <= ano _
                + 10, "20", "19") _
                & Mid(meb, 5, 2)
    End Select
    FmtData = Mid(meb, 1, 2) & _
        "/" & Mid(meb, 3, 2) & _
        "/" & Mid(meb, 5, 4)
    Data = FmtData
    dia = Val(Mid(Data, 1, 2))
    mes = Val(Mid(Data, 4, 2))
    ano = Val(Mid(Data, 7, 4))
    
    'Verificando os meses que podem ter até o dia 31
    Select Case mes
        Case 1, 3, 5, 7, 8, 10, 12
            If (dia < 1) Or (dia > 31) Then
                MsgBox ("Data Inválida!")
                FmtData = ""
                Exit Property
            End If
    
    'Verificando o mes de fevereiro
        Case 2
            If (dia >= 30) Then
                MsgBox ("Data Inválida!")
                FmtData = ""
                Exit Property
            End If
            fevereiro = ano Mod 4
            If (fevereiro <> 0) And (dia = 29) Then
                MsgBox ("Data Inválida!")
                FmtData = ""
                Exit Property
            End If
    
    'Verificar os meses que não podem ter dia até 31 e sim até 30
        Case 4, 6, 9, 11
            If (dia < 1) Or (dia > 30) Then
                MsgBox ("Data Inválida!")
                FmtData = ""
                Exit Property
            End If
    
    'Verificar os meses 1 A 12
        Case Else
            MsgBox ("Data Inválida!")
            FmtData = ""
            Exit Property
    End Select
End Property

Public Property Get FmtHora(meb As String) As String
    FmtHora = Empty
    If Val(Mid(meb, 1)) = 0 Then
        FmtHora = "0000"
        Exit Property
    End If
    Select Case Len(meb)
        Case 1
            FmtHora = "0" & meb & "00"
        Case 2
            If Val(meb) > 24 Then
                MsgBox "Hora Inválida"
                Exit Property
            ElseIf Val(meb) = 24 Then
                FmtHora = "0000"
            Else
                FmtHora = Format(meb, "00") & "00"
            End If
        Case 3
            If Val(meb) > 249 Then
                MsgBox "Hora Inválida"
                Exit Property
            ElseIf Val(Mid(meb, 1, 2)) = 24 Then
                FmtHora = "000" & Mid(meb, 3, 1)
            Else
                FmtHora = Mid(meb, 1, 2) & "0" & Mid(meb, 3, 1)
            End If
        Case 4
            If Val(meb) > 2459 Then
                MsgBox "Hora Inválida"
                Exit Property
            ElseIf Val(Mid(meb, 3, 2)) > 59 Then
                MsgBox "Hora Inválida"
                Exit Property
            ElseIf Val(Mid(meb, 1, 2)) = 24 Then
                FmtHora = "00" & Mid(meb, 3, 2)
            Else
                FmtHora = Format(meb, "0000")
            End If
    End Select
End Property

Public Property Get MyToVb(Data As String)
    MyToVb = Mid(Data, 9, 2) & "/" & Mid(Data, 6, 2) & "/" & Mid(Data, 1, 4)
End Property

Public Property Get VbToMy(Data As Date)
    VbToMy = Mid(Data, 7, 4) & "-" & Mid(Data, 4, 2) & "-" & Mid(Data, 1, 2)
End Property

