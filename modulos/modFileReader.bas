Attribute VB_Name = "modFileReader"
Option Explicit

Public Property Get Conf(LookedValue As String)
    ' Disponibiliza o próximo número de arquivo disponível
    Dim Arq As Integer
    Arq = FreeFile
    ' Abre Arquivo Para Leitura
    Open App.Path & "\config.ini" For Input As #Arq
    ' Lê o Arquivo Linha por Linha até o Fim
    Do While Not EOF(1)
        Dim lineFile As String
        ' Lê a Linha Corrente
        Line Input #Arq, lineFile
        ' Se a Linha Contiver a String Procurada
        If InStr(1, lineFile, LookedValue) <> 0 Then
            Dim posIgual As Integer
            ' Procura o Sinal de "="
            posIgual = InStr(1, lineFile, "=")
            ' Atribui à Variável "conf" o Valor da Linha Partindo
            ' do Sinal de "=" até o Fim da Linha
            Conf = Trim(Mid(lineFile, posIgual + 1))
            ' Fecha o Arquivo
            Close #Arq
            Exit Property
        End If
    Loop
    Close #Arq
End Property

Public Property Get Read(LookedValue As String, Separator As String, File As String)
    ' Disponibiliza o próximo número de arquivo disponível
    Dim Arq As Integer
    Arq = FreeFile
    ' Abre Arquivo Para Leitura
    Open File For Input As #Arq
    ' Lê o Arquivo Linha por Linha até o Fim
    Do While Not EOF(1)
        Dim lineFile As String
        ' Lê a Linha Corrente
        Line Input #Arq, lineFile
        ' Se a Linha Contiver a String Procurada
        If InStr(1, lineFile, LookedValue) <> 0 Then
            Dim posSeparator As Integer
            ' Procura o Separator
            posSeparator = InStr(1, lineFile, Separator)
            ' Atribui à Variável "Read" o Valor da Linha Partindo
            ' do Separator até o Fim da Linha
            Read = Trim(Mid(lineFile, posSeparator + 1))
            ' Fecha o Arquivo
            Close #Arq
            Exit Property
        End If
    Loop
    Close #Arq
End Property

