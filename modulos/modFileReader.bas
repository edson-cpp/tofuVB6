Attribute VB_Name = "modFileReader"
Option Explicit

Public Property Get Conf(LookedValue As String)
    ' Disponibiliza o pr�ximo n�mero de arquivo dispon�vel
    Dim Arq As Integer
    Arq = FreeFile
    ' Abre Arquivo Para Leitura
    Open App.Path & "\config.ini" For Input As #Arq
    ' L� o Arquivo Linha por Linha at� o Fim
    Do While Not EOF(1)
        Dim lineFile As String
        ' L� a Linha Corrente
        Line Input #Arq, lineFile
        ' Se a Linha Contiver a String Procurada
        If InStr(1, lineFile, LookedValue) <> 0 Then
            Dim posIgual As Integer
            ' Procura o Sinal de "="
            posIgual = InStr(1, lineFile, "=")
            ' Atribui � Vari�vel "conf" o Valor da Linha Partindo
            ' do Sinal de "=" at� o Fim da Linha
            Conf = Trim(Mid(lineFile, posIgual + 1))
            ' Fecha o Arquivo
            Close #Arq
            Exit Property
        End If
    Loop
    Close #Arq
End Property

Public Property Get Read(LookedValue As String, Separator As String, File As String)
    ' Disponibiliza o pr�ximo n�mero de arquivo dispon�vel
    Dim Arq As Integer
    Arq = FreeFile
    ' Abre Arquivo Para Leitura
    Open File For Input As #Arq
    ' L� o Arquivo Linha por Linha at� o Fim
    Do While Not EOF(1)
        Dim lineFile As String
        ' L� a Linha Corrente
        Line Input #Arq, lineFile
        ' Se a Linha Contiver a String Procurada
        If InStr(1, lineFile, LookedValue) <> 0 Then
            Dim posSeparator As Integer
            ' Procura o Separator
            posSeparator = InStr(1, lineFile, Separator)
            ' Atribui � Vari�vel "Read" o Valor da Linha Partindo
            ' do Separator at� o Fim da Linha
            Read = Trim(Mid(lineFile, posSeparator + 1))
            ' Fecha o Arquivo
            Close #Arq
            Exit Property
        End If
    Loop
    Close #Arq
End Property

