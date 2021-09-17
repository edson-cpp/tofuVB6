Attribute VB_Name = "modFone"
Public Property Get fone(foneNumber As String)
    foneNumber = Replace(foneNumber, "(", "")
    foneNumber = Replace(foneNumber, ")", "")
    foneNumber = Replace(foneNumber, "-", "")
    Select Case Len(Trim(foneNumber))
        Case 7
            fone = Format(foneNumber, "(41)000-0000")
        Case 8
            fone = Format(foneNumber, "(41)0000-0000")
        Case 9
            fone = Format(foneNumber, "(00)000-0000")
        Case 10
            fone = Format(foneNumber, "(00)0000-0000")
        Case Else
            fone = Empty
    End Select
End Property
