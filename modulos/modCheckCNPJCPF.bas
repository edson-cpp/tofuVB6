Attribute VB_Name = "modCheckCNPJCPF"
Option Explicit

Public Property Get CheckCNPJ(cnpj As String) As Boolean
   Dim VAR1, VAR2, VAR3, VAR4, VAR5
   If Len(cnpj) = 8 And Val(cnpj) > 0 Then
      VAR1 = 0
      VAR2 = 0
      VAR4 = 0
      For VAR3 = 1 To 7
         VAR1 = Val(Mid(cnpj, VAR3, 1))
         If (VAR1 Mod 2) <> 0 Then
            VAR1 = VAR1 * 2
         End If
         If VAR1 > 9 Then
            VAR2 = VAR2 + Int(VAR1 / 10) + (VAR1 Mod 10)
         Else
            VAR2 = VAR2 + VAR1
         End If
      Next VAR3
      VAR4 = IIf((VAR2 Mod 10) <> 0, 10 - (VAR2 Mod 10), 0)
      If VAR4 = Val(Mid(cnpj, 8, 1)) Then
         CheckCNPJ = True
      Else
         CheckCNPJ = False
      End If
   Else
      If Len(cnpj) = 14 And Val(cnpj) > 0 Then
         VAR1 = 0
         VAR3 = 0
         VAR4 = 0
         VAR5 = 0
         VAR2 = 5
         For VAR3 = 1 To 12
            VAR1 = VAR1 + (Val(Mid(cnpj, VAR3, 1)) * VAR2)
            VAR2 = IIf(VAR2 > 2, VAR2 - 1, 9)
         Next VAR3
         VAR1 = VAR1 Mod 11
         VAR4 = IIf(VAR1 > 1, 11 - VAR1, 0)
         VAR1 = 0
         VAR3 = 0
         VAR2 = 6
         For VAR3 = 1 To 13
            VAR1 = VAR1 + (Val(Mid(cnpj, VAR3, 1)) * VAR2)
            VAR2 = IIf(VAR2 > 2, VAR2 - 1, 9)
         Next VAR3
         VAR1 = VAR1 Mod 11
         VAR5 = IIf(VAR1 > 1, 11 - VAR1, 0)
         If (VAR4 = Val(Mid(cnpj, 13, 1)) And VAR5 = Val(Mid(cnpj, 14, 1))) Then
            CheckCNPJ = True
         Else
            CheckCNPJ = False
         End If
      Else
         CheckCNPJ = False
      End If
   End If
End Property

Public Property Get CheckCPF(CPF As String) As Boolean
   Dim EVAR1 As Integer
   Dim evar2 As Integer
   Dim F As Integer
   If Len(Trim(CPF)) <> 11 Then
      CheckCPF = False
      Exit Property
   End If
   EVAR1 = 0
   For F = 1 To 9
      EVAR1 = EVAR1 + Val(Mid(CPF, F, 1)) * (11 - F)
   Next F
   evar2 = 11 - (EVAR1 - (Int(EVAR1 / 11) * 11))
   If evar2 = 10 Or evar2 = 11 Then evar2 = 0
   If evar2 <> Val(Mid(CPF, 10, 1)) Then
      CheckCPF = False
      Exit Property
   End If
   EVAR1 = 0
   For F = 1 To 10
       EVAR1 = EVAR1 + Val(Mid(CPF, F, 1)) * (12 - F)
   Next F
   evar2 = 11 - (EVAR1 - (Int(EVAR1 / 11) * 11))
   If evar2 = 10 Or evar2 = 11 Then evar2 = 0
   If evar2 <> Val(Mid(CPF, 11, 1)) Then
      CheckCPF = False
      Exit Property
  End If
  CheckCPF = True
End Property

