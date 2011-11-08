Attribute VB_Name = "Validate_lib"
'-------------------------------------------------------------------------
' Validate_lib.bas
'-------------------------------------------------------------------------
' Descricao   : Biblioteca utilizada no objeto Validate
'               Exposto aqui para evitar Referências cíclicas
' Criaçao     : 16/01/2000 19:03
' Local       : Brasilia/DF
' Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
'               Ridai Govinda Pombo <ridai@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 1997, 1998, 1999, 2000 by Zevallos(r) Tecnologia em Informacao
'-------------------------------------------------------------------------

Rem -------------------------------------------------------------------------
Rem Testa se e um CPF
Rem -------------------------------------------------------------------------
Public Function HTIsCPF(ByVal strCPF As String) As Boolean

  Dim intSoma As Integer, i As Integer, strNumber As String
  Dim IntDigit2 As Integer, intMultiply As Integer
  Dim IntDigit1 As Integer, intNet As Integer
  'Dim sobjString As HiperTools30.String
  
  If strCPF > "" Then
    intSoma = 0
    intMultiply = 10
    
    strCPF = HTNormalizeNumber(strCPF)
    strCPF = HTLeadingZeroes(strCPF, 11)
    
    For i = 1 To 9
        strNumber = Int(Mid(strCPF, i, 1))
        intSoma = intSoma + (strNumber * intMultiply)
        intMultiply = intMultiply - 1
    
    Next
    
    intNet = intSoma Mod 11
    
    If (intNet = 0) Or (intNet = 1) Then
       IntDigit1 = 0
    
    Else
       IntDigit1 = CInt(11 - intNet)
    
    End If
    
    If IntDigit1 = Int(Mid(strCPF, 10, 1)) Then
       intSoma = 0
       intMultiply = 11
    
       For i = 1 To 10
           strNumber = Int(Mid(strCPF, i, 1))
           intSoma = intSoma + (strNumber * intMultiply)
           intMultiply = intMultiply - 1
    
       Next
    
       intNet = intSoma Mod 11
    
       If (intNet = 0) Or (intNet = 1) Then
          IntDigit2 = 0
       Else
          IntDigit2 = 11 - intNet
       End If
    
       If IntDigit2 = Int(Mid(strCPF, 11, 1)) Then
         HTIsCPF = True
 
       Else
         HTIsCPF = False
 
       End If
    Else
      HTIsCPF = False

    End If
    
    strCPF = HTFormatCPF(strCPF)

  Else
    HTIsCPF = True

  End If
  
  Set sobjString = Nothing
End Function


Rem -------------------------------------------------------------------------
Rem Testa se e um CGC
Rem -------------------------------------------------------------------------
Public Function HTIsCGC(ByVal strCGC As String) As Boolean
  Dim intSoma As Integer, i As Integer, strNumber As String
  Dim IntDigit1 As Integer, IntDigit2 As Integer, intMultiply As Integer, intNet As Integer
  
  If strCGC > "" Then
    intMultiply = 5
    intSoma = 0
  
    strCGC = HTNormalizeNumber(strCGC)
    strCGC = HTLeadingZeroes(strCGC, 14)
    For i = 1 To 12
        strNumber = Int(Mid(strCGC, i, 1))
        intSoma = intSoma + (intMultiply * strNumber)
        intMultiply = intMultiply - 1
        If intMultiply < 2 Then
           intMultiply = 9
        End If
    Next
    intNet = intSoma Mod 11
    If (intNet = 0) Or (intNet = 1) Then
       IntDigit1 = 0
    Else
       IntDigit1 = 11 - intNet
    End If
    If IntDigit1 = Int(Mid(strCGC, 13, 1)) Then
       intMultiply = 6
       intSoma = 0
  
       For i = 1 To 13
           strNumber = Int(Mid(strCGC, i, 1))
           intSoma = intSoma + (intMultiply * strNumber)
           intMultiply = intMultiply - 1
  
           If intMultiply < 2 Then
              intMultiply = 9
           End If
       Next
  
       intNet = intSoma Mod 11
  
       If (intNet = 0) Or (intNet = 1) Then
          IntDigit2 = 0
       Else
          IntDigit2 = 11 - intNet
       End If
  
       If IntDigit2 = Int(Mid(strCGC, 14, 1)) Then
          HTIsCGC = True
  
       Else
          HTIsCGC = False
       End If
  
     Else
       HTIsCGC = False
     End If
  
  Else
    HTIsCGC = True
  
  End If
  
  Set sobjString = Nothing

End Function
Rem ---------------------------------------------------------------------
Rem Fim do /HiperTools/CPF.inc

Public Function HTIsCPF_CGC(ByVal strSomething As String, Optional ByVal bytWhichFirst As Byte = 0) As Boolean
  
  HTIsCPF_CGC = False
  
  Select Case bytWhichFirst
  Case vlCp_CPFFirst
    
    If HTIsCPF(strSomething) Then
      HTIsCPF_CGC = True
    Else
      If HTIsCGC(strSomething) Then HTIsCPF_CGC = True
    End If
  
  Case vlCp_CGCFirst
    
    If HTIsCGC(strSomething) Then
      HTIsCPF_CGC = True
    Else
      If HTIsCPF(strSomething) Then HTIsCPF_CGC = True
    End If
  
  End Select

End Function

