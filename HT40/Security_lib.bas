Attribute VB_Name = "Security_lib"
'-------------------------------------------------------------------------
' Security_lib.bas
'-------------------------------------------------------------------------
' Descricao   : Biblioteca utilizada no objeto Security
'               Exposto aqui para evitar Referências cíclicas
' Criaçao     : 16/01/2000 19:03
' Local       : Brasilia/DF
' Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
'               Ridai Govinda Pombo <ridai@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 1997, 1998, 1999, 2000 by Zevallos(r) Tecnologia em Informacao
'-------------------------------------------------------------------------

Option Explicit

Rem Variáveis e constantes utilizadas p/ encriptar
Private Const conLocalKey = "JcHIxjdk3z0Oq87TfrNu_wo1sEiDLPnZa.W9QeGUFYXRCKBtAhS65Mylvp4mb*Vg2$"
Private Const conLocalKeySize = 66

Private intNextChar As Long
Private intCont As Integer
Private intVerifier As Long

Public Function HTEncript(ByVal strText As String) As String
  Dim strResult As String, i As Integer
  intVerifier = 1
  intNextChar = 0
  intCont = 0

  For i = 1 To Len(strText)
    strResult = strResult & HTGetKey(Mid(strText, i, 1))

  Next
  If intCont > 0 Then
    Randomize
    strResult = strResult & Mid(conLocalKey, (intNextChar * (4 ^ (3 - intCont)) + ((Rnd * 100) Mod (4 ^ (3 - intCont))) + 1), 1)

  End If
  
  Do While intVerifier > 0
    strResult = Mid(conLocalKey, (intVerifier Mod conLocalKeySize) + 1, 1) & strResult
    intVerifier = intVerifier \ conLocalKeySize

  Loop
  HTEncript = strResult

End Function

Public Function HTUncript(ByVal strText As String) As String
  Dim strResult As String
  Dim intTestVerifier As Long, i As Integer
  
  intVerifier = 1
  intNextChar = 0
  intCont = 0
  intTestVerifier = 0

  For i = 1 To 3
    intTestVerifier = (intTestVerifier * conLocalKeySize) + InStr(conLocalKey, Mid(strText, i, 1)) - 1

  Next
  
  strText = Right(strText, Len(strText) - 3)
  Do While strText > ""
    strResult = strResult & HTGetValue(Left(strText, 4))
    If Len(strText) > 3 Then
      strText = Right(strText, Len(strText) - 4)
    
    Else
      strText = ""

    End If

  Loop
  
  If intVerifier = intTestVerifier Then
    HTUncript = strResult
  
  Else
    HTUncript = ""
  
  End If

End Function

Public Function HTGetKey(ByVal chrText As String)
  Dim strResult As String

  strResult = strResult & Mid(conLocalKey, (Asc(chrText) Mod conLocalKeySize) + 1, 1)
  intVerifier = ((intVerifier * Asc(chrText)) Mod (conLocalKeySize ^ 3)) + 1

  intCont = intCont + 1
  If intCont = 3 Then
    intNextChar = (intNextChar * 4) + (Asc(chrText) \ conLocalKeySize)
    strResult = strResult & Mid(conLocalKey, intNextChar + 1, 1)
    intNextChar = 0
    intCont = 0

  Else
    intNextChar = (intNextChar * 4) + (Asc(chrText) \ conLocalKeySize)

  End If

  HTGetKey = strResult

End Function

Public Function HTGetValue(ByVal strText As String)
  Dim strResult As String
  Dim intParc As Integer, i As Integer
  
  For i = 1 To (Len(strText) - 1)
    intParc = ((((InStr(conLocalKey, Right(strText, 1)) - 1) \ (4 ^ (3 - i))) Mod 4) * conLocalKeySize) - 1 + InStr(conLocalKey, Mid(strText, i, 1))
    If intParc < 255 And intParc > 0 Then
      strResult = strResult & Chr(intParc)
      intVerifier = ((intVerifier * intParc) Mod (conLocalKeySize ^ 3)) + 1

    End If
  Next
  HTGetValue = strResult

End Function
