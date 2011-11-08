VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Security"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------------------------------------------
' Security.cls
'-------------------------------------------------------------------------
' Descricao   : Objeto de Segurança
' Criacao     : 11:30h 02/03/2001
' Local       : Brasilia/DF
' Elaborado   : Eduardo Alves Cortes <edualves@zevallos.com.br>
'               Ruben Zevallos Jr. <zevallos@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 2001 by Zevallos(r) Tecnologia em Informacao
'-------------------------------------------------------------------------

Option Explicit

Private Const key = "JcHIxjdk3z0Oq87TfrNu_wo1sEiDLPnZa.W9QeGUFYXRCKBtAhS65Mylvp4mb*Vg2$"
Private Const KeySize = 66
Private intNextChar As Long
Private intCont As Integer
Private intVerifier As Long

Private Request As Request, _
        Response As Response, _
        Server As Server
        
'
Public QueryString As HiperTools30.Pack

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  Rem TimeBombX xtbSecurity, 15656, 12934, 3080, 14624
  
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  
  Set QueryString = Server.CreateObject("HiperTools30.Pack")
  
  QSCreate

End Sub

Public Sub OnEndPage()
  Set QueryString = Nothing
  
  Set Request = Nothing
  Set Response = Nothing
  Set Server = Nothing

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function URLEncript(ByVal strText As String) As String
    URLEncript = "ZSQS=" & Encript(strText)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function GetKey(ByVal chrText As String) As String
  Dim strResult As String

  strResult = strResult & Mid(key, (Asc(chrText) Mod KeySize) + 1, 1)
  intVerifier = ((intVerifier * Asc(chrText)) Mod (KeySize ^ 3)) + 1

  intCont = intCont + 1
  If intCont = 3 Then
    intNextChar = (intNextChar * 4) + (Asc(chrText) \ KeySize)
    strResult = strResult & Mid(key, intNextChar + 1, 1)
    intNextChar = 0
    intCont = 0

  Else
    intNextChar = (intNextChar * 4) + (Asc(chrText) \ KeySize)

  End If

  GetKey = strResult

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Encript(ByVal strText As String) As String
  Dim strResult As String, i As Integer
  
  intVerifier = 1
  intNextChar = 0
  intCont = 0

  For i = 1 To Len(strText)
    strResult = strResult & GetKey(Mid(strText, i, 1))

  Next
  If intCont > 0 Then
    Randomize
    strResult = strResult & Mid(key, (intNextChar * (4 ^ (3 - intCont)) + ((Rnd * 100) Mod (4 ^ (3 - intCont))) + 1), 1)

  End If
  
  Do While intVerifier > 0
    strResult = Mid(key, (intVerifier Mod KeySize) + 1, 1) & strResult
    intVerifier = intVerifier \ KeySize

  Loop
  Encript = strResult

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function GetValue(ByVal strText As String) As String
Dim strResult As String
Dim intParc As Integer, i As Integer
  
  For i = 1 To (Len(strText) - 1)
    intParc = ((((InStr(key, Right(strText, 1)) - 1) \ (4 ^ (3 - i))) Mod 4) * KeySize) - 1 + InStr(key, Mid(strText, i, 1))
    If intParc < 255 And intParc > 0 Then
      strResult = strResult & Chr(intParc)
      intVerifier = ((intVerifier * intParc) Mod (KeySize ^ 3)) + 1

    End If
  Next
  GetValue = strResult

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Uncript(ByVal strText As String) As String
Dim strResult As String
Dim intTestVerifier As Long, i As Integer
  intVerifier = 1
  intNextChar = 0
  intCont = 0
  intTestVerifier = 0

  For i = 1 To 3
    intTestVerifier = (intTestVerifier * KeySize) + InStr(key, Mid(strText, i, 1)) - 1

  Next
  
  strText = Right(strText, Len(strText) - 3)
  Do While strText > ""
    strResult = strResult & GetValue(Left(strText, 4))
    If Len(strText) > 3 Then
      strText = Right(strText, Len(strText) - 4)
    
    Else
      strText = ""

    End If

  Loop
  
  If intVerifier = intTestVerifier Then
    Uncript = strResult
  
  Else
    Uncript = ""
  
  End If

End Function

Rem =========================================================================
Rem Calcula o digito padrao modulo 11 dos bancos
Rem -------------------------------------------------------------------------
Public Function DVModulo11(ByVal strValor As String, Optional ByVal bytMax As Byte = 9) As Integer
  Dim lngPosicao As Long
  Dim bytMultiplicador As Integer
  Dim intSoma As Integer
  Dim intDigito As Integer
  
  intSoma = 0
  bytMultiplicador = 2
  
  For lngPosicao = Len(strValor) To 1 Step -1
    intSoma = intSoma + CInt(Mid(strValor, lngPosicao, 1)) * bytMultiplicador
    bytMultiplicador = bytMultiplicador + 1
    
    If bytMultiplicador = bytMax + 1 Then
      bytMultiplicador = 2
      
    End If
  Next
  
  intDigito = 11 - (intSoma Mod 11)
  
  If intDigito < 1 Or intDigito > 9 Then
    intDigito = 1
    
  End If

  DVModulo11 = intDigito
  
End Function

'Alterado (Ridai)
Rem =========================================================================
Rem Calcula o Modulo 11 com o X padrao Banco do Brasil
Rem -------------------------------------------------------------------------
Public Function DVModulo11X(ByVal strValor As String, Optional ByVal bytMax As Byte = 9) As String
  Dim lngPosicao As Long
  Dim bytMultiplicador As Byte
  Dim intSoma As Integer
  Dim strDigito As String
  
  intSoma = 0
  bytMultiplicador = 2
  
  For lngPosicao = Len(strValor) To 1 Step -1
    intSoma = intSoma + (CInt(Mid(strValor, lngPosicao, 1)) * bytMultiplicador)
    bytMultiplicador = bytMultiplicador + 1
    
    If bytMultiplicador = bytMax + 1 Then
      bytMultiplicador = 2
      
    End If
  Next
  
  strDigito = 11 - (intSoma Mod 11)
  
  If strDigito < 1 Or strDigito > 9 Then
    strDigito = "X"
    
  End If
  
  DVModulo11X = strDigito
  
End Function

Rem =========================================================================
Rem Calculo do Digito Verificador Modulo 10 padrao dos Bancos
Rem -------------------------------------------------------------------------
Public Function DVModulo10(ByVal strValor As String) As Integer
  Dim lngPosicao As Long
  Dim bytMultiplicador As Integer
  Dim intProduto As Integer
  Dim lngSoma As Long
  Dim intDigito As Integer
  
  lngSoma = 0
  bytMultiplicador = 2
  
  For lngPosicao = Len(strValor) To 1 Step -1
    intProduto = CInt(Mid(CStr(strValor), lngPosicao, 1)) * bytMultiplicador

    lngSoma = lngSoma + (intProduto \ 10) + (intProduto Mod 10)

    bytMultiplicador = bytMultiplicador - 1

    If bytMultiplicador = 0 Then
      bytMultiplicador = 2

    End If
  Next
  
  intDigito = 10 - (lngSoma Mod 10)

  If intDigito = 10 Then
    intDigito = 0
      
  End If
  
  DVModulo10 = intDigito

End Function

Private Sub QSCreate()
  Dim strQueryString As String
  Dim intPos As Integer
  Dim strValor As String
  Dim strQuery As String, i As Variant

  For Each i In Request.QueryString
    If i <> "ZSQS" Then
      QueryString.Add i, Request.QueryString(i)

    Else
      strQueryString = Uncript(Request.QueryString(i))
      Do While strQueryString > ""
        intPos = InStr(strQueryString, "=")
        If intPos > 0 Then
          If Left(strQueryString, 1) = "&" Then
            strQueryString = Right(strQueryString, Len(strQueryString) - 1)

          End If
          strQuery = Left(strQueryString, intPos - 1)
          strQueryString = Right(strQueryString, Len(strQueryString) - intPos)
          intPos = InStr(strQueryString, "&")
          If intPos > 0 Then
            strValor = Left(strQueryString, intPos - 1)
            strQueryString = Right(strQueryString, Len(strQueryString) - intPos)

          Else
            strValor = strQueryString
            strQueryString = ""

          End If
          QueryString.Remove strQuery
          QueryString.Add strQuery, strValor

        Else
          strQueryString = ""

        End If
      Loop
    End If
  Next

End Sub

