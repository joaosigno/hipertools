Attribute VB_Name = "String_lib"
Option Explicit

Private sstrPieceString As String
Private sintPieceStringPosition  As Integer

Rem -------------------------------------------------------------------------
Rem Normaliza String retornando apenas os números
Rem -------------------------------------------------------------------------
Public Function HTNormalizeNumber(ByVal strValue As String, Optional ByVal blnFix As Boolean = False) As String
  Dim Target As String, _
      tempstr As String, _
      cOneChar As String, _
      i As Integer
    
  Target = ""
  tempstr = UCase(Trim(strValue))
  
  If tempstr > "" Then
    For i = 1 To Len(tempstr)
      cOneChar = Mid(tempstr, i, 1)
      'Alterado (Ridai Govinda)
      If (cOneChar >= "0" And cOneChar <= "9") Or (blnFix And i < 2 And cOneChar = "-") Then
        Target = Target & cOneChar
      End If
    Next
  End If

  HTNormalizeNumber = Target
End Function

Rem -------------------------------------------------------------------------
Rem Formata strings com o padrao de CGC 999.999.999/9999-99
Rem -------------------------------------------------------------------------
Public Function HTFormatCGC(ByVal strString As String) As String
  Dim strResult, i, strRest
  
  strResult = ""

  strString = HTNormalizeNumber(strString)
  HTSetPiece HTReverse(strString)
 
  If strString > "" Then
  
    strResult = HTGetPiece(2) & "-" & HTGetPiece(4) & "/"

    For i = 7 To Len(strString) Step 3
      strResult = strResult & HTGetPiece(3) & "."
  
    Next
  
  
    strRest = HTGetPieceRest
  
    If strRest > "" Then
      strResult = strResult & strRest
     
    Else
      strResult = Left(strResult, Len(strResult) - 1)
    
    End If
  
  End If
  
  
  HTFormatCGC = HTReverse(strResult)
  
End Function
Rem -------------------------------------------------------------------------
Rem Final da Function ZTIFormatCGC

Public Function HTFormatCPF(ByVal strString As String) As String
  Dim strResult, i, strRest
  
  strResult = ""
  
  strString = HTNormalizeNumber(strString)
  HTSetPiece HTReverse(strString)
 
  If strString > "" Then
  
    strResult = HTGetPiece(2) & "-"

    For i = 3 To Len(strString) Step 3
      strResult = strResult & HTGetPiece(3) & "."
  
    Next
  
    strRest = HTGetPieceRest
  
    If strRest > "" Then
      strResult = strResult & strRest
     
    Else
      strResult = Left(strResult, Len(strResult) - 1)
    
    End If
  
  End If
  
  
  HTFormatCPF = HTReverse(strResult)
  
End Function
Rem -------------------------------------------------------------------------
Rem Final da Function ZTIFormatCPF

Public Function HTReverse(ByVal strString As String) As String
  Dim i As Integer
  Dim strTarget As String
    
  strTarget = ""

  If strString > "" Then
    For i = Len(strString) To 1 Step -1
      strTarget = strTarget & Mid(strString, i, 1)
    Next
    
  End If
    
  HTReverse = strTarget
End Function
Rem -------------------------------------------------------------------------
Rem Final da Function ZTIReverse


Rem -------------------------------------------------------------------------
Rem Obtem a proxima string
Rem -------------------------------------------------------------------------
Public Sub HTSetPiece(ByVal strString As String)
  Dim intPos

  Do
    intPos = InStr(strString, Chr(13))
    
    If intPos > 0 Then
      strString = Left(strString, intPos - 1) & Mid(strString, intPos + 1)
      
    End If
  Loop While intPos > 0

    sstrPieceString = strString

  sintPieceStringPosition = 1
    
End Sub

Rem -------------------------------------------------------------------------
Rem Obtem a string inteira
Rem -------------------------------------------------------------------------
Public Function HTGetEntirePiece() As String
    HTGetEntirePiece = sstrPieceString
    
End Function

Rem -------------------------------------------------------------------------
Rem Obtem a proxima string
Rem -------------------------------------------------------------------------
Public Function HTGetPiece(ByVal intSize As Integer) As String
    HTGetPiece = Trim(Mid(sstrPieceString, sintPieceStringPosition, intSize))
    
    HTSkipPiece intSize

End Function

Rem -------------------------------------------------------------------------
Rem Obtem a proxima string
Rem -------------------------------------------------------------------------
Public Function HTGetPieceRest() As String
    HTGetPieceRest = Trim(Mid(sstrPieceString, sintPieceStringPosition))
    
End Function

Rem -------------------------------------------------------------------------
Rem Obtem a proxima string colocando-a entre delimitadores de texto
Rem -------------------------------------------------------------------------
Public Function HTGetPieceAsStr(ByVal intSize As Integer) As String
    HTGetPieceAsStr = "'" & HTGetPiece(intSize) & "'"
    
End Function

Rem -------------------------------------------------------------------------
Rem Obtem a proxima string formatando-a como data
Rem -------------------------------------------------------------------------
Public Function HTGetPieceAsDate() As String
    HTGetPieceAsDate = Chr(34) & HTGetPiece(8) & Chr(34)
    
End Function

Rem -------------------------------------------------------------------------
Rem Salta a proxima string
Rem -------------------------------------------------------------------------
Public Sub HTSkipPiece(ByVal intSize As Integer)

    sintPieceStringPosition = sintPieceStringPosition + intSize

End Sub

Rem -------------------------------------------------------------------------
Rem Format number with leading zeroes
Rem -------------------------------------------------------------------------
'Alterado (Ridai Govinda) - Agora suportando números negativos
Public Function HTLeadingZeroes(ByVal strNumber As String, ByVal intPlaces As Integer) As String
  
  Dim blnNegative As Boolean
  
  blnNegative = False
  strNumber = Trim(strNumber)
  
  If intPlaces >= Len(strNumber) Then
    If InStr(strNumber, "-") > 0 Then
      strNumber = HTCropLeft(strNumber, 1)
      blnNegative = True
    End If
    
    If blnNegative Then intPlaces = intPlaces - 1
    If intPlaces < 0 Then intPlaces = 0
    
    HTLeadingZeroes = IIf(blnNegative, "-", "") & String(intPlaces - Len(strNumber), "0") & strNumber
  Else
    HTLeadingZeroes = strNumber
  End If

End Function

Rem -------------------------------------------------------------------------
Rem retorna uma string composta pela repeticao de uma outra string
Rem -------------------------------------------------------------------------
Public Function HTReplicate(ByVal strChar As String, ByVal intCount As Integer)
    Dim i As Integer, strReturn As String
    
    For i = 1 To intCount
        strReturn = strReturn & strChar
    Next
    
    HTReplicate = strReturn
End Function


Rem -------------------------------------------------------------------------
Rem retorna uma string removendo caracteres determinados do início e do fim
Rem -------------------------------------------------------------------------
Public Function HTRemoveChar(ByVal strValue As String, ByVal strBefore As String, Optional ByVal strAfter As String = vbNullString)
    If strAfter = vbNullString Then
        strValue = Replace(strValue, strBefore, "")
        
    Else
        If Left(strValue, 1) = strBefore Then strValue = Right(strValue, Len(strValue) - 1)
        If Right(strValue, 1) = strAfter Then strValue = Left(strValue, Len(strValue) - 1)
    
    End If
    
    HTRemoveChar = strValue
End Function


Public Function HTFormatInt(ByVal intValue As Long) As String
  Dim Reais As Long, _
      Centena As Long, _
      Negative As Boolean, _
      Target As String

  Reais = Abs(intValue)
  Negative = intValue < 0
  Target = ""

  Do While Reais > 0
    Centena = Reais Mod 1000
    Reais = Reais \ 1000

    If Reais Then
      Target = "." & HTLeadingZeroes(Centena, 3) & Target
            
    Else
      Target = Trim(CStr(Centena)) & Target

    End If
  Loop
    
  If Target = "" Then
    Target = "0"

  End If

  Target = Target

  If Negative Then
    Target = "(" & Target & ")"

  End If

  HTFormatInt = Target

End Function

'Criado (Ridai Govinda)
Public Function HTCropLeft(ByVal strString As String, ByVal intChar As Integer) As String
    
  HTCropLeft = strString
  
  If strString > "" Then
    HTCropLeft = Right(strString, Len(strString) - intChar)
    
  End If
  
End Function

Public Function HTCropRight(ByVal strString As String, ByVal intChar As Integer) As String
  If strString > "" Then
    HTCropRight = Left(strString, Len(strString) - intChar)
  
  Else
    HTCropRight = strString
  
  End If

End Function

Public Function HTFormatMoney(ByVal monValue As Currency) As String
    HTFormatMoney = "R$ " & HTFormatNumber(monValue)
    
End Function

Public Function HTFormatNumber(ByVal monValue As Currency, Optional ByVal intCasas As Integer = 2) As String
    Dim Centavos As String
    Dim Reais As Long, Centena As String
    Dim Target As String
    Dim Negative As Boolean

    Centavos = HTLeadingZeroes(Abs(Round(monValue, intCasas) - Fix(monValue)) * (10 ^ intCasas), intCasas)
    
    Reais = Abs(Fix(monValue))
    
    If Len(Centavos) > intCasas Then
      If Left(Centavos, 1) = "1" Then Reais = Reais + 1
      Centavos = String(intCasas, "0")
    End If
    
    Negative = monValue < 0
    Target = ""
    
    Do While Reais > 0
        Centena = Reais Mod 1000
        Reais = Reais \ 1000
        
        If Reais Then
            Target = "." & HTLeadingZeroes(Centena, 3) & Target

        Else
            Target = Trim(CStr(Centena)) & Target

        End If
    Loop
    
    If Target = "" Then
        Target = "0"
    End If
    
    Target = Target & "," & Centavos
    
    If Negative Then
        Target = "(" & Target & ")"
        
    End If
    
    HTFormatNumber = Target
    
End Function

Public Function HTFormatText(ByVal strMask As String, ParamArray arrStrings() As Variant) As String
  Const lconReplaceStr = "$"
  Const lconIndexStr = ":"
  Const lconLeftJustify = "-"
  Const lconDecSeparator = "."
  
  Const lconTypeDecimal = "d"
  Const lconTypeUnsigned = "u"
  Const lconTypeScientific = "e"
  Const lconTypeFixed = "f"
  Const lconTypeGeneral = "g"
  Const lconTypeNumber = "n"
  Const lconTypeMoney = "m"
  Const lconTypeString = "s"
  Const lconTypeHexa = "x"
  
  Dim strResult As String, strBuff As String, strChar As String
  Dim i As Long, j As Long, lngIndex As Long, lngAux As Long
  Dim intWidth As Integer, intPrec As Integer
  Dim blnEndType As Boolean, blnLeftJust As Boolean, blnOnPrec As Boolean
  
  blnEndType = True
  blnLeftJust = False
  blnOnPrec = False
  
  i = 1
  j = -1
  lngIndex = -1
  intPrec = 0
  
  strResult = ""
  
  If Not IsMissing(arrStrings) Then
    Do While i <= Len(strMask)
      strChar = Mid(strMask, i, 1)
      
      If strChar = lconReplaceStr Or Not blnEndType Then
        blnEndType = False
        i = i + 1
        
        strChar = Mid(strMask, i, 1)
        
        Select Case strChar
        Case lconTypeDecimal, lconTypeUnsigned, lconTypeUnsigned, _
             lconTypeScientific, lconTypeFixed, lconTypeGeneral, _
             lconTypeNumber, lconTypeMoney, lconTypeString, lconTypeHexa
          
          j = j + 1
          
          blnEndType = True
          If strBuff > "" Then
            If blnOnPrec Then
              intPrec = CInt(strBuff)
            Else
              intWidth = CInt(strBuff)
            End If
          End If
          
          'Guarda o valor de j...
          If lngIndex > -1 Then
            lngAux = j
            j = lngIndex
          End If
          
          If j <= UBound(arrStrings) Then
            strChar = GetFormatedString(strChar, arrStrings(j), blnLeftJust, intWidth, intPrec)
          Else
            strChar = ""
          End If
          
          'Recupera o valor de j
          If lngIndex > -1 Then j = lngAux
          
          lngIndex = -1
          intPrec = 0
          intWidth = 0
          blnLeftJust = False
          blnOnPrec = False
          strBuff = ""
        
        Case lconLeftJustify
          blnLeftJust = True
        
        Case lconDecSeparator
          If strBuff > "" Then
            intWidth = CInt(strBuff)
            blnOnPrec = True
            strBuff = ""
          End If
        
        Case lconIndexStr
          If Not blnEndType Then
            lngIndex = CLng(strBuff)
            strBuff = ""
          End If
        
        Case lconReplaceStr
          i = i + 1
          strResult = strResult & lconReplaceStr
          blnEndType = True
        
        Case Else
          If IsNumeric(strChar) And Not blnEndType Then
            strBuff = strBuff & strChar
          End If
        
        End Select
      
      End If
      
      If blnEndType Then
        strResult = strResult & strChar
        i = i + 1
      End If
      
    Loop
  End If
  
  HTFormatText = strResult

End Function

Private Function GetFormatedString(ByVal strType As String, ByVal vntValue As Variant, _
                                   ByVal blnLeftJustify As Boolean, _
                                   ByVal intWidth As Integer, ByVal intPrec As Integer) As String
  Const lconTypeDecimal = "d"
  Const lconTypeUnsigned = "u"
  Const lconTypeScientific = "e"
  Const lconTypeFixed = "f"
  Const lconTypeGeneral = "g"
  Const lconTypeNumber = "n"
  Const lconTypeMoney = "m"
  Const lconTypeString = "s"
  Const lconTypeHexa = "x"
  
  Dim strResult As String, strFormat As String
  
  strResult = vntValue
  strFormat = ""
  
  Select Case strType
  Case lconTypeDecimal
    Select Case LCase(TypeName(vntValue))
    Case "integer", "long"
      strResult = HTLeadingZeroes(CStr(vntValue), intWidth)
    End Select
  
  Case lconTypeUnsigned
    Select Case LCase(TypeName(vntValue))
    Case "integer", "long"
      strResult = CStr(Abs(vntValue))
      strResult = HTLeadingZeroes(strResult, intWidth)
    End Select
    
  Case lconTypeScientific
    Select Case LCase(TypeName(vntValue))
    Case "double", "single", "currency", "long"
      strResult = Format(CCur(vntValue), "Scientific")
    End Select
  
  Case lconTypeFixed
    Select Case LCase(TypeName(vntValue))
    Case "double", "single", "long", "currency", "integer"
      strFormat = String(intWidth, "0") & "." & String(intPrec, "0")
      strResult = Format(vntValue, strFormat)
    End Select
  
  Case lconTypeGeneral
    Select Case LCase(TypeName(vntValue))
    Case "double", "single", "long", "currency", "integer"
      If intPrec <= 0 Then intPrec = 15
      strFormat = String(intWidth, "0") & "." & String(intPrec, "0")
      strResult = Format(vntValue, strFormat)
    End Select
  
  Case lconTypeNumber
    Select Case LCase(TypeName(vntValue))
    Case "double", "single", "currency"
      strResult = HTFormatNumber(CCur(vntValue))
    End Select
  
  Case lconTypeMoney
    Select Case LCase(TypeName(vntValue))
    Case "double", "single", "currency"
      strResult = HTFormatMoney(CCur(vntValue))
    End Select
  
  Case lconTypeString
    strResult = Mid(CStr(vntValue), 1, intWidth)
  
  Case lconTypeHexa
    strResult = Mid(Hex(CStr(vntValue)), 1, intWidth)
  
  End Select
  
  If blnLeftJustify Then
    If intWidth > Len(strResult) Then
      strResult = strResult & Space(intWidth - Len(strResult))
    End If
  End If
  
  GetFormatedString = strResult

End Function
'----------------------------

