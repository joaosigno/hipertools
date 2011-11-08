VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "String"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =========================================================================
' String.cls
' -------------------------------------------------------------------------
' Descricao   : Include de String
' Criacao     : 11:23h 7/2/1998
' Local       : Brasilia/DF
' Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------
Option Explicit

Private sstrComments As String, _
        sstrTitle As String, _
        sstrVersion As String, _
        sstrCompanyName As String, _
        sstrProductName As String, _
        sstrLegalCopyright As String

'Criado (Ridai Govinda)
Private sstrText3 As String

Private Request As Request, _
        Response As Response, _
        Server As Server, _
        Session As Session, _
        Application As Application, _
        ScriptContext As ScriptingContext

Private sstrDateSeparator As String

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  TimeBombX xtbString, 15656, 12934, 3080, 14624
  
  Set ScriptContext = ScriptContext
  
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application
  
  sstrDateSeparator = "/"
  'Criado (Ridai Govinda)
  sstrText3 = ""
  '--------------------

End Sub

Public Sub OnEndPage()

  Set ScriptContext = Nothing
  
  Set Request = Nothing
  Set Response = Nothing
  Set Server = Nothing
  Set Session = Nothing
  Set Application = Nothing
End Sub

Private Sub Class_Initialize()

    sstrComments = App.Comments
    sstrTitle = App.Title
    sstrVersion = App.Major & "." & _
                App.Minor & "." & _
                App.Revision
    sstrCompanyName = App.CompanyName
    sstrProductName = App.ProductName

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get DateSeparator() As String
  DateSeparator = sstrDateSeparator
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Replicate(ByVal strChar As String, ByVal intCount As Integer)
    Replicate = HTReplicate(strChar, intCount)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let DateSeparator(ByVal strDateSeparator As String)
  sstrDateSeparator = strDateSeparator
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Piece() As String
  Piece = HTGetEntirePiece
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Piece(ByVal strPiece As String)
  HTSetPiece strPiece
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub SetPiece(ByVal strString As String)
  HTSetPiece strString
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub ShowPiece()
    Response.Write "<BR>" & GetEntirePiece
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub ShowPieceRest()
    Response.Write "<BR>" & GetPieceRest
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function GetEntirePiece() As String
    GetEntirePiece = HTGetEntirePiece
    
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function GetPiece(ByVal intSize As Integer) As String
    GetPiece = HTGetPiece(intSize)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function GetPieceRest() As String
    GetPieceRest = HTGetPieceRest
End Function

Rem =========================================================================
Rem Obtem a proxima string colocando-a entre delimitadores de texto
Rem -------------------------------------------------------------------------
Public Function GetPieceAsStr(ByVal intSize As Integer) As String
    GetPieceAsStr = "'" & GetPiece(intSize) & "'"
End Function

Rem =========================================================================
Rem Obtem a proxima string formatando-a como data
Rem -------------------------------------------------------------------------
Public Function GetPieceAsDate() As String
    GetPieceAsDate = HTGetPieceAsDate()
End Function

Rem =========================================================================
Rem Salta a proxima string
Rem -------------------------------------------------------------------------
Public Sub SkipPiece(ByVal intSize As Integer)
    HTSkipPiece intSize
End Sub

Rem =========================================================================
Rem Obtem n-esima palavra de uma string
Rem -------------------------------------------------------------------------
Public Function GetWord(ByVal strPhrase As String, ByVal intWord As Integer) As String
    Dim blnOnWord As Boolean, _
        nWordCount As Integer, _
        nPos As Integer, _
        nLength As Integer, _
        strWord As String, _
        cOneChar As String, _
        strTempPhrase As String, _
        strTempPhrase2 As String
   
    blnOnWord = False
    nWordCount = 0
    nPos = 1
    strTempPhrase = Trim(strPhrase)
    strTempPhrase2 = ChangeChar(UCase(Trim(strPhrase)))
    nLength = Len(strTempPhrase)
    
    Do While nPos <= nLength
        cOneChar = Mid(strTempPhrase2, nPos, 1)
        If (cOneChar >= "A" And cOneChar <= "Z") Or (cOneChar >= "0" And cOneChar <= "9") Then
            If blnOnWord Then
                strWord = strWord & Mid(strTempPhrase, nPos, 1)
            Else
                blnOnWord = True
                nWordCount = nWordCount + 1
                strWord = Mid(strTempPhrase, nPos, 1)
            End If
        Else
            If blnOnWord Then
                blnOnWord = False
                If nWordCount = intWord Then
                    Exit Do
                End If
            End If
        End If
        nPos = nPos + 1
    Loop
    
    If nWordCount <> intWord Then
        strWord = ""
    End If
    
    GetWord = strWord
End Function

Rem =========================================================================
Rem Retorna a quantidade de palavras contidas da frase
Rem -------------------------------------------------------------------------
Public Function GetWordCount(ByVal strPhrase As String) As Long
    Dim blnOnWord As Boolean, _
        intWordCount As Integer, _
        nPos As Integer, _
        cOneChar As String, _
        strTempPhrase As String

    blnOnWord = False
    intWordCount = 0
    nPos = 1
    strTempPhrase = ChangeChar(UCase(Trim(strPhrase)))
    
    Do While nPos <= Len(strTempPhrase)
       cOneChar = Mid(strTempPhrase, nPos, 1)
       If (cOneChar >= "A" And cOneChar <= "Z") Or (cOneChar >= "0" And cOneChar <= "9") Then
          If Not blnOnWord Then
             blnOnWord = True
             intWordCount = intWordCount + 1
          End If
       Else
          If blnOnWord Then
             blnOnWord = False
          End If
       End If
       nPos = nPos + 1
    Loop
    GetWordCount = intWordCount
End Function

Rem =========================================================================
Rem Checa as Strings e coloca os " ' " para a chamada SQL
Rem -------------------------------------------------------------------------
Public Function CheckString(ByVal strString As String, Optional ByVal strEndchar As String = "") As String
  'Criado (Ridai Govinda)
  Dim nPos As Long
    
  nPos = InStr(strString, "'")

  While nPos > 0
      strString = Mid(strString, 1, nPos) & "'" & Mid(strString, nPos + 1)

      nPos = InStr(nPos + 2, strString, "'")

  Wend

  CheckString = "'" & strString & "'" & strEndchar & " "

End Function

Rem =========================================================================
Rem Prepara valores numericos para serem adicionados a uma chamada SQL INSERT
Rem -------------------------------------------------------------------------
Public Function CheckNumber(ByVal vntValue As Variant, Optional ByVal strEndchar As String = "") As String
  Dim intPos As Integer, _
      strChar As String, _
      strResult As String, _
      strValue As String
  
  strResult = ""
  strValue = Trim(CStr(vntValue))
  For intPos = 1 To Len(strValue)
    strChar = Mid(strValue, intPos, 1)
    If (strChar >= "0" And strChar <= "9") Or (strChar = ".") Or (strChar = ",") Then
      If (strChar <> ",") Then
         strResult = strResult & strChar
      End If
    Else
      Exit For
    End If
  Next

  If Not IsNumeric(strResult) Then
    strResult = "NULL"
  End If
  
  CheckNumber = strResult & strEndchar
End Function

Rem =========================================================================
Rem Prepara Datas do Sistema para serem adicionados a uma chamada SQL INSERT
Rem -------------------------------------------------------------------------
Public Function CheckSystemDate(ByVal strDate As String, Optional ByVal strEndchar As String = "") As String
    CheckSystemDate = CheckString(strDate, strEndchar)
End Function

Rem =========================================================================
Rem Prepara datas AAMMDD para serem adicionadas a uma chamada SQL INSERT no
Rem formato MMDDAA
Rem -------------------------------------------------------------------------
Public Function CheckDate(ByVal strDate As String, Optional ByVal strEndchar As String = "") As String
    Dim strDia As String, _
        strMes As String, _
        strAno As String

    strAno = GetWord(strDate, 1)
    strMes = GetWord(strDate, 2)
    strDia = GetWord(strDate, 3)

    strDate = strMes & sstrDateSeparator & strDia & sstrDateSeparator & strAno

    If IsDate(strDate) Then
        strDate = "'" & strDate & "'"
    Else
        strDate = "NULL"
    End If

    CheckDate = strDate & strEndchar
End Function

Rem =========================================================================
Rem Prepara datas para serem utilizadas dentro de uma Query SQL
Rem -------------------------------------------------------------------------
Public Function ImportDate(ByVal strDate As String, ByVal intDateType As Integer) As Date

  Dim strDia As String, _
      strMes As String, _
      strAno As String, _
      intDia As Integer, _
      intMes As Integer, _
      intAno As Integer, _
      intAnoSize As Integer

  strDate = NormalizeString(Trim(strDate))
  
  intAnoSize = 2

  Select Case intDateType
    Case stDateTypeAAMMDD
      intDia = 5
      intMes = 3
      intAno = 1
      
    Case stDateTypeAAAAMMDD
      intDia = 7
      intMes = 5
      intAno = 1
      intAnoSize = 4
      
    Case stDateTypeMMDDAA
      intDia = 3
      intMes = 1
      intAno = 5

    Case stDateTypeMMDDAAAA
      intDia = 3
      intMes = 1
      intAno = 5
      intAnoSize = 4
      
    Case stDateTypeDDMMAA
      intDia = 1
      intMes = 3
      intAno = 5

    Case stDateTypeDDMMAAAA
      intDia = 1
      intMes = 3
      intAno = 5
      intAnoSize = 4

  End Select

  strAno = Mid(strDate, intAno, intAnoSize)
  strMes = Mid(strDate, intMes, 2)
  strDia = Mid(strDate, intDia, 2)

  strDate = strMes & sstrDateSeparator & strDia & sstrDateSeparator & strAno

  If IsDate(strDate) Then
    ImportDate = CDate(strDate)
  End If

End Function

Rem =========================================================================
Rem Prepara Money para serem utilizadas dentro de uma Query SQL
Rem -------------------------------------------------------------------------
Public Function ImportMoney(ByVal strMoney As String, ByVal intMoneyType As Integer) As String
  
  strMoney = NormalizeNumber(Trim(strMoney))
  
  If Len(strMoney) > 0 Then
    If Fix(CDbl(strMoney)) > 0 Then
      Select Case intMoneyType
        Case stMoneyType9999D99, stMoneyType9D999D99
          ImportMoney = Fix(Left(strMoney, Len(strMoney) - 2)) & "." & Right(strMoney, 2)
        Case Else
          ImportMoney = strMoney
      End Select
    Else
        ImportMoney = 0#
    End If
  Else
    ImportMoney = 0#
  End If

End Function

Rem =========================================================================
Rem Prepara um Int para ser importado
Rem -------------------------------------------------------------------------
Public Function ImportInt(ByVal strInteger As String) As Integer

  strInteger = NormalizeString(Trim(strInteger))
  
  If Len(strInteger) > 0 And IsNumeric(strInteger) Then
    ImportInt = Fix(strInteger)
  Else
    ImportInt = 0
  End If

End Function

Rem =========================================================================
Rem Prepara um Int para ser importado
Rem -------------------------------------------------------------------------
Public Function ImportLong(ByVal strLong As String) As Long

  strLong = NormalizeString(Trim(strLong))
  
  If Len(strLong) > 0 And IsNumeric(strLong) Then
    ImportLong = CLng(strLong)
  Else
    ImportLong = 0
  End If

End Function

Rem =========================================================================
Rem Prepara um Int para ser importado
Rem -------------------------------------------------------------------------
Public Function ImportDouble(ByVal strDouble As String) As Double

  strDouble = NormalizeString(Trim(strDouble))
  
  If Len(strDouble) > 0 And IsNumeric(strDouble) Then
    ImportDouble = CDbl(strDouble)
  Else
    ImportDouble = 0
  End If

End Function

Rem =========================================================================
Rem Format number with leading zeroes
Rem -------------------------------------------------------------------------
Public Function LeadingZeroes(ByVal strNumber As String, ByVal intPlaces As Integer) As String
  LeadingZeroes = HTLeadingZeroes(strNumber, intPlaces)
End Function

Rem =========================================================================
Rem Apaga os zeros a esquerda
Rem -------------------------------------------------------------------------
Public Function LTrimZeroes(ByVal strValue As String) As String

  strValue = LTrim(strValue)

  Do While Len(strValue) > 0
    If Left(strValue, 1) = "0" Then
      strValue = Right(strValue, Len(strValue) - 1)
    Else
      Exit Do
    End If
  Loop
  
  LTrimZeroes = strValue

End Function

Rem =========================================================================
Rem Devolve um valor monetario completo R$#.###,##
Rem -------------------------------------------------------------------------
Public Function ZTIMoney(ByVal monValue As Currency) As String
    ZTIMoney = Money(monValue)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function FormatMoney(ByVal monValue As Currency, Optional ByVal intCasas As Integer = 2) As String
    FormatMoney = "R$ " & FormatNumber(monValue, intCasas)
    
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Money(ByVal monValue As Currency) As String
    Money = "R$ " & number(monValue)
    
End Function

Rem =========================================================================
Rem Devolve um valor monetario completo #.###,##
Rem -------------------------------------------------------------------------
Public Function ZTINumber(ByVal monValue As Currency) As String
    ZTINumber = FormatNumber(monValue)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function number(ByVal monValue As Currency) As String
  number = FormatNumber(monValue)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function FormatNumber(ByVal monValue As Currency, Optional ByVal intCasas As Integer = 2) As String
  FormatNumber = HTFormatNumber(monValue, intCasas)
End Function

Rem =========================================================================
Rem Devolve um valor numerico completo #.###
Rem -------------------------------------------------------------------------
Public Function ZTIInt(ByVal intValue As Long) As String
    ZTIInt = FormatInt(intValue)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function FormatInt(ByVal lngValue As Long) As String
  FormatInt = HTFormatInt(lngValue)
End Function

Rem =========================================================================
Rem Devolve a data e a hora completos DD/MM/YYYY HH:MM:SS
Rem -------------------------------------------------------------------------
Public Function ZTILongDateTime(ByVal vntDateTime As Variant) As String
    ZTILongDateTime = LongDateTime(vntDateTime)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function LongDateTime(ByVal vntDateTime As Variant) As String
    LongDateTime = LongDate(vntDateTime) & " " & LongTime(vntDateTime)
End Function

Rem =========================================================================
Rem Devolve a data longa DD/MM/YYYY
Rem -------------------------------------------------------------------------
Public Function ZTILongDate(ByVal vntDate As Variant) As String
    ZTILongDate = LongDate(vntDate)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function LongDate(ByVal vntDate As Variant) As String
  If IsDate(vntDate) Then
    LongDate = LeadingZeroes(Day(vntDate), 2) & sstrDateSeparator & _
                  LeadingZeroes(Month(vntDate), 2) & sstrDateSeparator & _
                  LeadingZeroes(Year(vntDate), 4)

  End If

End Function

Rem =========================================================================
Rem Devolve a data longa MM/DD/YYYY
Rem -------------------------------------------------------------------------
Public Function ZTIUSADate(ByVal vntDate As Variant) As String
    ZTIUSADate = USADate(vntDate)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function USADate(ByVal vntDate As Variant) As String
  If IsDate(vntDate) Then
    USADate = LeadingZeroes(Month(vntDate), 2) & sstrDateSeparator & _
                 LeadingZeroes(Day(vntDate), 2) & sstrDateSeparator & _
                 LeadingZeroes(Year(vntDate), 4)
                  
  End If
    
End Function

Rem =========================================================================
Rem Devolve a hora longa 24H
Rem -------------------------------------------------------------------------
Public Function ZTILongTime(ByVal vntTime As Variant) As String
    ZTILongTime = LongTime(vntTime)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function LongTime(ByVal vntTime As Variant) As String
    If IsDate(vntTime) Then
        LongTime = LeadingZeroes(Hour(vntTime), 2) & ":" & _
                   LeadingZeroes(Minute(vntTime), 2) & ":" & _
                   LeadingZeroes(Second(vntTime), 2)
        
    End If
    
End Function

Rem =========================================================================
Rem Devolve a hora curta 24H
Rem -------------------------------------------------------------------------
Public Function ZTIShortTime(ByVal vntTime As Variant) As String
    ZTIShortTime = ShortTime(vntTime)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function ShortTime(ByVal vntTime As Variant) As String
    If IsDate(vntTime) Then
      ShortTime = LeadingZeroes(Hour(vntTime), 2) & ":" & _
                  LeadingZeroes(Minute(vntTime), 2)
                              
    End If
    
End Function

Rem =========================================================================
Rem Devolve o numero de segundos de uma hora
Rem -------------------------------------------------------------------------
Public Function ZTITimeSeconds(ByVal vntTime As Variant) As Long
    ZTITimeSeconds = TimeAsSeconds(vntTime)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function TimeAsSeconds(ByVal vntTime As Variant) As Long
    TimeAsSeconds = Hour(vntTime) * 3600 + Minute(vntTime) * 60 + Second(vntTime)
End Function

Rem =========================================================================
Rem Devolve o numero de segundos de uma hora
Rem -------------------------------------------------------------------------
Public Function ZTISecondsTime(ByVal intSeconds As Integer) As Date
    ZTISecondsTime = SecondsAsTime(intSeconds)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function SecondsAsTime(ByVal lngSeconds As Long) As Date
    SecondsAsTime = TimeSerial(Int(lngSeconds / 3600), Int(lngSeconds / 216000), lngSeconds Mod 3600)
End Function

Rem =========================================================================
Rem Prepara datas AAMMDD para serem adicionadas a uma chamada SQL INSERT
Rem -------------------------------------------------------------------------
Public Function SQLLongDate(ByVal vntDate As Variant) As String

  SQLLongDate = "'" & LeadingZeroes(Month(vntDate), 2) & sstrDateSeparator & LeadingZeroes(Day(vntDate), 2) & sstrDateSeparator & LeadingZeroes(Year(vntDate), 4) & "'"
    
End Function

Rem =========================================================================
Rem Prepara datas MMDDAA para serem adicionadas a uma chamada SQL INSERT - Brasil
Rem -------------------------------------------------------------------------
Public Function SQLLongDateBR(ByVal vntDate As Variant) As String

  SQLLongDateBR = "'" & LeadingZeroes(Day(vntDate), 2) & sstrDateSeparator & LeadingZeroes(Month(vntDate), 2) & sstrDateSeparator & LeadingZeroes(Year(vntDate), 4) & "'"

End Function

Rem =========================================================================
Rem Prepara data e hora para serem adicionadas a uma chamada SQL
Rem -------------------------------------------------------------------------
Public Function SQLDateTime(ByVal vntDateTime As Variant) As String
    SQLDateTime = "'" & LongDateTime(vntDateTime) & "'"
End Function

Rem =========================================================================
Rem Prepara um Int para ser importado
Rem -------------------------------------------------------------------------
Public Function SQLLong(ByVal lngValue As Long) As String
  lngValue = NormalizeString(Trim(lngValue))
  
  If Len(lngValue) > 0 Then
    If IsNumeric(lngValue) Then
      lngValue = CLng(lngValue)
  
    Else
      lngValue = 0
    
    End If
  Else
    lngValue = 0
    
  End If
  
  SQLLong = lngValue
                    
End Function

Rem =========================================================================
Rem Prepara um Int para ser importado
Rem -------------------------------------------------------------------------
Public Function SQLString(ByVal strString As String) As String
  SQLString = "'" & Trim(strString) & "'"
                    
End Function

Rem =========================================================================
Rem Normaliza String
Rem -------------------------------------------------------------------------
Public Function NormalizeString(ByVal strValue As String) As String

    Dim Target As String, _
        tempstr As String, _
        cOneChar As String, _
        i As Integer


  If strValue > "" Then
    Target = ""
    'tempstr = UCase(Trim(strValue))
    tempstr = Trim(strValue)
  
    For i = 1 To Len(tempstr)
        cOneChar = Mid(tempstr, i, 1)
        If (cOneChar >= "A" And cOneChar <= "Z") Or (cOneChar >= "a" And cOneChar <= "z") Or (cOneChar >= "0" And cOneChar <= "9") Then
            Target = Target & cOneChar
        End If
    Next
    
    NormalizeString = Target
  Else
    NormalizeString = strValue
    
  End If
  
End Function

Rem =========================================================================
Rem Normaliza String
Rem -------------------------------------------------------------------------
Public Function TotalNormalizeString(ByVal strValue As String) As String
    TotalNormalizeString = HTNormalizeNumber(strValue)
End Function

Rem =========================================================================
Rem Normaliza String
Rem -------------------------------------------------------------------------
Public Function NormalizeNumber(ByVal strValue As String, Optional ByVal blnFix As Boolean = False) As String
  NormalizeNumber = HTNormalizeNumber(strValue, blnFix)
End Function

Rem =========================================================================
Rem Formata strings com o padrao com digito verificador 999.999-X
Rem -------------------------------------------------------------------------
Public Function ZTIFormatDigito(ByVal strString As String) As String
  ZTIFormatDigito = FormatDigit(strString)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function FormatDigito(ByVal strString As String, Optional ByVal intDigito As Integer = 1) As String
  FormatDigito = FormatDigit(strString, intDigito)
End Function


Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function FormatDigit(ByVal strString As String, Optional ByVal intDigito As Integer = 1) As String
  
  Dim strResult As String, _
      strRest As String
      
  strResult = Left(Reverse(NormalizeNumber(strString)), intDigito) & "-"
  strRest = Reverse(CropRight(NormalizeNumber(strString), intDigito))
  
  Do While Len(strRest) > 0
    If Len(strResult) > intDigito + 1 Then strResult = strResult & "."
    strResult = strResult & Left(strRest, 3)
    If Len(strRest) < 3 Then strRest = ""
    strRest = CropLeft(strRest, 3)
  Loop
  
  FormatDigit = Reverse(strResult)
  
End Function

Rem =========================================================================
'Criado/Alterado (Ridai) p/ manter compatibilidade:
Rem
Rem -------------------------------------------------------------------------
Public Function ZTIFormatCGC(ByVal strString As String) As String
    ZTIFormatCGC = FormatCGC(strString)
End Function

Rem =========================================================================
Rem Formata strings com o padrao de CGC 999.999.999/9999-99
Rem -------------------------------------------------------------------------
Public Function FormatCGC(ByVal strString As String) As String
  FormatCGC = HTFormatCGC(strString)
End Function

Rem =========================================================================
'Criado/Alterado (Ridai)
Rem
Rem -------------------------------------------------------------------------
Public Function ZTIFormatCPF(ByVal strString As String) As String
    ZTIFormatCPF = HTFormatCPF(strString)
End Function

Rem =========================================================================
Rem Formata strings com o padrao de CPF 999.999.999-99
Rem -------------------------------------------------------------------------
Public Function FormatCPF(ByVal strString As String) As String
  FormatCPF = HTFormatCPF(strString)
End Function
Rem -------------------------------------------------------------------------
Rem Final da Function ZTIFormatCPF

Rem =========================================================================
'Criado Ueslei/Ridai
Rem
Rem -------------------------------------------------------------------------
Public Function ZTIFormatCGCCPF(ByVal strCGCCPF As String) As String
    ZTIFormatCGCCPF = FormatCGCCPF(strCGCCPF)
End Function

Rem =========================================================================
Rem Formata Valor Para CGC ou CPF
Rem -------------------------------------------------------------------------
Public Function FormatCGCCPF(ByVal strCGCCPF As String) As String
  If HTIsCGC(strCGCCPF) Then
    FormatCGCCPF = HTFormatCGC(strCGCCPF)

  ElseIf HTIsCPF(strCGCCPF) Then
    FormatCGCCPF = HTFormatCPF(strCGCCPF)

  Else
    FormatCGCCPF = strCGCCPF

  End If

End Function

Rem =========================================================================
Rem Formata strings com o padrao de telefone
Rem -------------------------------------------------------------------------
Public Function ZTIFormatFone(ByVal strString As String) As String
    ZTIFormatFone = FormatPhone(strString)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function FormatFone(ByVal strString As String) As String
  FormatFone = FormatPhone(strString)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function FormatPhone(ByVal strString As String) As String

  If Not IsNull(strString) Then
    strString = NormalizeString(strString)
    
    Select Case Len(strString)
      Case 7
        FormatPhone = Left(strString, 3) & "-" & Right(strString, 4)
        
      Case 8
        FormatPhone = Left(strString, 4) & "-" & Right(strString, 4)
        
      Case 9
        FormatPhone = "(0_" & Left(strString, 2) & ") " & Mid(strString, 4, 3) & "-" & Right(strString, 4)
  
      Case 10
        FormatPhone = "(" & Left(strString, 3) & ") " & Mid(strString, 4, 3) & "-" & Right(strString, 4)
  
      Case 11
        FormatPhone = "(" & Left(strString, 3) & ") " & Mid(strString, 4, 4) & "-" & Right(strString, 4)
        
      Case Else
        FormatPhone = strString
              
    End Select
  End If
  
End Function

Rem =========================================================================
Rem Inverte o conteudo de uma string
Rem -------------------------------------------------------------------------
Public Function ZTIReverse(ByVal strString As String) As String
    ZTIReverse = Reverse(strString)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Reverse(ByVal strString As String) As String
  Reverse = HTReverse(strString)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function DayOfWeek(ByVal dtmDate As Date, Optional ByVal intSize As Integer = 0, Optional blnFull As Boolean = False) As String
  Dim strReturn As String, intDay As Integer
  
  intDay = Weekday(dtmDate)
  Select Case intDay
    Case 1
      strReturn = "Domingo"
    Case 2
      strReturn = "Segunda"
    Case 3
      strReturn = "Terça"
    Case 4
      strReturn = "Quarta"
    Case 5
      strReturn = "Quinta"
    Case 6
      strReturn = "Sexta"
    Case 7
      strReturn = "Sábado"
  End Select
  
  If intSize > 0 Then
    strReturn = Left(strReturn, intSize)
  Else
    If blnFull And (intDay > 1 And intDay < 7) Then strReturn = strReturn & "-feira"
  End If

  DayOfWeek = strReturn

End Function

Rem =========================================================================
Rem Retorna o Nome de um Mes
Rem -------------------------------------------------------------------------
Public Function MonthName(ByVal intMes As Integer) As String
  Select Case intMes
    Case 1
      MonthName = "Janeiro"
  
    Case 2
      MonthName = "Fevereiro"
  
    Case 3
      MonthName = "Março"
  
    Case 4
      MonthName = "Abril"
  
    Case 5
      MonthName = "Maio"
  
    Case 6
      MonthName = "Junho"
  
    Case 7
      MonthName = "Julho"
  
    Case 8
      MonthName = "Agosto"
  
    Case 9
      MonthName = "Setembro"
  
    Case 10
      MonthName = "Outubro"
  
    Case 11
      MonthName = "Novembro"
  
    Case 12
      MonthName = "Dezembro"

  End Select
  
End Function

Rem =========================================================================
Rem Texto em Negrito
Rem -------------------------------------------------------------------------
Public Function BoldText(ByVal strText As String) As String
  BoldText = "<B>" & strText & "</B>"

End Function
                
Rem =========================================================================
Rem Texto em Italico
Rem -------------------------------------------------------------------------
Public Function ItalicText(ByVal strText As String) As String
  ItalicText = "<I>" & strText & "</I>"
End Function
                
Rem =========================================================================
Rem Texto em Sublinhado
Rem -------------------------------------------------------------------------
Public Function UnderlineText(ByVal strText As String) As String
  UnderlineText = "<U>" & strText & "</U>"
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Friend Function MaskNextNumber(ByVal strMaskCompleter As String) As String

  Dim blnOK As Boolean, _
      strChar As String

  blnOK = False
  Do While (Not blnOK) And (Len(sstrText3) > 0)
    If (Left(sstrText3, 1) >= "0") And (Left(sstrText3, 1) <= "9") Then
      strChar = Left(sstrText3, 1)
      blnOK = True

    End If
    sstrText3 = Right(sstrText3, Len(sstrText3) - 1)

  Loop

  If (Not sstrText3 > "") And (Not strChar > "") Then
    MaskNextNumber = strMaskCompleter

  Else
    MaskNextNumber = strChar

  End If

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Friend Function MaskNextAlfa(ByVal strMaskCompleter As String) As String

  Dim blnOK As Boolean, strChar As String

  blnOK = False
  Do While (Not blnOK) And (Len(sstrText3) > 0)
    If ((Left(sstrText3, 1) >= "a") And (Left(sstrText3, 1) <= "z")) Or ((Left(sstrText3, 1) >= "A") And (Left(sstrText3, 1) <= "Z")) Then
      strChar = Left(sstrText3, 1)
      blnOK = True

    End If
    sstrText3 = Right(sstrText3, Len(sstrText3) - 1)

  Loop

  If (Not sstrText3 > "") And (Not strChar > "") Then
    MaskNextAlfa = strMaskCompleter

  Else
    MaskNextAlfa = strChar

  End If

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Mask(ByVal strText As String, ByVal strMask As String, ByVal strMaskCompleter As String) As String

  Dim strAux As String, _
      i As Integer
  
  If Not IsNull(strText) And strText > "" Then
    If strMask > "" Then
      strAux = ""
      sstrText3 = ToString(Reverse(strText))
      For i = Len(strMask) To 1 Step -1
        Select Case Mid(strMask, i, 1)
          Case "0"
            strAux = strAux & MaskNextNumber(strMaskCompleter)

          Case "A"
            strAux = strAux & MaskNextAlfa(strMaskCompleter)

          Case Else
            strAux = strAux & Mid(strMask, i, 1)

        End Select
      Next
      Mask = Reverse(strAux)

    Else
      Mask = strText

    End If
  Else
    Mask = ""

  End If

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function ToString(ByVal vntText As Variant) As String
  
  If Not IsNull(vntText) And vntText > "" Then
    ToString = CStr(vntText)
  Else
    ToString = ""
  End If

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function MonthSize(ByVal intMonth As Byte, Optional ByVal intYear As Integer = -1) As Byte
  If intYear = -1 Then intYear = Year(Now())
  Select Case intMonth
    Case 1, 3, 5, 7, 8, 10, 12
      MonthSize = 31

    Case 4, 6, 9, 11
      MonthSize = 30

    Case 2
      If intYear Mod 4 = 0 Then
        MonthSize = 29
      
      Else
        MonthSize = 28
        
      End If
                                                
  End Select
  
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function BreakString(ByVal strValue As String, ByVal intSize As Integer) As String
  Dim strText As String
  
  strText = Trim(strValue)
  
  If Len(strText) > intSize And intSize > 4 Then
    BreakString = Left(strText, intSize - 3) & "..."
  Else
    BreakString = strText
  End If

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function NormalizeAccents(ByVal strString As String) As String
  strString = Replace(strString, "á", "a")
  strString = Replace(strString, "ä", "a")
  strString = Replace(strString, "â", "a")
  strString = Replace(strString, "à", "a")
  strString = Replace(strString, "ã", "a")
  strString = Replace(strString, "é", "e")
  strString = Replace(strString, "ë", "e")
  strString = Replace(strString, "ê", "e")
  strString = Replace(strString, "è", "e")
  strString = Replace(strString, "í", "i")
  strString = Replace(strString, "ï", "i")
  strString = Replace(strString, "î", "i")
  strString = Replace(strString, "ì", "i")
  strString = Replace(strString, "ó", "o")
  strString = Replace(strString, "ö", "o")
  strString = Replace(strString, "ô", "o")
  strString = Replace(strString, "ò", "o")
  strString = Replace(strString, "õ", "o")
  strString = Replace(strString, "ú", "u")
  strString = Replace(strString, "ü", "u")
  strString = Replace(strString, "û", "u")
  strString = Replace(strString, "ù", "u")
  strString = Replace(strString, "ç", "c")
  strString = Replace(strString, "ñ", "n")
  strString = Replace(strString, "ý", "y")
  strString = Replace(strString, "ÿ", "y")
 
  strString = Replace(strString, "Á", "A")
  strString = Replace(strString, "Ä", "A")
  strString = Replace(strString, "Â", "A")
  strString = Replace(strString, "À", "A")
  strString = Replace(strString, "Ã", "A")
  strString = Replace(strString, "É", "E")
  strString = Replace(strString, "Ë", "E")
  strString = Replace(strString, "Ê", "E")
  strString = Replace(strString, "È", "E")
  strString = Replace(strString, "Í", "I")
  strString = Replace(strString, "Ï", "I")
  strString = Replace(strString, "Î", "I")
  strString = Replace(strString, "Ì", "I")
  strString = Replace(strString, "Ó", "O")
  strString = Replace(strString, "Ö", "O")
  strString = Replace(strString, "Ô", "O")
  strString = Replace(strString, "Ò", "O")
  strString = Replace(strString, "Õ", "O")
  strString = Replace(strString, "Ú", "U")
  strString = Replace(strString, "Ü", "U")
  strString = Replace(strString, "Û", "U")
  strString = Replace(strString, "Ù", "U")
  strString = Replace(strString, "Ç", "C")
  strString = Replace(strString, "Ñ", "N")
  strString = Replace(strString, "Ý", "Y")

  NormalizeAccents = strString

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function ChangeChar(ByVal strString As String) As String
  ChangeChar = NormalizeAccents(strString)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function CropRight(ByVal strString As String, ByVal intChar As Integer) As String
  If strString > "" Then
    CropRight = Left(strString, Len(strString) - intChar)
  
  Else
    CropRight = strString
  
  End If

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function CropLeft(ByVal strString As String, ByVal intChar As Integer) As String
    
  CropLeft = strString
  
  If strString > "" Then
    CropLeft = Right(strString, Len(strString) - intChar)
    
  End If
  
End Function

Rem =========================================================================
Rem Função que retorna quantas vezes um caracter (strDiv) se repete em
Rem uma string (strString)
Rem -------------------------------------------------------------------------
Public Function CountSubString(ByVal strString As String, ByVal strSubString As String) As Long
    
    Dim intPos As Long
    
    CountSubString = 0
    
    If Not strSubString > "" Then strString = ""
    
    Do While strString > ""
        intPos = InStr(strString, strSubString)
        If intPos = 0 Then
            intPos = Len(strString)
        Else
            CountSubString = CountSubString + 1
        End If
        strString = Mid(strString, intPos + 1)
    Loop
    
End Function

'Criado (Ridai Govinda)
Rem =========================================================================
Rem Função clone do printf() do C/C++ e do Format do Delphi
Rem Passe uma string com $'s um paramarray com um monte de strings que serão
Rem preenchidas substituindo os $'s
Rem -------------------------------------------------------------------------
Public Function FormatText(ByVal strMask As String, ParamArray arrStrings() As Variant) As String
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
  
  FormatText = strResult

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
    strResult = Mid(CStr(vntValue), 1, IIf(intWidth > 0, intWidth, Len(vntValue)))
  
  Case lconTypeHexa
    strResult = Mid(Hex(CStr(vntValue)), 1, IIf(intWidth > 0, intWidth, Len(vntValue)))
  
  End Select
  
  If blnLeftJustify Then
    If intWidth > Len(strResult) Then
      strResult = strResult & Space(intWidth - Len(strResult))
    End If
  End If
  
  GetFormatedString = strResult

End Function
'----------------------------------------------------------------------------

Rem Convert Numbers and Currency Values to English
Public Function NumToString(ByVal nNumber As Currency) As String

Dim bNegative As Boolean
Dim bHundred As Boolean

If nNumber < 0 Then
    bNegative = True
End If

nNumber = Abs(Int(nNumber))

If nNumber < 1000 Then
    If nNumber \ 100 > 0 Then
        NumToString = NumToString & _
             NumToString(nNumber \ 100) & " hundred"
        bHundred = True
    End If
    nNumber = nNumber - ((nNumber \ 100) * 100)
    Dim bNoFirstDigit As Boolean
    bNoFirstDigit = False
    Select Case nNumber \ 10
        Case 0
            Select Case nNumber Mod 10
                Case 0
                    If Not bHundred Then
                        NumToString = NumToString & " zero"
                    End If
                Case 1: NumToString = NumToString & " one"
                Case 2: NumToString = NumToString & " two"
                Case 3: NumToString = NumToString & " three"
                Case 4: NumToString = NumToString & " four"
                Case 5: NumToString = NumToString & " five"
                Case 6: NumToString = NumToString & " six"
                Case 7: NumToString = NumToString & " seven"
                Case 8: NumToString = NumToString & " eight"
                Case 9: NumToString = NumToString & " nine"
            End Select
            bNoFirstDigit = True
        Case 1
            Select Case nNumber Mod 10
                Case 0: NumToString = NumToString & " ten"
                Case 1: NumToString = NumToString & " eleven"
                Case 2: NumToString = NumToString & " twelve"
                Case 3: NumToString = NumToString & " thirteen"
                Case 4: NumToString = NumToString & " fourteen"
                Case 5: NumToString = NumToString & " fifteen"
                Case 6: NumToString = NumToString & " sixteen"
                Case 7: NumToString = NumToString & " seventeen"
                Case 8: NumToString = NumToString & " eighteen"
                Case 9: NumToString = NumToString & " nineteen"
            End Select
            bNoFirstDigit = True
        Case 2: NumToString = NumToString & " twenty"
        Case 3: NumToString = NumToString & " thirty"
        Case 4: NumToString = NumToString & " forty"
        Case 5: NumToString = NumToString & " fifty"
        Case 6: NumToString = NumToString & " sixty"
        Case 7: NumToString = NumToString & " seventy"
        Case 8: NumToString = NumToString & " eighty"
        Case 9: NumToString = NumToString & " ninety"
    End Select
    If Not bNoFirstDigit Then
        If nNumber Mod 10 <> 0 Then
            NumToString = NumToString & "-" & _
                          Mid(NumToString(nNumber Mod 10), 2)
        End If
    End If
Else
    Dim nTemp As Currency
    nTemp = 10 ^ 12 'trillion
    Do While nTemp >= 1
        If nNumber >= nTemp Then
            NumToString = NumToString & _
                          NumToString(Int(nNumber / nTemp))
            Select Case Int(Log(nTemp) / Log(10) + 0.5)
                Case 12: NumToString = NumToString & " trillion"
                Case 9: NumToString = NumToString & " billion"
                Case 6: NumToString = NumToString & " million"
                Case 3: NumToString = NumToString & " thousand"
            End Select
           
            nNumber = nNumber - (Int(nNumber / nTemp) * nTemp)
        End If
        nTemp = nTemp / 1000
    Loop
End If

If bNegative Then
    NumToString = " negative" & NumToString
End If
    
End Function

Public Function DollarToString(ByVal nAmount As Currency) As _
String

    Dim nDollar As Currency
    Dim nCent As Currency
    
    nDollar = Int(nAmount)
    nCent = (Abs(nAmount) * 100) Mod 100
    
    DollarToString = NumToString(nDollar) & " dollar"
    
    If Abs(nDollar) <> 1 Then
        DollarToString = DollarToString & "s"
    End If
    
    DollarToString = DollarToString & " and" & _
                     NumToString(nCent) & " cent"
                     
    If Abs(nCent) <> 1 Then
        DollarToString = DollarToString & "s"
    End If
    
End Function


' * Comments  : Convert an integer number to text in French

Public Function ConvertirEnText(ValNum As Double) As String

Static Unites(0 To 9) As String
Static Dixaines(0 To 9) As String
Static LesDixaines(0 To 9) As String
Static Milliers(0 To 4) As String

Dim i As Integer
Dim nPosition As Integer
Dim ValNb As Integer
Dim LesZeros As Integer
Dim strResultat As String
Dim strTemp As String
Dim tmpBuff As String

Unites(0) = "zero"
Unites(1) = "un"
Unites(2) = "deux"
Unites(3) = "trois"
Unites(4) = "quatre"
Unites(5) = "cinq"
Unites(6) = "six"
Unites(7) = "sept"
Unites(8) = "huit"
Unites(9) = "neuf"

Dixaines(0) = "dix"
Dixaines(1) = "onze"
Dixaines(2) = "douze"
Dixaines(3) = "treize"
Dixaines(4) = "quatorze"
Dixaines(5) = "quinze"
Dixaines(6) = "seize"
Dixaines(7) = "dix-sept"
Dixaines(8) = "dix-huit"
Dixaines(9) = "dix-neuf"

LesDixaines(0) = ""
LesDixaines(1) = "dix"
LesDixaines(2) = "vingt"
LesDixaines(3) = "trente"
LesDixaines(4) = "quarante"
LesDixaines(5) = "cinquante"
LesDixaines(6) = "soixante"
LesDixaines(7) = "soixante-dix"
LesDixaines(8) = "quatre-vingt"
LesDixaines(9) = "quatre-vingt-dix"

Milliers(0) = ""
Milliers(1) = "mille"
Milliers(2) = "million"
Milliers(3) = "millard"
Milliers(4) = "mille"

On Error GoTo NbVersTexteError

strTemp = CStr(Int(ValNum))

For i = Len(strTemp) To 1 Step -1
   ValNb = Val(Mid$(strTemp, i, 1))
   nPosition = (Len(strTemp) - i) + 1
   Select Case (nPosition Mod 3)
      Case 1
         LesZeros = False
         If i = 1 Then
            If ValNb > 1 Then
               tmpBuff = Unites(ValNb) & " "
            Else
               tmpBuff = ""
            End If
         ElseIf Mid$(strTemp, i - 1, 1) = "1" Then
            tmpBuff = Dixaines(ValNb) & " "
            i = i - 1
         ElseIf Mid$(strTemp, i - 1, 1) = "9" Then
            tmpBuff = LesDixaines(8) & " " & _
                Dixaines(ValNb) & " "
            i = i - 1
         ElseIf Mid$(strTemp, i - 1, 1) = "7" Then
            tmpBuff = LesDixaines(6) & " " & _
                 Dixaines(ValNb) & " "
            i = i - 1
         ElseIf ValNb > 0 Then
            tmpBuff = Unites(ValNb) & " "
         Else
            LesZeros = True
            If i > 1 Then
               If Mid$(strTemp, i - 1, 1) <> "0" Then
                  LesZeros = False
               End If
            End If
            If i > 2 Then
               If Mid$(strTemp, i - 2, 1) <> "0" Then
                  LesZeros = False
               End If
            End If
            tmpBuff = ""
         End If
         If LesZeros = False And nPosition > 1 Then
            tmpBuff = tmpBuff & Milliers(nPosition / 3) & " "
         End If
         strResultat = tmpBuff & strResultat
      Case 2
         If ValNb > 0 Then
            strResultat = LesDixaines(ValNb) & " " & _
               strResultat
         End If
      Case 0
         If ValNb > 0 Then
            If ValNb > 1 Then
               strResultat = Unites(ValNb) & " cent " & _
                  strResultat
            Else
               strResultat = "cent " & strResultat
            End If
         End If
   End Select
Next i
If Len(strResultat) > 0 Then
   strResultat = UCase$(Left$(strResultat, 1)) & _
      Mid$(strResultat, 2)
End If

EndNbVersTexte:
ConvertirEnText = strResultat
Exit Function

NbVersTexteError:
strResultat = "Une Erreur !"
Resume EndNbVersTexte
End Function

Rem Increment a Text String (from ABY to ABZ to AZA etc.)
Function IncrementTextString(txt As String) As String
    Dim l As Integer, i As Integer, c As Integer
    Dim s As String
    s = txt
    l = Len(s)
    
    For i = l To 1 Step -1    'go thru the string, right to left
        c = Asc(Mid(s, i, 1)) 'ASCII code of the i-th character
        Select Case c
            Case 65 To 89, 97 To 121  'A-Y or a-y
                s = Left(s, i - 1) & Chr(c + 1) & Mid(s, i + 1)
                Exit For
            Case 90    'Z
                s = Left(s, i - 1) & "A" & Mid(s, i + 1)
            Case 122   'z
                s = Left(s, i - 1) & "a" & Mid(s, i + 1)
        End Select
        'in the last two cases, we need to continue the loop:
    Next i
    If i = 0 Then
        IncrementTextString = String(l + 1, 65) 'grow the string
    Else
        IncrementTextString = s
    End If
End Function

Public Function PasswordGenerator(ByVal lngLength As Long) _
  As String

' Description: Generate a random password of 'user input' length
' Parameters : lngLength - the length of the password to be
                'generated
' Returns    : String    - Randomly generated password
' Created    : 08/21/1999 Andrew Ells-O'Brien
             '(andrew@ellsobrien@msn.com)
  
On Error GoTo Err_Proc
  
 Dim iChr As Integer
 Dim c As Long
 Dim strResult As String
 Dim iAsc As String
 
 Randomize Timer

 For c = 1 To lngLength
   
   ' Randomly decide what set of ASCII chars we will use
   iAsc = Int(3 * Rnd + 1)
   
    'Randomly pick a char from the random set
   Select Case iAsc
     Case 1
       iChr = Int((Asc("Z") - Asc("A") + 1) * Rnd + Asc("A"))
     Case 2
       iChr = Int((Asc("z") - Asc("a") + 1) * Rnd + Asc("a"))
     Case 3
       iChr = Int((Asc("9") - Asc("0") + 1) * Rnd + Asc("0"))
     Case Else
       Err.Raise 20000, , "PasswordGenerator has a problem."
   End Select
   
   strResult = strResult & Chr(iChr)
 
 Next c
 
 PasswordGenerator = strResult
 
Exit_Proc:
 Exit Function
 
Err_Proc:
 MsgBox Err.number & ": " & Err.Description, _
    vbOKOnly + vbCritical
 PasswordGenerator = vbNullString
 Resume Exit_Proc
 
End Function

Rem Read Tokens using a Delimiter
Public Function ParameterValue(ParseCharacter As String, _
                               tString As Variant, _
                               Index As Integer) As String

'USAGE: Debug.Print ParameterValue(",", "A,B,C,D,E", 4)
'Prints "D" to the Debug Window
  Dim CurrentPosition As Integer
  Dim ParseToPosition As Integer
  Dim CurrentToken As Integer
  Dim TempString As String
  TempString = Trim(tString) + ParseCharacter
  If Len(TempString) = 1 Then Exit Function
  CurrentPosition = 1
  CurrentToken = 1
  Do
    ParseToPosition = InStr(CurrentPosition, TempString, _
                            ParseCharacter)
    If Index = CurrentToken Then
      ParameterValue = Mid$(TempString, CurrentPosition, _
                            ParseToPosition - CurrentPosition)
      Exit Function
    End If
    CurrentToken = CurrentToken + 1
    CurrentPosition = ParseToPosition + 1
  Loop Until (CurrentPosition >= Len(TempString))
End Function

Public Function ParameterCount(ParseCharacter As String, _
                               tString As Variant) As Integer
            
'USAGE: ParameterCount(",", "A,B,C,D,E")
'Prints "5" to the Debug Window
  Dim CurrentPosition As Integer
  Dim ParseToPosition As Integer
  Dim CurrentToken As Integer
  Dim TempString As String
  TempString = Trim(tString) + ParseCharacter
  If Len(TempString) = 1 Then Exit Function
  CurrentPosition = 1
  CurrentToken = 1
  Do
      ParseToPosition = InStr(CurrentPosition, TempString, _
                              ParseCharacter)
      CurrentToken = CurrentToken + 1
      CurrentPosition = ParseToPosition + 1
  Loop Until (CurrentPosition >= Len(TempString))
  ParameterCount = CurrentToken - 1
End Function

Rem Remove All Spaces from a String
Public Function RemoveAllSpaces(ByVal InputString As String) _
As String

'NOTE: Intended for VB5 and below; if you have VB6, use built-in
'Replace Function. (e.g., MyString = Replace(MyString, " ", "")

Dim sAns As String
Dim lLen As String
Dim lCtr As Long, lCtr2 As Long
Dim sChar As String


lLen = Len(InputString)
sAns = InputString
lCtr2 = 1

For lCtr = 1 To lLen
    sChar = Mid(InputString, lCtr, 1)
    If sChar <> " " Then
        Mid(sAns, lCtr2, 1) = sChar
        lCtr2 = lCtr2 + 1
    End If
Next

If lCtr2 > 1 Then
    sAns = Left(sAns, lCtr2 - 1)
Else
    sAns = ""
End If

RemoveAllSpaces = sAns
End Function

Rem Replace Special Vowels (à, é, è, ...) with "Regular" (7-bit ASCII) Values
Public Function ReplaceAccent(sString As String) As String
    Dim i As Integer
    
' Replace A & a
    For i = 192 To 197
        sString = Replace(sString, Chr(i), "A")
    Next
    For i = 224 To 229
        sString = Replace(sString, Chr(i), "a")
    Next
' Replace C & c
    sString = Replace(sString, Chr(199), "C")
    sString = Replace(sString, Chr(231), "c")
' Replace E & e
    For i = 200 To 203
        sString = Replace(sString, Chr(i), "E")
    Next
    For i = 232 To 235
        sString = Replace(sString, Chr(i), "e")
    Next
' Replace I & i
    For i = 204 To 207
        sString = Replace(sString, Chr(i), "I")
    Next
    For i = 236 To 239
        sString = Replace(sString, Chr(i), "i")
    Next
' Replace D
    sString = Replace(sString, Chr(208), "D")
' Replace N & n
    sString = Replace(sString, Chr(209), "N")
    sString = Replace(sString, Chr(241), "n")
' Replace O & o
    For i = 210 To 214
        sString = Replace(sString, Chr(i), "O")
    Next
    For i = 242 To 246
        sString = Replace(sString, Chr(i), "o")
    Next
' Replace U & u
    For i = 217 To 220
        sString = Replace(sString, Chr(i), "U")
    Next
    For i = 249 To 252
        sString = Replace(sString, Chr(i), "u")
    Next
' Replace Y & y
    sString = Replace(sString, Chr(221), "Y")
    sString = Replace(sString, Chr(253), "y")
    sString = Replace(sString, Chr(255), "y")
    
    ReplaceAccent = sString
End Function
Public Function UniqueChars(ByVal OrigString As String) _
  As String

'INPUT: ANY STRING
'OUTPUT: The Unique Characters of the string, in the order
'        They First appear
'Example: Debug.Print(UniqueChars("FreeVBCode")):
'         Outputs: FreVBCod

Dim oCol As New Collection
Dim sAns As String
Dim lCtr As Long, lCount As Long
Dim sChar As String

lCount = Len(OrigString)

For lCtr = 1 To lCount
    sChar = Mid(OrigString, lCtr, 1)
    On Error Resume Next
    oCol.Add sChar, sChar
    If Err.number = 0 Then sAns = sAns & sChar
    Err.Clear
Next

UniqueChars = sAns

End Function

Rem Return and Count Unique Characters from a String
Public Function UniqueCharCount(ByVal TheString As String) _
   As Long
    UniqueCharCount = Len(UniqueChars(TheString))
End Function

' Return a paded string with blancs on right
Public Function PadR(ByVal sTmp As String, lNbr As Long) _
  As String
   
   PadR = Left(RTrim(sTmp) & Space(lNbr), lNbr)

End Function

' Return a padded string with blanks on left

Public Function PadL(ByVal sTmp As String, lNbr As Long) _
   As String
   

   PadL = Right(Space(lNbr) & RTrim(sTmp), lNbr)

End Function


Public Function OnlyAlphaNumericChars(ByVal OrigString As _
  String) As String
'***********************************************************
'INPUT:  Any String
'OUTPUT: The Input String with all non-alphanumeric characters
'        removed
'EXAMPLE Debug.Print OnlyAlphaNumericChars("Hello World!")
         'output = "HelloWorld")
'NOTES:  Not optimized for speed and will run slow on long
'        strings.  If you plan on using long strings, consider
'        using alternative method of appending to output string,
'        such as the method at
'        http://www.freevbcode.com/ShowCode.Asp?ID=154
'***********************************************************
    Dim lLen As Long
    Dim sAns As String
    Dim lCtr As Long
    Dim sChar As String
    
    OrigString = Trim(OrigString)
    lLen = Len(OrigString)
    For lCtr = 1 To lLen
        sChar = Mid(OrigString, lCtr, 1)
        If IsAlphaNumeric(Mid(OrigString, lCtr, 1)) Then
            sAns = sAns & sChar
        End If
    DoEvents '(optional, but if processing long string,
    'necessary to prevent program from appearing to hang)
    'if used, write your app so no re-entrancy into this function
    'can occur)
    Next
        
    OnlyAlphaNumericChars = sAns

End Function

Private Function IsAlphaNumeric(sChr As String) As Boolean
    IsAlphaNumeric = sChr Like "[0-9A-Za-z]"
End Function

Rem Strip special characters from a string
Public Function StripSpecialCharacters(ByVal sIn As String) _
As String

Dim sWkg As String, sOrigString As String, sNewString As String

Dim lLen As Long
Dim lCtr As Long, lCtr2 As Long
Dim sChar As String

lLen = Len(sIn)
'create buffer

sOrigString = Space$(lLen)
sOrigString = sIn
sNewString = sOrigString
lCtr2 = 1

For lCtr = 1 To lLen
    sChar = Mid(sOrigString, lCtr, 1)
    If Asc(sChar) < 128 Then
        Mid(sNewString, lCtr2, 1) = sChar
        lCtr2 = lCtr2 + 1
    End If
Next

If lCtr2 > 1 Then
    sNewString = Left(sNewString, lCtr2 - 1)
Else
    sNewString = ""
End If

StripSpecialCharacters = sNewString
End Function

'Designed for newbies to cryptology.
'Because XOR's inverse is itself, you use this function to
'encrypt and decrypt.
' e.g. Cipher_text = transform(Plain_text,"123")
'      Plain_text = transform(Cipher_text,"123")

Public Function StringXOR(ByVal text As String, _
   ByVal Key As String) As String
   Dim i, a
For i = 1 To Len(text)
      a = i Mod Len(Key): If a = 0 Then a = Len(Key)
      StringXOR = StringXOR & _
         Chr(Asc(Mid(Key, a, 1)) Xor Asc(Mid(text, i, 1)))
Next i

End Function

Public Function IsTopDomainName(ByVal strDomainCode As String) As Boolean
  
  If GetTopDomainName(strDomainCode) > "" Then
    IsTopDomainName = True
  
  Else
    IsTopDomainName = False
    
  End If
End Function

Public Function GetTopDomainName(ByVal strDomainCode As String) As String
  Dim strResult

  Select Case LCase(strDomainCode)
  Case "ac"
     strResult = "Ascension Island"
     
    Case "ad"
      strResult = "Andorra"


    Case "ae"
      strResult = "United Arab Emirates"


    Case "af"
      strResult = "Afghanistan"

    Case "ag"
      strResult = "Antigua and Barbuda"

    Case "ai"
      strResult = "Anguilla"

    Case "al"
      strResult = "Albania"

    Case "am"
      strResult = "Armenia"

    Case "an"
      strResult = "Netherlands Antilles"

    Case "ao"
      strResult = "Angola"

    Case "aq"
      strResult = "Antartica"

    Case "ar"
      strResult = "Argentina"

    Case "as"
      strResult = "American Samoa"

    Case "at"
      strResult = "Austria"

    Case "au"
      strResult = "Australia"

    Case "aw"
      strResult = "Aruba"

    Case "az"
      strResult = "Azerbaijan"

    Case "ba"
      strResult = "Bosnia and Herzegovina"

    Case "bb"
      strResult = "Barbados"

    Case "bd"
      strResult = "Bangladesh"

    Case "be"
      strResult = "Belgium"

    Case "bf"
      strResult = "Burkina Faso"

    Case "bg"
      strResult = "Bulgaria"

    Case "bh"
      strResult = "Bahrain"

    Case "bi"
      strResult = "Burundi"

    Case "bj"
      strResult = "Benin"

    Case "bm"
      strResult = "Bermuda"

    Case "bn"
      strResult = "Brunei Darussalam"

    Case "bo"
      strResult = "Bolivia"

    Case "br"
      strResult = "Brazil"

    Case "bs"
      strResult = "Bahamas"

    Case "bt"
      strResult = "Bhutan"

    Case "bv"
      strResult = "Bouvet Island"

    Case "bw"
      strResult = "Botswana"

    Case "by"
      strResult = "Belarus"

    Case "bz"
      strResult = "Belize"

    Case "ca"
      strResult = "Canada"

    Case "cc"
      strResult = "Cocos (Keeling) Islands"

    Case "cd"
      strResult = "Congo, Democratic People's Republic"

    Case "cf"
      strResult = "Central African Republic"

    Case "cg"
      strResult = "Congo, Republic of"

    Case "ch"
      strResult = "Switzerland"

    Case "ci"
      strResult = "Cote d'Ivoire"

    Case "ck"
      strResult = "Cook Islands"

    Case "cl"
      strResult = "Chile"

    Case "cm"
      strResult = "Cameroon"

    Case "cn"
      strResult = "China"

    Case "co"
      strResult = "Colombia"

    Case "cr"
      strResult = "Costa Rica"

    Case "cu"
      strResult = "Cuba"

    Case "cv"
      strResult = "Cap Verde"

    Case "cx"
      strResult = "Christmas Island"

    Case "cy"
      strResult = "Cyprus"

    Case "cz"
      strResult = "Czech Republic"

    Case "de"
      strResult = "Germany"

    Case "dj"
      strResult = "Djibouti"

    Case "dk"
      strResult = "Denmark"

    Case "dm"
      strResult = "Dominica"

    Case "do"
      strResult = "Dominican Republic"

    Case "dz"
      strResult = "Algeria"

    Case "ec"
      strResult = "Ecuador"

    Case "ee"
      strResult = "Estonia"

    Case "eg"
      strResult = "Egypt"

    Case "eh"
      strResult = "Western Sahara"

    Case "er"
      strResult = "Eritrea"

    Case "es"
      strResult = "Spain"

    Case "et"
      strResult = "Ethiopia"

    Case "fi"
      strResult = "Finland"

    Case "fj"
      strResult = "Fiji"

    Case "fk"
      strResult = "Falkland Islands (Malvina)"

    Case "fm"
      strResult = "Micronesia, Federal State of"

    Case "fo"
      strResult = "Faroe Islands"

    Case "fr"
      strResult = "France"

    Case "ga"
      strResult = "Gabon"

    Case "gd"
      strResult = "Grenada"

    Case "ge"
      strResult = "Georgia"

    Case "gf"
      strResult = "French Guiana"

    Case "gg"
      strResult = "Guernsey"

    Case "gh"
      strResult = "Ghana"

    Case "gi"
      strResult = "Gibraltar"

    Case "gl"
      strResult = "Greenland"

    Case "gm"
      strResult = "Gambia"

    Case "gn"
      strResult = "Guinea"

    Case "gp"
      strResult = "Guadeloupe"

    Case "gq"
      strResult = "Equatorial Guinea"

    Case "gr"
      strResult = "Greece"

    Case "gs"
      strResult = "South Georgia and the South Sandwich Islands"

    Case "gt"
      strResult = "Guatemala"

    Case "gu"
      strResult = "Guam"

    Case "gw"
      strResult = "Guinea-Bissau"

    Case "gy"
      strResult = "Guyana"

    Case "hk"
      strResult = "Hong Kong"

    Case "hm"
      strResult = "Heard and McDonald Islands"

    Case "hn"
      strResult = "Honduras"

    Case "hr"
      strResult = "Croatia/Hrvatska"

    Case "ht"
      strResult = "Haiti"

    Case "hu"
      strResult = "Hungary"

    Case "id"
      strResult = "Indonesia"

    Case "ie"
      strResult = "Ireland"

    Case "il"
      strResult = "Israel"

    Case "im"
      strResult = "Isle of Man"

    Case "in"
      strResult = "India"

    Case "io"
      strResult = "British Indian Ocean Territory"

    Case "iq"
      strResult = "Iraq"

    Case "ir"
      strResult = "Iran (Islamic Republic of)"

    Case "is"
      strResult = "Iceland"

    Case "it"
      strResult = "Italy"

    Case "je"
      strResult = "Jersey"

    Case "jm"
      strResult = "Jamaica"

    Case "jo"
      strResult = "Jordan"

    Case "jp"
      strResult = "Japan"

    Case "ke"
      strResult = "Kenya"

    Case "kg"
      strResult = "Kyrgyzstan"

    Case "kh"
      strResult = "Cambodia"

    Case "ki"
      strResult = "Kiribati"

    Case "km"
      strResult = "Comoros"

    Case "kn"
      strResult = "Saint Kitts and Nevis"

    Case "kp"
      strResult = "Korea, Democratic People's Republic"

    Case "kr"
      strResult = "Korea, Republic of"

    Case "kw"
      strResult = "Kuwait"

    Case "ky"
      strResult = "Cayman Islands"

    Case "kz"
      strResult = "Kazakhstan"

    Case "la"
      strResult = "Lao People's Democratic Republic"

    Case "lb"
      strResult = "Lebanon"

    Case "lc"
      strResult = "Saint Lucia"

    Case "li"
      strResult = "Liechtenstein"

    Case "lk"
      strResult = "Sri Lanka"

    Case "lr"
      strResult = "Liberia"

    Case "ls"
      strResult = "Lesotho"

    Case "lt"
      strResult = "Lithuania"

    Case "lu"
      strResult = "Luxembourg"

    Case "lv"
      strResult = "Latvia"

    Case "ly"
      strResult = "Libyan Arab Jamahiriya"

    Case "ma"
      strResult = "Morocco"

    Case "mc"
      strResult = "Monaco"

    Case "md"
      strResult = "Moldova, Republic of"

    Case "mg"
      strResult = "Madagascar"

    Case "mh"
      strResult = "Marshall Islands"

    Case "mk"
      strResult = "Macedonia, Former Yugoslav Republic"

    Case "ml"
      strResult = "Mali"

    Case "mm"
      strResult = "Myanmar"

    Case "mn"
      strResult = "Mongolia"

    Case "mo"
      strResult = "Macau"

    Case "mp"
      strResult = "Northern Mariana Islands"

    Case "mq"
      strResult = "Martinique"

    Case "mr"
      strResult = "Mauritania"

    Case "ms"
      strResult = "Montserrat"

    Case "mt"
      strResult = "Malta"

    Case "mu"
      strResult = "Mauritius"

    Case "mv"
      strResult = "Maldives"

    Case "mw"
      strResult = "Malawi"

    Case "mx"
      strResult = "Mexico"

    Case "my"
      strResult = "Malaysia"

    Case "mz"
      strResult = "Mozambique"

    Case "na"
      strResult = "Namibia"

    Case "nc"
      strResult = "New Caledonia"

    Case "ne"
      strResult = "Niger"

    Case "nf"
      strResult = "Norfolk Island"

    Case "ng"
      strResult = "Nigeria"

    Case "ni"
      strResult = "Nicaragua"

    Case "nl"
      strResult = "Netherlands"

    Case "no"
      strResult = "Norway"

    Case "np"
      strResult = "Nepal"

    Case "nr"
      strResult = "Nauru"

    Case "nu"
      strResult = "Niue"

    Case "nz"
      strResult = "New Zealand"

    Case "om"
      strResult = "Oman"

    Case "pa"
      strResult = "Panama"

    Case "pe"
      strResult = "Peru"

    Case "pf"
      strResult = "French Polynesia"

    Case "pg"
      strResult = "Papua New Guinea"

    Case "ph"
      strResult = "Philippines"

    Case "pk"
      strResult = "Pakistan"

    Case "pl"
      strResult = "Poland"

    Case "pm"
      strResult = "St. Pierre and Miquelon"

    Case "pn"
      strResult = "Pitcairn Island"

    Case "pr"
      strResult = "Puerto Rico"

    Case "ps"
      strResult = "Palestinian Territories"

    Case "pt"
      strResult = "Portugal"

    Case "pw"
      strResult = "Palau"

    Case "py"
      strResult = "Paraguay"

    Case "qa"
      strResult = "Qatar"

    Case "re"
      strResult = "Reunion Island"

    Case "ro"
      strResult = "Romania"

    Case "ru"
      strResult = "Russian Federation"

    Case "rw"
      strResult = "Rwanda"

    Case "sa"
      strResult = "Saudi Arabia"

    Case "sb"
      strResult = "Solomon Islands"

    Case "sc"
      strResult = "Seychelles"

    Case "sd"
      strResult = "Sudan"

    Case "se"
      strResult = "Sweden"

    Case "sg"
      strResult = "Singapore"

    Case "sh"
      strResult = "St. Helena"

    Case "si"
      strResult = "Slovenia"

    Case "sj"
      strResult = "Svalbard and Jan Mayen Islands"

    Case "sk"
      strResult = "Slovak Republic"

    Case "sl"
      strResult = "Sierra Leone"

    Case "sm"
      strResult = "San Marino"

    Case "sn"
      strResult = "Senegal"

    Case "so"
      strResult = "Somalia"

    Case "sr"
      strResult = "Suriname"

    Case "st"
      strResult = "Sao Tome and Principe"

    Case "sv"
      strResult = "El Salvador"

    Case "sy"
      strResult = "Syrian Arab Republic"

    Case "sz"
      strResult = "Swaziland"

    Case "tc"
      strResult = "Turks and Ciacos Islands"

    Case "td"
      strResult = "Chad"

    Case "tf"
      strResult = "French Southern Territories"

    Case "tg"
      strResult = "Togo"

    Case "th"
      strResult = "Thailand"

    Case "tj"
      strResult = "Tajikistan"

    Case "tk"
      strResult = "Tokelau"

    Case "tm"
      strResult = "Turkmenistan"

    Case "tn"
      strResult = "Tunisia"

    Case "to"
      strResult = "Tonga"

    Case "tp"
      strResult = "East Timor"

    Case "tr"
      strResult = "Turkey"

    Case "tt"
      strResult = "Trinidad and Tobago"

    Case "tv"
      strResult = "Tuvalu"

    Case "tw"
      strResult = "Taiwan"

    Case "tz"
      strResult = "Tanzania"

    Case "ua"
      strResult = "Ukraine"

    Case "ug"
      strResult = "Uganda"

    Case "uk"
      strResult = "United Kingdom"

    Case "um"
      strResult = "US Minor Outlying Islands"

    Case "us"
      strResult = "United States"

    Case "uy"
      strResult = "Uruguay"

    Case "uz"
      strResult = "Uzbekistan"

    Case "va"
      strResult = "Holy See (City Vatican State)"

    Case "vc"
      strResult = "Saint Vincent and the Grenadines"

    Case "ve"
      strResult = "Venezuela"

    Case "vg"
      strResult = "Virgin Islands (British)"

    Case "vi"
      strResult = "Virgin Islands (USA)"

    Case "vn"
      strResult = "Vietnam"

    Case "vu"
      strResult = "Vanuatu"

    Case "wf"
      strResult = "Wallis and Futuna Islands"

    Case "ws"
      strResult = "Western Samoa"

    Case "ye"
      strResult = "Yemen"

    Case "yt"
      strResult = "Mayotte"

    Case "yu"
      strResult = "Yugoslavia"

    Case "za"
      strResult = "South Africa"

    Case "zm"
      strResult = "Zambia"

    Case "zr"
      strResult = "Zaire"

    Case "zw"
      strResult = "Zimbabwe"
      
    Case Else
      strResult = ""
  End Select
  
  GetTopDomainName = strResult
  
End Function

