VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Character"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =========================================================================
' Character.cls
' -------------------------------------------------------------------------
' Descricao   : Include de Caracteres
' Criaçao     : 11:23h 7/2/1998
' Local       : Brasilia/DF
' Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 1997, 1998, 1999 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------

Option Explicit
 
Private ScriptContext As ScriptingContext

Private sstrComments As String, _
        sstrTitle As String, _
        sstrVersion As String, _
        sstrCompanyName As String, _
        sstrProductName As String, _
        sstrLegalCopyright As String

Private Request As Request, _
        Response As Response, _
        Server As Server, _
        Session As Session, _
        Application As Application

Public FontDirectory As String
Public Font As String

Public Sub OnStartPage(ScriptContext As ScriptingContext)

  TimeBombX xtbCharacter, 15656, 12934, 3080, 14624
  
  Set ScriptContext = ScriptContext
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application
 
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
  sstrVersion = App.Major & "." & App.Minor & "." & App.Revision
  sstrCompanyName = App.CompanyName
  sstrProductName = App.ProductName

  FontDirectory = "/Img/GraphChar/"
  Font = "Cil"

End Sub

Rem -------------------------------------------------------------------------
Rem  Mostra os caracteres graficos
Rem -------------------------------------------------------------------------
Public Sub DrawString(ByVal strString As String)
  
  Dim i As Integer

  For i = 1 To Len(strString)
    ShowGraphChar Mid(strString, i, 1), Font
  Next

End Sub

Rem -------------------------------------------------------------------------
Rem  Desenha um botão
Rem -------------------------------------------------------------------------
Public Sub DrawButton(ByVal strString As String)

  ShowGraphChar "Left", Font
  DrawString strString
  ShowGraphChar "Right", Font

End Sub

Rem -------------------------------------------------------------------------
Rem E uma Tag de final de linha
Rem -------------------------------------------------------------------------
Public Function IsValidChar(ByVal strChar As String) As Boolean
  
  Dim Result

  Result = False

  If InStr("abcdefghijklmnopqrstuvwxyz", strChar) Then
    Result = True
  ElseIf InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", strChar) Then
    Result = True
  ElseIf InStr("0123456789", strChar) Then
    Result = True
  ElseIf InStr("áéíóúÁÉÍÓÚýÝ", strChar) Then
    Result = True
  ElseIf InStr("àèìòùÀÈÌÒÙ", strChar) Then
    Result = True
  ElseIf InStr("äëïöüÄËÏÖÜ", strChar) Then
    Result = True
  ElseIf InStr("âêîôûÂÊÎÔÛ", strChar) Then
    Result = True
  ElseIf InStr("ãõÃÕ", strChar) Then
    Result = True
  ElseIf InStr("Çç", strChar) Then
    Result = True
  End If

  IsValidChar = Result

End Function

Rem -------------------------------------------------------------------------
Rem  Mostra os caracteres graficos
Rem -------------------------------------------------------------------------
Public Function IsCapital(ByVal strChar As String) As Boolean
  Dim blnResult
  
  blnResult = False
  
  If strChar > "" Then
    strChar = Left(strChar, 1)
    
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÝÀÈÌÒÙÄËÏÖÜÂÊÎÔÛÃÕÇ", strChar) Then
      blnResult = True
      
    End If
  End If
  
  IsCapital = blnResult

End Function

Rem -------------------------------------------------------------------------
Rem  Mostra os caracteres graficos
Rem -------------------------------------------------------------------------
Public Function IsSmall(ByVal strChar As String) As Boolean
  Dim blnResult
  
  blnResult = False
  
  If strChar > "" Then
    strChar = Left(strChar, 1)
    
    If InStr("abcdefghijklmnopqrstuvwxyzáéíóúýàèìòùäëïöüâêîôûãõç", strChar) Then
      blnResult = True
      
    End If
  End If
  
  IsSmall = blnResult

End Function

Rem -------------------------------------------------------------------------
Rem  Mostra os caracteres graficos
Rem -------------------------------------------------------------------------
Public Function IsNumber(ByVal strChar As String) As Boolean
  Dim blnResult
  
  blnResult = False
  
  If strChar > "" Then
    strChar = Left(strChar, 1)
    
    If InStr("0123456789", strChar) Then
      blnResult = True
      
    End If
  End If
  
  IsNumber = blnResult

End Function

Rem -------------------------------------------------------------------------
Rem  Mostra os caracteres graficos
Rem -------------------------------------------------------------------------
Public Sub ShowGraphString(ByVal strString As String, _
                           ByVal strFont As String)
  
  DrawString strString

End Sub

Rem -------------------------------------------------------------------------
Rem  Mostra os caracteres graficos
Rem -------------------------------------------------------------------------
Public Sub ShowGraphChar(ByVal strChar As String, _
                         ByVal strFont As String)
  
  Dim strFontDir, strFontFull, strExt, strCaps, strCharResult
  
  If strChar > "" Then
    strExt = GetGraphCharExt(strFont)
    strCaps = ""
  
    If Len(strChar) > 1 Then
      Select Case LCase(strChar)
        Case "left"
        Case "right"
      Case Else
        strChar = ""
        
      End Select
    Else
      
'      If IsCapital(strChar) Then
'        strCaps = "C"
'      ElseIf IsSmall(strChar) Then
'        strCaps = "S"
'      End If

      Select Case strChar
        Case "?" ' question mark
          strCharResult = "quest"

        Case "*" ' asterix sign
          strCharResult = "ast"

        Case "/" ' forward slash
          strCharResult = "fslash"

        Case "\" ' backward slash
          strCharResult = "bslash"

        Case "#" ' Number sign
          strCharResult = "Number"

        Case ":" ' Number sign
          strCharResult = "colon"

        Case Else
          If strChar = " " Then
            strCharResult = "Space"
          End If
        End Select
  
      If strCharResult > "" Then
        strChar = strCharResult
      End If
  
    End If
    
    strFontDir = Session("htDefaultHiperToolsWorkPath") & FontDirectory & strFont & "/"
    strFontFull = strFontDir & strFont & strCaps & strChar & strExt
 
    Response.Write "<IMG SRC=""" & strFontFull & """ ALT=""" & strChar & """>"
    
  End If
End Sub

Rem -------------------------------------------------------------------------
Rem  Mostra os caracteres graficos
Rem -------------------------------------------------------------------------
Public Function GetGraphCharExt(ByVal strFont As String) As String
  
  Dim strExt
  
  If IsFileExist(Server.MapPath(Session("htDefaultHiperToolsWorkPath") & _
     FontDirectory) & "\" & strFont & "\" & strFont & "Space.gif") Then
     strExt = ".gif"
    
  ElseIf IsFileExist(Server.MapPath(Session("htDefaultHiperToolsWorkPath") & _
     FontDirectory) & "\" & strFont & "\" & strFont & "Space.jpg") Then
     strExt = ".jpg"
      
  Else
     strExt = ""
      
  End If

  GetGraphCharExt = strExt
  
End Function

Rem -------------------------------------------------------------------------
Rem  Mostra os caracteres graficos
Rem -------------------------------------------------------------------------
Public Sub ShowButton(ByVal strString As String, _
                      ByVal strFont As String)

  ShowGraphChar "Left", Font
  DrawString strString
  ShowGraphChar "Right", Font

End Sub

Rem -------------------------------------------------------------------------
Rem  Mostra o botao do tipo cilindrico
Rem -------------------------------------------------------------------------
Public Sub ShowButtonCil(ByVal strString As String)
  
  Dim strOldFont As String
  
  strOldFont = Font
  
  Font = "Cil"
  DrawButton strString
  
  Font = strOldFont
End Sub

Rem -------------------------------------------------------------------------
Rem  Mostra o botao do tipo cilindrico
Rem -------------------------------------------------------------------------
Public Sub ShowButtonWin(ByVal strString As String)
  
  Dim strOldFont As String
  
  strOldFont = Font
  
  Font = "Win"
  DrawButton strString
  
  Font = strOldFont

End Sub

Rem -------------------------------------------------------------------------
Rem  ShowNumbers(nNumbers, cDigit)
Rem -------------------------------------------------------------------------
Rem  Entradas    : nNumbers = intero a ser mostrado
Rem                (123456789)
Rem                cDigit = Tipo do digitos
Rem                ("Green")
Rem  Saidas      :
Rem -------------------------------------------------------------------------
Public Sub ShowNumbers(ByVal intNumbers As Integer, _
                       ByVal strFont As String)

  If intNumbers = "" Then
    intNumbers = 0

  End If
  
  If UCase(strFont) = "ABACUS" Then
    ShowGraphChar "Left", strFont
  End If
  
  ShowGraphString intNumbers, strFont

  If UCase(strFont) = "ABACUS" Then
    ShowGraphChar "Left", strFont
  End If

End Sub

Rem -------------------------------------------------------------------------
Rem Verifica se o arquivo existe
Rem -------------------------------------------------------------------------
Public Function IsFileExist(ByVal strFullPath As String) As Boolean
  Dim objFS

  Set objFS = Server.CreateObject("Scripting.FileSystemObject")

  If objFS.GetDriveName(strFullPath) = "" Then
    strFullPath = Server.MapPath(strFullPath)
  
  End If
  
  IsFileExist = objFS.FileExists(strFullPath)

  Set objFS = Nothing
  
End Function
