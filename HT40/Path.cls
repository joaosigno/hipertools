VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Path"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =========================================================================
' Path.cls
' -------------------------------------------------------------------------
' Descricao   : Include de Path
' Criacao     : 11:23h 7/2/1998
' Local       : Brasilia/DF
' Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
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

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  TimeBombX xtbPath, 15656, 12934, 3080, 14624
  
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
    sstrVersion = App.Major & "." & _
                App.Minor & "." & _
                App.Revision
    sstrCompanyName = App.CompanyName
    sstrProductName = App.ProductName

End Sub

Rem =========================================================================
Rem Retorna o tipo de barra utilizado no path
Rem -------------------------------------------------------------------------
Public Function SlashType(ByVal strFullPath As String) As String
  Dim i As Long
  SlashType = ""

  For i = 1 To Len(strFullPath)
    Select Case Mid(strFullPath, i, 1)
    Case "/", "\"
      SlashType = Mid(strFullPath, i, 1)
      Exit For
    End Select
  Next
  
End Function

Rem =========================================================================
Rem Estrai o nome do arquivo de um Full Path
Rem -------------------------------------------------------------------------
Public Function InvertSlash(ByVal strFullPath As String) As String
  Dim strResult, strSlash, strInvSlash
  Dim intSlashPos, intLen

  strResult = ""
  
  If strFullPath > "" Then
    intSlashPos = 0
    strResult = strFullPath
    intLen = Len(strFullPath)
    strSlash = SlashType(strFullPath)
    
    If strSlash = "\" Then
      strInvSlash = "/"
      
    ElseIf strSlash = "/" Then
      strInvSlash = "\"
      
    End If
    
    strResult = Replace(strResult, strSlash, strInvSlash)
    
    'Alterado (Ridai Govinda)
'    Do While intSlashPos > 0
'      intSlashPos = InStr(intSlashPos + 1, strResult, strSlash)
'
'      If intSlashPos > 0 Then
'        strResult = Left(strResult, intSlashPos - 1) & strInvSlash & Right(strResult, intLen - intSlashPos)
'
'      End If
'
'    Loop

  End If
  
  InvertSlash = strResult
  
End Function

Rem =========================================================================
Rem Converte o path para barras /
Rem -------------------------------------------------------------------------
Public Function URLSlash(ByVal strFullPath As String) As String
  Dim strResult

  strResult = strFullPath

  If strFullPath > "" Then
    If SlashType(strFullPath) = "\" Then
      strResult = InvertSlash(strFullPath)

      If Drive(strFullPath) > "" Then
        strResult = Mid(strResult, 3)
        
      End If
    End If
  End If
  
  URLSlash = strResult
  
End Function

Rem =========================================================================
Rem Converte o path para barras \
Rem -------------------------------------------------------------------------
Public Function NormalSlash(ByVal strFullPath As String) As String
  Dim strResult
  
  strResult = strFullPath

  If strFullPath > "" Then
    If SlashType(strFullPath) = "/" Then
      strResult = InvertSlash(strFullPath)
      
    End If
  End If
  
  NormalSlash = strResult
  
End Function

Rem =========================================================================
Rem Remove a última barra
Rem -------------------------------------------------------------------------
Public Function RemoveLastSlash(ByVal strFullPath As String) As String
  Dim strResult

  strResult = strFullPath
    
  If strFullPath > "" Then
    If Right(strResult, 1) = SlashType(strResult) Then
      strResult = Left(strResult, Len(strResult) - 1)
      
    End If
  End If
    
  RemoveLastSlash = strResult
  
End Function

Rem =========================================================================
Rem Extrai o a letra do Drive de um Full Path
Rem -------------------------------------------------------------------------
Public Function Only(ByVal strFullPath As String) As String
  Dim strResult, intColon
  
  strResult = ""

  If strFullPath > "" Then
    strFullPath = Path(strFullPath)
    
    intColon = InStr(strFullPath, ":")
    
    If intColon > 0 Then
      strResult = Mid(strFullPath, intColon + 1)
      
    End If
  End If
  
  Only = strResult

End Function

Rem =========================================================================
Rem Extrai a letra do Drive de um Full Path
Rem -------------------------------------------------------------------------
Public Function Drive(ByVal strFullPath As String) As String
  Dim strResult
  
  strResult = ""

  If strFullPath > "" Then
    If InStr(strFullPath, ":") = 2 Then
      strResult = Left(strFullPath, 2)
      
    End If
  End If
  
  Drive = strResult

End Function

Rem =========================================================================
Rem Mantidas apenas por questão de compatibilidade
Rem ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Public Function GetSlash(ByVal strFullPath As String) As String 'Obsoleto
  GetSlash = SlashType(strFullPath)
End Function

Public Function URL(ByVal strFullPath As String) As String 'Obsoleto
  URL = URLSlash(strFullPath)
End Function

Public Function Normal(ByVal strFullPath As String) As String 'Obsoleto
  Normal = NormalSlash(strFullPath)
End Function

Public Function DelRightSlash(ByVal strFullPath As String) As String 'Obsoleto
  DelRightSlash = RemoveLastSlash(strFullPath)
End Function

Rem =========================================================================
Rem Verifica se é uma URL
Rem -------------------------------------------------------------------------
Public Function IsURL(ByVal strURL As String) As Boolean 'Obsoleto

  strURL = UCase(Trim(strURL))

  IsURL = True

  If Left(strURL, 6) = "HTTP:/" Then
  ElseIf Left(strURL, 6) = "FILE:/" Then
  ElseIf Left(strURL, 7) = "MAILTO:" Then
  ElseIf Left(strURL, 5) = "FTP:/" Then
  ElseIf Left(strURL, 8) = "GOPHER:/" Then
  ElseIf Left(strURL, 6) = "NEWS:/" Then
  ElseIf Left(strURL, 7) = "HTTPS:/" Then
  ElseIf Left(strURL, 8) = "TELNET:/" Then
  ElseIf Left(strURL, 6) = "NNTP:/" Then
  Else
    IsURL = False
  End If

End Function

Rem =========================================================================
Rem Extrai o nome do arquivo de um Full Path
Rem -------------------------------------------------------------------------
Public Function FileName(ByVal strFullPath As String) As String 'Obsoleto
  Dim strResult, strSlash
  
  strResult = ""

  If strFullPath > "" Then
    If InStr(strFullPath, ".") > 0 Then
      strResult = strFullPath
  
      strSlash = SlashType(strFullPath)
      
      If strSlash > "" Then
        strResult = Mid(strFullPath, InStrRev(strFullPath, strSlash) + 1)
        
      End If
    End If
  End If
    
  FileName = strResult
    
End Function

Rem =========================================================================
Rem Extrai o Path de um Full Path
Rem -------------------------------------------------------------------------
Public Function Path(ByVal strFullPath As String) As String 'Obsoleto
  Dim strResult

  strResult = ""
  
  If strFullPath > "" Then
    strResult = strFullPath
    
    If InStr(strFullPath, ".") > 0 Then
      strResult = Left(strFullPath, Len(strFullPath) - Len(FileName(strFullPath)))
      
    End If
  End If
   
  Path = strResult

End Function

Rem =========================================================================
Rem Demo do Path
Rem -------------------------------------------------------------------------
Public Sub Demo() 'Obsoleto
  
  Dim strFullPath
  
  strFullPath = Request.ServerVariables("PATH_TRANSLATED")

  Response.Write "<BR>FullPath=" & strFullPath
  Response.Write "<BR>SlashType=" & SlashType(strFullPath)
  Response.Write "<BR>InvertSlash=" & InvertSlash(strFullPath)
  Response.Write "<BR>URLSlash=" & URLSlash(strFullPath)
  Response.Write "<BR>NormalSlash=" & NormalSlash(strFullPath)
  Response.Write "<BR>RemoveLastSlash=" & RemoveLastSlash(strFullPath)
  Response.Write "<BR>Only=" & Only(strFullPath)
  Response.Write "<BR>Drive=" & Drive(strFullPath)
  Response.Write "<BR>IsURL=" & IsURL(strFullPath)
  
  strFullPath = Request.ServerVariables("PATH_INFO")
  
  Response.Write "<BR>"
  Response.Write "<BR>FullPath=" & strFullPath
  Response.Write "<BR>SlashType=" & SlashType(strFullPath)
  Response.Write "<BR>InvertSlash=" & InvertSlash(strFullPath)
  Response.Write "<BR>URLSlash=" & URLSlash(strFullPath)
  Response.Write "<BR>NormalSlash=" & NormalSlash(strFullPath)
  Response.Write "<BR>RemoveLastSlash=" & RemoveLastSlash(strFullPath)
  Response.Write "<BR>Only=" & Only(strFullPath)
  Response.Write "<BR>Drive=" & Drive(strFullPath)
  Response.Write "<BR>IsURL=" & IsURL(strFullPath)

End Sub
