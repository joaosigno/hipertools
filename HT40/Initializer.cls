VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Initializer"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =========================================================================
'  /HiperTools/Initializer.inc
' -------------------------------------------------------------------------
'  Descricao   : Include do Inicializador
'  Cria‡ao     : 11:23h 14/2/1998
'  Local       : Brasilia/DF
'  Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
'  Versao      : 1.0.0
'  Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------
'  ALTERACOES
' -------------------------------------------------------------------------
'  Responsavel : [Nome do executante da alteracao]
'  Data/Hora   : [Data e hora da alteracao]
'  Resumo      : [Resumo descritivo da alteracao executada]
' -------------------------------------------------------------------------
  
  Option Explicit

' -------------------------------------------------------------------------
' Constantes dos Tipos de Conexoes
' -------------------------------------------------------------------------

' -------------------------------------------------------------------------
' Criado (Ridai) - constate de retorno para o property HiperToolsWorkPath
' Quando o não houver nada no Session("htHiperToolsWorkPath");
Private Const htDefaultError = -201
Private Const htDefaultHiperToolsWorkPath = "/HiperTools"
Private Const cfDefaultErrorMessagePortuguese = "Ocorreu um erro interno do sistema. " & vbCrLf & _
              "Entre em contato com o(a) desenvolvedor(a) " & vbCrLf & _
              "e relate o ocorrido com o maior número de detalhes."

Private Const cfDefaultErrorMessageEnglish = "System internal error." & vbCrLf & _
              "Contact your developer and report with more detailed information as possible."

Private sstrScriptURL As String, _
  sstrScriptName As String, _
  sstrHTTPHost As String, _
  sstrScriptFull As String, _
  sstrSiteRootDir As String, _
  sstrQueryString As String, _
  sstrProtocol As String

Private ScriptContext As ScriptingContext

Private sstrComments As String, _
    sstrClassTitle As String, _
    sstrVersion As String, _
    sstrCompanyName As String, _
    sstrProductName As String, _
    sstrLegalCopyright As String

Private Registry As HiperTools30.Registry

'Private HTConfig As HiperTools30.Config

Private Request As Request, _
    Response As Response, _
    Server As Server, _
    Session As Session, _
    Application As Application
    
Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  TimeBombX xtbInitializer, 15656, 12934, 3080, 14624
  
  Set ScriptContext = ScriptContext
  
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application
  
  Set Registry = Server.CreateObject("HiperTools30.Registry")
  
  'Alterado (Ridai):
  If Not Session("Default") > "" Then Response.Expires = 0
  Session("htDefaultHiperToolsWorkPath") = htDefaultHiperToolsWorkPath
  '---
  
  SystemOptions

  sstrProtocol = IIf(Request.ServerVariables("HTTPS") = "on", "https", "http")
  sstrHTTPHost = Request.ServerVariables("HTTP_HOST")
  sstrScriptName = Request.ServerVariables("SCRIPT_NAME")
  sstrScriptURL = sstrProtocol & "://" & sstrHTTPHost & sstrScriptName
  sstrSiteRootDir = Request.ServerVariables("APPL_PHYSICAL_PATH")
  sstrScriptFull = Request.ServerVariables("PATH_TRANSLATED")
  
  SaveQueryString
  
  If Not Session("Default") > "" Then InitDefault

End Sub

Public Sub OnEndPage()
  
  Set ScriptContext = Nothing
  
  Set Request = Nothing
  Set Response = Nothing
  Set Server = Nothing
  Set Session = Nothing
  Set Application = Nothing
  Set Registry = Nothing
  
  
End Sub

Public Property Get Language() As Integer
  Language = Session("HTLanguage")

End Property

Public Property Let Language(intLanguage As Integer)
  Session("HTLanguage") = intLanguage
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ScriptURL() As String
  ScriptURL = sstrScriptURL
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ScriptName() As String
  ScriptName = sstrScriptName
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get HTTPHost() As String
  HTTPHost = sstrHTTPHost
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ScriptFull() As String
  ScriptFull = sstrScriptFull
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get SiteRootDir() As String
  SiteRootDir = sstrSiteRootDir
End Property

Rem =========================================================================
Rem QueryString anterior
Rem -------------------------------------------------------------------------
Public Property Get LastQueryString() As String
  LastQueryString = Session("LastQueryString")
End Property

Rem =========================================================================
Rem QueryString atual
Rem -------------------------------------------------------------------------
Public Property Get QueryString() As String
  QueryString = sstrQueryString
End Property

Rem =========================================================================
Rem QueryString atual
Rem -------------------------------------------------------------------------
Public Property Get Protocol() As String
  Protocol = sstrProtocol
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub InitDefault()
  
  Session("Default") = "HiperTools"
  
  Session("HTLanguage") = 0
  Session("HTLicenseDate") = ""
  Session("HTLicenserName") = ""
  Session("HTLicenserCompany") = ""
  Session("HTDemoDate") = ""
  
  Registry.AutoCreate = True
  
  Registry.hKey = rgHKEYLocalMachine
 
  On Error Resume Next
  
  Rem [HKEY_LOCAL_MACHINE\SOFTWARE\HiperTools]
  Registry.Key = "SOFTWARE\HiperTools"
  
  Session("HTLanguage") = Registry.ValueAsLong("DefaultLanguage")
  
  Rem [HKEY_LOCAL_MACHINE\SOFTWARE\HiperTools\3.0]
  Registry.Key = "SOFTWARE\HiperTools\" & App.Major & "." & App.Minor
  
  Session("HTLicenseDate") = Registry.ValueAsString("LicenseDate")
  Session("HTLicenserName") = Registry.ValueAsString("LicenserName")
  Session("HTLicenserCompany") = Registry.ValueAsString("LicenserCompany")
  Session("HTDemoDate") = Registry.ValueAsString("DemoDate")

  If Session("HTLicenseDate") > "" Then
    Session("HTLicenseDate") = Left(Session("HTLicenseDate"), Len(Session("HTLicenseDate")) - 1)
    
  End If
  
  If Session("HTLicenserName") > "" Then
    Session("HTLicenserName") = Left(Session("HTLicenserName"), Len(Session("HTLicenserName")) - 1)
    
  End If
  
  If Session("HTLicenserCompany") > "" Then
    Session("HTLicenserCompany") = Left(Session("HTLicenserCompany"), Len(Session("HTLicenserCompany")) - 1)
    
  End If
  
  If Session("HTDemoDate") > "" Then
    Session("HTDemoDate") = Left(Session("HTDemoDate"), Len(Session("HTDemoDate")) - 1)
    
  End If
  
  On Error GoTo 0
  
Rem  Session("MetaPragma") = "NO-CACHE"
Rem  Session("MetaContentType") = "text/html; charset=iso-8859-1"
Rem  Session("MetaReplyTo") = "webmaster@zevallos.com.br"
Rem  Session("MetaExpires") = Now + 90
Rem  Session("MetaKeywords") = "Zevallos"
Rem  Session("MetaKeyPhrase") = ""
Rem  Session("MetaDescription") = ""
Rem  Session("MetaProduct") = "HiperTools v3.0 - Biblioteca ASP"
Rem  Session("MetaLocale") = "PO-BR"
Rem  Session("MetaCharSet") = "US-ASCII"
Rem  Session("MetaCategory") = "HOME PAGE"
Rem  Session("MetaAuthor") = "ZTI - Zevallos(r) Tecnologia em Informacao - Brasilia - DF -" & _
                                " Brasil - webmaster@zevallos.com.br - http://www.zevallos.com.br"
Rem  Session("MetaDistribution") = "global"
Rem  Session("MetaTechnology") = "SCRIPTING"
  
  Session("MetaPicsLabel") = "(PICS-1.1 ""http://www.rsac.org/ratingsv01.html""" & _
                                " l gen true comment ""RSACi North America Server"" by" & _
                                " ""webmaster@zevallos.com.br"" for """ & _
                                "http://www.zevallos.com.br""" & _
                                " on ""1997.06.26T21:24-0500"" r (n 0 s 0 v 0 l 0))"
  Session("MetaGenerator") = App.Comments
  Session("MetaFormatter") = App.Comments

' -------------------------------------------------------------------------
' Constantes BASE FONT de configuracao do sistema
' -------------------------------------------------------------------------
  Session("BaseFont") = "Arial, Helvetica, Sans-Serif"

' -------------------------------------------------------------------------
' Constantes LINK de configuracao do sistema
' -------------------------------------------------------------------------
  Session("LinkHomeHRef") = "/Default.asp"
  Session("LinkHomeTitle") = Session("MetaFormatter")

  Session("LinkCopyrightHRef") = "/copyright.htm"
  Session("LinkCopyrightTitle") = "Copyright da Zevallos"

  Session("LinkMadeHRef") = "mailto:webmaster@zevallos.com.br"
  Session("LinkMadeTitle") = "WebMaster do HiperTools"

  Session("LinkStyleSheetHRef") = ""
  Session("LinkStyleSheetTitle") = ""
  
' -------------------------------------------------------------------------
' Constantes BODY de configuracao do sistema
' -------------------------------------------------------------------------
  Session("BodyBGColor") = "white"
  Session("BodyBackground") = ""
  Session("BodyBGProperties") = ""
  Session("BodyStyle") = ""
  Session("BodyText") = ""
  Session("BodyLink") = ""
  Session("BodyVLink") = ""
  Session("BodyALink") = ""
  Session("BodyTopMargin") = 0
  Session("BodyLeftMargin") = 0
  Session("BodyWidth") = 780

' -------------------------------------------------------------------------
' Constantes FOOTER de configuracao do sistema
' -------------------------------------------------------------------------
  Select Case Session("HTLanguage")
    Case 1 ' English
      Session("FooterTitle") = "A product from <A HREF=http://www.zevallos.com.br>" & _
                                    "Zevallos&reg;</A> Tecnologia em Informa&ccedil;&atilde;o."
      Session("FooterSugestions") = "Sugestions or problems send to " & _
                                    "<A HREF=mailto:hipertools@hipertools.com.br>" & _
                                    "<img src=""" & HiperToolsWorkPath & "/img/mailto.gif"" alt=""Mail To"" border=""0""> " & _
                                    "<i>&lt;webmaster@zevallos.com.br&gt;</i></A>"
      Session("FooterCopyright") = "&copy; 97-2000 <A HREF =/copyright.asp>Zevallos&reg;</a>" & _
                                    " all rights reserved."
    Case Else
      Session("FooterTitle") = "Um produto da <A HREF=http://www.zevallos.com.br>" & _
                                    "Zevallos&reg;</A> Tecnologia em Informa&ccedil;&atilde;o."
      Session("FooterSugestions") = "Sugest&otilde;es ou problemas encaminhar para o " & _
                                    "<A HREF=mailto:hipertools@hipertools.com.br>" & _
                                    "<img src=""" & HiperToolsWorkPath & "/img/mailto.gif"" alt=""Mail To"" border=""0""> " & _
                                    "<i>&lt;webmaster@zevallos.com.br&gt;</i></A>"
      Session("FooterCopyright") = "&copy; 97-2000 <A HREF =/copyright.asp>Zevallos&reg;</a>" & _
                                    " todos os direitos reservados."
    End Select
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub QueryStringSave() 'obsoleto
  SaveQueryString
End Sub

Rem =========================================================================
Rem usado para salvar a QueryString anterior e a atual
Rem -------------------------------------------------------------------------
Public Sub SaveQueryString()
  sstrQueryString = Request.ServerVariables("QUERY_STRING")
  If Session("LastQueryString") <> Session("CurrentQueryString") Then
    Session("LastQueryString") = Session("CurrentQueryString")
    Session("CurrentQueryString") = sstrQueryString
  End If
End Sub


Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let HiperToolsWorkPath(ByVal strPath As String)
  Dim strAux As String
  
  strPath = Replace(strPath, "\", "/")
  strAux = Right(strPath, 1)
  
  If strAux = "/" Then strPath = Left(strPath, Len(strPath) - 1)
  
  'SaveSetting "HiperTools30", "Config", "HiperToolsWorkPath", strPath
  Session("htDefaultHiperToolsWorkPath") = strPath
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get HiperToolsWorkPath() As String
    If Not Session("htDefaultHiperToolsWorkPath") > "" Then
        Rem Session("htDefaultHiperToolsWorkPath") = GetSetting("HiperTools30", _
                "Config", "HiperToolsWorkPath", htDefaultHiperToolsWorkPath)
        Session("htDefaultHiperToolsWorkPath") = htDefaultHiperToolsWorkPath
    End If
    
    HiperToolsWorkPath = Session("htDefaultHiperToolsWorkPath")
End Property

Private Sub SystemOptions()
  ' -----------------------------------------------------------------------
  ' Inicializa os defaults do sistema
  ' -----------------------------------------------------------------------
  If Request.QueryString("Default") > "" Then
    Session("Default") = ""

  End If
  
  ' -----------------------------------------------------------------------
  ' Ativa o icone de ASP Source Code
  ' -----------------------------------------------------------------------
  If Request.QueryString("Source") > "" Then
    Select Case Request.QueryString("Source")
      Case 0
        Session("SourceCode") = False

      Case 1
        Session("SourceCode") = True

      Case Else
    End Select
  Else
    If IsEmpty(Session("SourceCode")) Then Session("SourceCode") = False

  End If

  ' -----------------------------------------------------------------------
  ' Define a borda das tabelas
  ' -----------------------------------------------------------------------
  If Request.QueryString("Border") > "" Then
    Session("TableBorder") = Request.QueryString("Border")

  Else
    If IsEmpty(Session("TableBorder")) Then Session("TableBorder") = "0"

  End If

End Sub

Public Property Let ErrorMessage(ByVal strMessage As String)
  Session("htITErrorMessage") = strMessage
End Property

Public Property Get ErrorMessage() As String
  Dim strMessage As String
  strMessage = Session("htITErrorMessage")
  
  If Not strMessage > "" Then
    If Language = 1 Then
      strMessage = cfDefaultErrorMessageEnglish
    
    Else
      strMessage = cfDefaultErrorMessagePortuguese
    
    End If
  End If
  Rem strMessage = Replace(strMessage, vbCrLf, vbCrLf & "<BR>" & vbCrLf)

  ErrorMessage = strMessage
End Property

Public Property Get WhenStarted() As Date
  WhenStarted = gdatServerStarted
End Property

'Criado (Ridai)
Rem =========================================================================
Rem Levanta uma excecao básica utilizando o ErrorMessage
Rem -------------------------------------------------------------------------
Public Sub ShowDefaultError()
  Err.Raise vbObject + htDefaultError, "HiperTools30", ErrorMessage
End Sub


