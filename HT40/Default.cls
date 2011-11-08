VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Default"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =========================================================================
'  HiperTools30.Default
' -------------------------------------------------------------------------
'  Descricao   : Informações Default do HiperTools
'  Criacao     : ???
'  Local       : Brasilia/DF
'  Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
'  Versao      : 1.0.0
'  Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------
'  ALTERACOES
' -------------------------------------------------------------------------
'  Responsavel : Flavio Costa <flavio@zevallos.com.br>
'  Data/Hora   : 07/03/2000 12:00
'  Resumo      : Alteracoes para lancamento do produto
' -------------------------------------------------------------------------

Option Explicit
  
Private sobjShow As HiperTools30.Show
Private Initializer As HiperTools30.Initializer

Private sstrHeaderTitle As String
Private sstrHeaderIcon As String

Private Request As Request, _
        Response As Response, _
        Server As Server, _
        Session As Session, _
        Application As Application
    
Private ScriptContext As ScriptingContext

Private sstrComments As String, _
        sstrTitle As String, _
        sstrVersion As String, _
        sstrCompanyName As String, _
        sstrProductName As String, _
        sstrLegalCopyright As String

Public IgnoreBodyWidth As Boolean
Public CenterBody As Boolean
Public PageHeader As String
Public PageFooter As String
Public BodyOnBeforeUnload As String
Public BodyOnBlur As String
Public BodyOnError As String
Public BodyOnFocus As String
Public BodyOnHelp As String
Public BodyOnLoad As String
Public BodyOnResize As String
Public BodyOnScroll As String
Public BodyOnUnload As String

Public Sub OnStartPage(ScriptContext As ScriptingContext)

  TimeBombX xtbDefault, 15656, 12934, 3080, 14624
  
  Set ScriptContext = ScriptContext
  
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application

  Set sobjShow = Server.CreateObject("HiperTools30.Show")
  Set Initializer = Server.CreateObject("HiperTools30.Initializer")

  IgnoreBodyWidth = False
  BodyWidth = ""
  CenterBody = False
  
  If Not Session("HTHiddenDefault") > "" Then
    Session("HTHiddenDefault") = "1"
    
    Session("HTHideUserID") = False
    Session("HTHideFileInfo") = False
    Session("HTHideSiteCounter") = False
    Session("HTHideInnerInfo") = False
    Session("HTHideLicenserName") = False
    
    Session("HTHideHeadMeta") = False
    Session("HTHideHeadBaseFont") = False
    Session("HTHideHeadLink") = False
  
  End If

End Sub

Public Sub OnEndPage()

  Set ScriptContext = Nothing
  
  Set Request = Nothing
  Set Response = Nothing
  Set Server = Nothing
  Set Session = Nothing
  Set Application = Nothing

  Set sobjShow = Nothing
  Set Initializer = Nothing

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub BeginHTML()
  
  sobjShow.CrLfAfter = True
  sobjShow.HTML "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN" & _
             """ ""http://www.w3.org/TR/REC-html40/loose.dtd"">"

  sobjShow.HTML "<HTML>"
  sobjShow.HTMLComment App.Comments & " - Build: " & App.Major & "." & App.Minor & "." & App.Revision & " | " & _
                       "1997-2001 (c) Zevallos(r) Tecnologia em Informacao | " & _
                       "Brasilia - DF - Brasil | " & _
                       "+55 61 328-3575 | " & _
                       "info@hipertools.com.br | " & _
                       "http://www.hipertools.com.br"

  If Session("HTLicenserName") > 0 And Not Session("HTHideLicenserName") Then
    sobjShow.HTMLComment "Licensed to : " & Session("HTLicenserName") & " | " & Session("HTLicenserCompany") & " | " & Session("HTLicenseDate")
    sobjShow.HTMLComment "Language : " & Session("HTLanguage")
  
  End If
    
  If Session("HTDemoDate") > 0 And Not Session("HTHideLicenserName") Then
    sobjShow.HTMLComment "Demo version since : " & Session("HTDemoDate")
  
  End If
  
  If Session("UserID") > 0 And Not Session("HTHideUserID") Then
    sobjShow.HTMLComment "User ID: " & Session("UserID")
    
  End If

  If Session("SiteCounter") > 0 And Not Session("HTHideSiteCounter") Then
    sobjShow.HTMLComment "Site Counter: " & Session("SiteCounter")
    
  End If
  
  Dim objFS, objFile
 
  Set objFS = Server.CreateObject("Scripting.FileSystemObject")
  Set objFile = objFS.GetFile(Initializer.ScriptFull)
  
  If Not Session("HTHideFileInfo") Then
    sobjShow.HTMLComment "File Name: " & Initializer.ScriptURL
    sobjShow.HTMLComment "Created: " & objFile.DateCreated
    sobjShow.HTMLComment "Last Accessed: " & objFile.DateLastAccessed
    sobjShow.HTMLComment "Last Modified: " & objFile.DateLastModified
  End If
  
  Set objFile = Nothing
  Set objFS = Nothing

  'sobjShow.HTMLComment "Script Engine: " & ScriptEngine & " - " & _
                  ScriptEngineMajorVersion & "." & _
                  ScriptEngineMinorVersion & " (Build:" & _
                  ScriptEngineBuildVersion & ")"
  
  If Not Session("HTHideInnerInfo") Then
    sobjShow.HTMLComment "Initializer.HTTPHost    = " & Initializer.HTTPHost
    sobjShow.HTMLComment "Initializer.ScriptName  = " & Initializer.ScriptName
    sobjShow.HTMLComment "Initializer.ScriptURL   = " & Initializer.ScriptURL
    sobjShow.HTMLComment "Initializer.SiteRootDir = " & Initializer.SiteRootDir
    sobjShow.HTMLComment "Initializer.ScriptFull  = " & Initializer.ScriptFull
    sobjShow.HTMLComment "Initializer.QueryString = " & Initializer.QueryString
  End If
  
  sobjShow.CrLfAfter = False

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub HTMLBegin()
  BeginHTML
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub EndHTML()
  sobjShow.HTMLCR "</HTML>"
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub HTMLEnd()
  EndHTML
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub BeginHead(Optional ByVal strTitle As String = "")
  sobjShow.HTMLComment "Head Section Begin"
  sobjShow.HTMLCR "<HEAD>"
  sobjShow.HTMLCR "<TITLE>" & strTitle & "</TITLE>"
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub HeadBegin(ByVal strTitle As String)
  BeginHead strTitle
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub EndHead()
  sobjShow.HTMLCR "</HEAD>"
  sobjShow.HTMLComment "Head Section End"
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub HeadEnd()
  EndHead
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub BodyBegin()
  BeginBody
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub BeginBody()
  Dim strResult

  strResult = "<BODY"

  If Session("BodyBGColor") > "" Then
    strResult = strResult & " BGCOLOR=" & Chr(34) & Session("BodyBGColor") & Chr(34)
    
  End If

  If Session("BodyBackground") > "" Then
    strResult = strResult & " BACKGROUND=" & Chr(34) & Session("BodyBackground") & Chr(34)
    
  End If

  If Session("BodyText") > "" Then
    strResult = strResult & " TEXT=" & Session("BodyText")
    
  End If

  If Session("BodyLink") > "" Then
    strResult = strResult & " LINK=" & Session("BodyLink")
    
  End If

  If Session("BodyVLink") > "" Then
    strResult = strResult & " VLINK=" & Session("BodyVLink")
    
  End If

  If Session("BodyALink") > "" Then
    strResult = strResult & " ALINK=" & Session("BodyALink")
    
  End If
  
  If Session("BodyTopMargin") > "" Then
    strResult = strResult & " TOPMARGIN=" & Session("BodyTopMargin")
    
  End If

  If Session("BodyLeftMargin") > "" Then
    strResult = strResult & " LEFTMARGIN=" & Session("BodyLeftMargin")
    
  End If

  If Session("BodyStyle") > "" Then
    strResult = strResult & " STYLE=" & Session("BodyStyle")
    
  End If

  If Session("BodyBGProperties") > "" Then
    strResult = strResult & " BGPROPERTIES=" & Session("BodyBGProperties")
    
  End If

  If BodyOnBeforeUnload > "" Then strResult = strResult & " OnBeforeUnload=""" & BodyOnBeforeUnload & """"
  If BodyOnBlur > "" Then strResult = strResult & " OnBlur=""" & BodyOnBlur & """"
  If BodyOnError > "" Then strResult = strResult & " OnError=""" & BodyOnError & """"
  If BodyOnFocus > "" Then strResult = strResult & " OnFocus=""" & BodyOnFocus & """"
  If BodyOnHelp > "" Then strResult = strResult & " OnHelp=""" & BodyOnHelp & """"
  If BodyOnLoad > "" Then strResult = strResult & " OnLoad=""" & BodyOnLoad & """"
  If BodyOnResize > "" Then strResult = strResult & " OnResize=""" & BodyOnResize & """"
  If BodyOnScroll > "" Then strResult = strResult & " OnScroll=""" & BodyOnScroll & """"
  If BodyOnUnload > "" Then strResult = strResult & " OnUnload=""" & BodyOnUnload & """"

  If Session("BodyJavaScript") > "" Then
    strResult = strResult & " " & Session("BodyJavaScript")
    
  End If

  sobjShow.HTMLComment "Main Body Section Begin"
  sobjShow.HTMLCR strResult & ">"

  If CenterBody Then sobjShow.HTMLCR "<CENTER>"
  
  If Not IgnoreBodyWidth And BodyWidth > "" Then
    strResult = "<TABLE WIDTH=" & Session("BodyWidth")
 
    If TableBorder > "" Then
      strResult = strResult & " BORDER=" & TableBorder
  
    Else
      strResult = strResult & " BORDER=0"
        
    End If
    
    sobjShow.HTMLCR strResult & " CELLPADDING=""0"" CELLSPACING=""0""><TR><TD>"
    
  End If
  
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub BodyLimit(ByVal blnState As Boolean)
  IgnoreBodyWidth = Not blnState
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub BodyCenter(ByVal blnState As Boolean)
  CenterBody = blnState
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub BodyEnd()
  EndBody
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub EndBody()
  
  If Not IgnoreBodyWidth And BodyWidth > "" Then sobjShow.HTMLCR "</TD></TR></TABLE>"
  If CenterBody Then sobjShow.HTMLCR "</CENTER>"
  
  sobjShow.HTMLCR "</BODY>"
  sobjShow.HTMLComment "Main Body Section End"

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub HeadAll(ByVal strTitle As String)
  WriteHead strTitle
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub WriteHead(Optional ByVal strTitle As String = "")
  BeginHead strTitle
  
  WriteMeta
    
  WriteBaseFont
    
  WriteLink
  
  EndHead

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub MetaDefault()
  
  WriteMeta
  
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub WriteMeta()
  If Not Session("HTHideHeadMeta") Then
    sobjShow.HTMLComment "Meta Section Begin"
  
    If Session("MetaPragma") > "" Then
      sobjShow.MetaHTTPEquiv "PRAGMA", Session("MetaPragma")
  
    End If
  
    If Session("MetaRefresh") > "" Then
      sobjShow.MetaHTTPEquiv "REFRESH", Session("MetaRefresh")
  
    End If
  
    If Session("MetaContentType") > "" Then
      sobjShow.MetaHTTPEquiv "Content-Type", Session("MetaContentType")
  
    End If
  
    If Session("MetaReplyTo") > "" Then
      sobjShow.MetaHTTPEquiv "REPLY-TO", Session("MetaReplyTo")
  
    End If
  
    If Session("MetaExpires") > "" Then
      sobjShow.MetaHTTPEquiv "EXPIRES", Session("MetaExpires")
    End If
  
    If Session("MetaPicsLabel") > "" Then
      sobjShow.MetaHTTPEquiv "PICS-Label", Session("MetaPicsLabel")
  
    End If
  
    If Session("MetaKeywords") > "" Then
      sobjShow.MetaName "KEYWORDS", Session("MetaKeywords") & _
                               ", ASP, VBScript, Active Server Pages, ZTI, HiperTools, Zevallos, Zevallos Tecnologia em Informacao, Brasil, Brazil, Brasilia, DF"
  
    End If
  
    If Session("MetaKeyPhrase") > "" Then
      sobjShow.MetaName "KEYPHASE", Session("MetaKeyPhrase")
  
    End If
  
    If Session("MetaDescription") > "" Then
      sobjShow.MetaName "DESCRIPTION", Session("MetaDescription")
  
    End If
  
    If Session("MetaProduct") > "" Then
      sobjShow.MetaName "PRODUCT", Session("MetaProduct")
  
    End If
  
    If Session("MetaLocale") > "" Then
      sobjShow.MetaName "LOCALE", Session("MetaLocale")
  
    End If
  
    If Session("MetaCharSet") > "" Then
      sobjShow.MetaName "CHARSET", Session("MetaCharSet")
  
    End If
  
    If Session("MetaCategory") > "" Then
      sobjShow.MetaName "CATEGORY", Session("MetaCategory")
    End If
  
    If Session("MetaGenerator") > "" Then
      sobjShow.MetaName "GENERATOR", Session("MetaGenerator")
  
    End If
  
    If Session("MetaAuthor") > "" Then
      sobjShow.MetaName "AUTHOR", Session("MetaAuthor")
  
    End If
  
    If Session("MetaDistribution") > "" Then
      sobjShow.MetaName "DISTRIBUTION", Session("MetaDistribution")
  
    End If
  
    If Session("MetaFormatter") > "" Then
      sobjShow.MetaName "FORMATTER", Session("MetaFormatter")
  
    End If
  
    If Session("MetaTechnology") > "" Then
      sobjShow.MetaName "TECHNOLOGY", Session("MetaTechnology")
  
    End If
  
    sobjShow.HTMLComment "Meta Section End"
  End If
  
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub BaseFontDefault()
  
  WriteBaseFont

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub WriteBaseFont()
  
  If Session("BaseFont") > "" And Not Session("HTHideHeadBaseFont") Then
    sobjShow.HTMLCR "<BASEFONT FACE=""" & Session("BaseFont") & """>"

  End If
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub LinkDefault()
  WriteLink
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub WriteLink()
  
  If Not Session("HTHideHeadLink") Then
    sobjShow.HTMLComment "Link Section Begin"
  
    If Session("LinkHomeHRef") > "" Then
      sobjShow.LinkRel "Home", Session("LinkHomeHRef"), Session("LinkHomeTitle")
  
    End If
  
    If Session("LinkCopyrightHRef") > "" Then
      sobjShow.LinkRel "Copyright", Session("LinkCopyrightHRef"), Session("LinkCopyrightTitle")
  
    End If
    
    If Session("LinkStyleSheetHRef") > "" Then
      sobjShow.LinkRelStyleSheet Session("LinkStyleSheetHRef")
  
    End If
  
    If Session("LinkMadeHRef") > "" Then
      sobjShow.LinkRev "Made", Session("LinkMadeHRef"), Session("LinkMadeTitle")
  
    End If
  
    sobjShow.HTMLComment "Link Section End"
  End If
  
End Sub
  
Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub WritePageHeader()
    sobjShow.HTMLComment "Main Header Begin"
    sobjShow.HTML PageHeader
    sobjShow.HTMLComment "Main Header End"
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub PageHeaderDefault(Optional ByVal strHeaderTitle As String = "")

  If strHeaderTitle > "" Then HeaderTitle = strHeaderTitle
  
  sobjShow.HTMLComment "Main Header Begin"
  sobjShow.Center
  sobjShow.HTML HeaderTitle
  sobjShow.CenterEnd
  sobjShow.HTMLComment "Main Header End"
  
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub WritePageFooter()
    sobjShow.HTMLComment "Main Footer Begin"
    sobjShow.HTML PageFooter
    sobjShow.HTMLComment "Main Footer End"
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub PageFooterDefault()
  sobjShow.HTMLCR "<br><HR>"
  sobjShow.HTMLCR "<table width=100% border=" & TableBorder + 0 & " cellpadding=0 cellspacing=0>"
  sobjShow.HTML "<tr>"

  If Session("SourceCode") Then
    sobjShow.HTMLCR "<td width=1% align=left>"
    sobjShow.HTMLCR "<A HREF=""" & Session("htDefaultHiperToolsWorkPath") & "/Code/code.asp?SourceFile=" & Request.ServerVariables("PATH_INFO") & """>"
    sobjShow.HTMLCR "<IMG SRC=""" & Session("htDefaultHiperToolsWorkPath") & "/Code/Vsource.GIF"" WIDTH=85 HEIGHT=45 ALT=""View Active Server Page Source"" ALIGN=CENTER BORDER=0></A>"
    sobjShow.HTMLCR "</td>"
  End If

  sobjShow.HTMLCR "<td align=center><font size=-3>"
  sobjShow.HTMLCR Session("FooterTitle") & "<br>" & Session("FooterSugestions") & _
             "<br>" & Session("FooterCopyright")

  
  sobjShow.HTMLCR "</font></td></tr></table>"
  
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get BaseFont() As String
  BaseFont = Session("BaseFont")
End Property
Public Property Let BaseFont(strValue As String)
  Session("BaseFont") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get BodyALink() As String
  BodyALink = Session("BodyALink")
End Property

Public Property Let BodyALink(strValue As String)
  Session("BodyALink") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get BodyBackground() As String
  BodyBackground = Session("BodyBackground")
End Property

Public Property Let BodyBackground(strValue As String)
  Session("BodyBackground") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get BodyBGColor() As String
  BodyBGColor = Session("BodyBGColor")
End Property

Public Property Let BodyBGColor(strValue As String)
  Session("BodyBGColor") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get BodyBGProperties() As String
  BodyBGProperties = Session("BodyBGProperties")
End Property

Public Property Let BodyBGProperties(strValue As String)
  Session("BodyBGProperties") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get BodyLeftMargin() As Integer
  BodyLeftMargin = Session("BodyLeftMargin")
End Property

Public Property Let BodyLeftMargin(strValue As Integer)
  Session("BodyLeftMargin") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get BodyLink() As String
  BodyLink = Session("BodyLink")
End Property

Public Property Let BodyLink(strValue As String)
  Session("BodyLink") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get BodyStyle() As String
  BodyStyle = Session("BodyStyle")
End Property

Public Property Let BodyStyle(strValue As String)
  Session("BodyStyle") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get BodyText() As String
  BodyText = Session("BodyText")
End Property

Public Property Let BodyText(strValue As String)
  Session("BodyText") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get BodyTopMargin() As String
  BodyTopMargin = Session("BodyTopMargin")
End Property

Public Property Let BodyTopMargin(strValue As String)
  Session("BodyTopMargin") = strValue
End Property


Rem =========================================================================
Rem define o top e left margin
Rem -------------------------------------------------------------------------
Public Property Let BodyMargin(strValue As String)
  BodyTopMargin = strValue
  BodyLeftMargin = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get BodyVLink() As String
  BodyVLink = Session("BodyVLink")
End Property

Public Property Let BodyVLink(strValue As String)
  Session("BodyVLink") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get BodyWidth() As String
  BodyWidth = Session("BodyWidth")
End Property

Public Property Let BodyWidth(intValue As String)
  Session("BodyWidth") = intValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get FooterCopyright() As String
  FooterCopyright = Session("FooterCopyright")
End Property

Public Property Let FooterCopyright(strValue As String)
  Session("FooterCopyright") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get FooterSugestions() As String
  FooterSugestions = Session("FooterSugestions")
End Property

Public Property Let FooterSugestions(strValue As String)
  Session("FooterSugestions") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get FooterTitle() As String
  FooterTitle = Session("FooterTitle")
End Property

Public Property Let FooterTitle(strValue As String)
  Session("FooterTitle") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get HeaderIcon() As String
  HeaderIcon = sstrHeaderIcon
End Property

Public Property Let HeaderIcon(strValue As String)
  sstrHeaderIcon = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get HeaderTitle() As String
  HeaderTitle = sstrHeaderTitle
End Property

Public Property Let HeaderTitle(strValue As String)
  sstrHeaderTitle = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get LinkCopyrightHRef() As String
  LinkCopyrightHRef = Session("LinkCopyrightHRef")
End Property

Public Property Let LinkCopyrightHRef(strValue As String)
  Session("LinkCopyrightHRef") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get LinkCopyrightTitle() As String
  LinkCopyrightTitle = Session("LinkCopyrightTitle")
End Property

Public Property Let LinkCopyrightTitle(strValue As String)
  Session("LinkCopyrightTitle") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get LinkHomeHRef() As String
  LinkHomeHRef = Session("LinkHomeHRef")
End Property

Public Property Let LinkHomeHRef(strValue As String)
  Session("LinkHomeHRef") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get LinkHomeTitle() As String
  LinkHomeTitle = Session("LinkHomeTitle")
End Property

Public Property Let LinkHomeTitle(strValue As String)
  Session("LinkHomeTitle") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get LinkMadeHRef() As String
  LinkMadeHRef = Session("LinkMadeHRef")
End Property

Public Property Let LinkMadeHRef(strValue As String)
  Session("LinkMadeHRef") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get LinkMadeTitle() As String
  LinkMadeTitle = Session("LinkMadeTitle")
End Property

Public Property Let LinkMadeTitle(strValue As String)
  Session("LinkMadeTitle") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get LinkStyleSheetHRef() As String
  LinkStyleSheetHRef = Session("LinkStyleSheetHRef")
End Property

Public Property Let LinkStyleSheetHRef(strValue As String)
  Session("LinkStyleSheetHRef") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaAuthor() As String
  MetaAuthor = Session("MetaAuthor")
End Property

Public Property Let MetaAuthor(strValue As String)
  Session("MetaAuthor") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaCategory() As String
  MetaCategory = Session("MetaCategory")
End Property

Public Property Let MetaCategory(strValue As String)
  Session("MetaCategory") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaCharSet() As String
  MetaCharSet = Session("MetaCharSet")
End Property

Public Property Let MetaCharSet(strValue As String)
  Session("XxXxXx") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaContentType() As String
  MetaContentType = Session("MetaContentType")
End Property

Public Property Let MetaContentType(strValue As String)
  Session("MetaContentType") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaDescription() As String
  MetaDescription = Session("MetaDescription")
End Property

Public Property Let MetaDescription(strValue As String)
  Session("MetaDescription") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaDistribution() As String
  MetaDistribution = Session("MetaDistribution")
End Property

Public Property Let MetaDistribution(strValue As String)
  Session("MetaDistribution") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaExpires() As String
  MetaExpires = Session("MetaExpires")
End Property

Public Property Let MetaExpires(strValue As String)
  Session("MetaExpires") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaFormatter() As String
  MetaFormatter = Session("MetaFormatter")
End Property

Public Property Let MetaFormatter(strValue As String)
  Session("MetaFormatter") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaGenerator() As String
  MetaGenerator = Session("MetaGenerator")
End Property

Public Property Let MetaGenerator(strValue As String)
  Session("MetaGenerator") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaKeyPhrase() As String
  MetaKeyPhrase = Session("MetaKeyPhrase")
End Property

Public Property Let MetaKeyPhrase(strValue As String)
  Session("MetaKeyPhrase") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaKeywords() As String
  MetaKeywords = Session("MetaKeywords")
End Property

Public Property Let MetaKeywords(strValue As String)
  Session("MetaKeywords") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaLocale() As String
  MetaLocale = Session("MetaLocale")
End Property

Public Property Let MetaLocale(strValue As String)
  Session("MetaLocale") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaPicsLabel() As String
  MetaPicsLabel = Session("MetaPicsLabel")
End Property

Public Property Let MetaPicsLabel(strValue As String)
  Session("MetaPicsLabel") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaPragma() As String
  MetaPragma = Session("MetaPragma")
End Property

Public Property Let MetaPragma(strValue As String)
  Session("MetaPragma") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaProduct() As String
  MetaProduct = Session("MetaProduct")
End Property

Public Property Let MetaProduct(strValue As String)
  Session("MetaProduct") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaRefresh() As String
  MetaRefresh = Session("MetaRefresh")
End Property

Public Property Let MetaRefresh(strValue As String)
  Session("MetaRefresh") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaReplyTo() As String
  MetaReplyTo = Session("MetaReplyTo")
End Property

Public Property Let MetaReplyTo(strValue As String)
  Session("MetaReplyTo") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MetaTechnology() As String
  MetaTechnology = Session("MetaTechnology")
End Property

Public Property Let MetaTechnology(strValue As String)
  Session("MetaTechnology") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get SiteCounter() As Long
  SiteCounter = Session("SiteCounter")
End Property

Public Property Let SiteCounter(lngValue As Long)
  Session("SiteCounter") = lngValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get SourceCode() As Boolean
  SourceCode = Session("SourceCode")
End Property

Public Property Let SourceCode(blnValue As Boolean)
  Session("SourceCode") = blnValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get TableBorder() As String
  TableBorder = Session("TableBorder")
End Property

Public Property Let TableBorder(strValue As String)
  Session("TableBorder") = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get UserID() As String
  UserID = Session("UserID")
End Property

Public Property Let UserID(strValue As String)
  Session("UserID") = strValue
End Property

Rem =========================================================================
Rem Propriedade de ativar e desativar a apresentação do User ID
Rem -------------------------------------------------------------------------
Public Property Get HideUserID() As Boolean
  HideUserID = Session("HTHideUserID")
End Property

Public Property Let HideUserID(blnValue As Boolean)
  Session("HTHideUserID") = blnValue
End Property
    
Rem =========================================================================
Rem Propriedade de ativar e desativar a apresentação do Informação do Script Corrente
Rem -------------------------------------------------------------------------
Public Property Get HideFileInfo() As Boolean
  HideFileInfo = Session("HTHideFileInfo")
End Property

Public Property Let HideFileInfo(blnValue As Boolean)
  Session("HTHideFileInfo") = blnValue
End Property
    
Rem =========================================================================
Rem Propriedade de ativar e desativar a apresentação do contador do site
Rem -------------------------------------------------------------------------
Public Property Get HideSiteCounter() As Boolean
  HideSiteCounter = Session("HTHideSiteCounter")
End Property

Public Property Let HideSiteCounter(blnValue As Boolean)
  Session("HTHideSiteCounter") = blnValue
End Property
    
Rem =========================================================================
Rem Propriedade de ativar e desativar a apresentação das outras informações
Rem -------------------------------------------------------------------------
Public Property Get HideInnerInfo() As Boolean
  HideInnerInfo = Session("HTHideInnerInfo")
End Property

Public Property Let HideInnerInfo(blnValue As Boolean)
  Session("HTHideInnerInfo") = blnValue
End Property
   
Rem =========================================================================
Rem Propriedade de ativar e desativar a apresentação do usuário licensiador
Rem -------------------------------------------------------------------------
Public Property Get HideLicenserName() As Boolean
  HideLicenserName = Session("HTHideLicenserName")
End Property

Public Property Let HideLicenserName(blnValue As Boolean)
  Session("HTHideLicenserName") = blnValue
End Property
   
Rem =========================================================================
Rem Propriedade de ativar e desativar a apresentação do Meta no Head
Rem -------------------------------------------------------------------------
Public Property Get HideHeadMeta() As Boolean
  HideHeadMeta = Session("HTHideHeadMeta")
End Property

Public Property Let HideHeadMeta(blnValue As Boolean)
  Session("HTHideHeadMeta") = blnValue
End Property
   
Rem =========================================================================
Rem Propriedade de ativar e desativar a apresentação do Base Font no Head
Rem -------------------------------------------------------------------------
Public Property Get HideHeadBaseFont() As Boolean
  HideHeadBaseFont = Session("HTHideHeadBaseFont")
End Property

Public Property Let HideHeadBaseFont(blnValue As Boolean)
  Session("HTHideHeadBaseFont") = blnValue
End Property

Rem =========================================================================
Rem Propriedade de ativar e desativar a apresentação do Link no Head
Rem -------------------------------------------------------------------------
Public Property Get HideHeadLink() As Boolean
  HideHeadLink = Session("HTHideHeadLink")
End Property

Public Property Let HideHeadLink(blnValue As Boolean)
  Session("HTHideHeadLink") = blnValue
End Property

