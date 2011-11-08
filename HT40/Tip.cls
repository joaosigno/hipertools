VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Tip"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =========================================================================
' Tip.cls
' -------------------------------------------------------------------------
' Descricao   : Include de Edicao
' Criaçao     : 11:23h 7/2/1998
' Local       : Brasilia/DF
' Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------

Option Explicit

Public TitleClass As String, _
        ContentClass As String, _
        OverColor As String, _
        OutColor As String

Private sobjShow As HiperTools30.Show, _
        sintTipCounter As Integer, _
        sintTipLevel As Integer

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
        ScriptContext As ScriptingContext, _
        Application As Application

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  TimeBombX xtbTip, 15656, 12934, 3080, 14624
  
  Set ScriptContext = ScriptContext
  
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application
 
  Set sobjShow = Server.CreateObject("HiperTools30.Show")
  
  TitleClass = ""
  ContentClass = ""
  OverColor = ""
  OutColor = ""
  sintTipCounter = 1
  sintTipLevel = 1
 
  SetTipHeadClass "dtitle"
  SetTipContentClass "dbody"

  SetTipOverColor "Black"
  SetTipNormalColor "Blue"
 
End Sub

Public Sub OnEndPage()

  Set ScriptContext = Nothing
  
  Set Request = Nothing
  Set Response = Nothing
  Set Server = Nothing
  Set Session = Nothing
  Set Application = Nothing

  Set sobjShow = Nothing

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
Public Sub AddTip(ByVal strTitle As String, ByVal strBody As String)

  Response.Write "<tr><td align=left>"
  Response.Write "<DIV ID=F" & sintTipCounter

  If TitleClass > "" Then
      Response.Write " class=" & TitleClass
      
  End If
  
  Response.Write " onMouseOver=""makeCool(F" & sintTipCounter & ");"" onMouseOut=""makeNormal(F" & sintTipCounter & ");"" onClick=""OpenDesc(T" & sintTipCounter & ");"">" & strTitle & "</DIV>"
  
  Response.Write "</td></tr><tr><td align=left>"
  Response.Write "<DIV ID=T" & sintTipCounter

  If ContentClass > "" Then
      Response.Write " class=" & ContentClass
      
  End If
  
  Response.Write " style=""display:'none'>"
  Response.Write strBody
  Response.Write "<hr color=white></DIV></td></tr>"

  sintTipCounter = sintTipCounter + 1

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub BeginTip()

  sobjShow.HTMLCR "<SCRIPT TYPE=""text/javascript"">" & Chr(13)
  sobjShow.HTMLCR "<!--"
  sobjShow.HTMLCR "function OpenDesc(text) {"
  sobjShow.HTMLCR "  if (text.style.display =="""") {"
  sobjShow.HTMLCR "    text.style.display = ""none"";"
  sobjShow.HTMLCR "  }else {"
  sobjShow.HTMLCR "    text.style.display="""";"
  sobjShow.HTMLCR "  }"
  sobjShow.HTMLCR "}"

  sobjShow.HTMLCR "function makeCool(text) {"
  sobjShow.HTMLCR "  text.style.color = '" & OverColor & "';"
  sobjShow.HTMLCR "  text.style.cursor = 'hand';"
  sobjShow.HTMLCR "}"

  sobjShow.HTMLCR "function makeNormal(text) {"
  sobjShow.HTMLCR "  text.style.color = """ & OutColor & """;"
  sobjShow.HTMLCR "  text.style.cursor = 'default';"
  sobjShow.HTMLCR "}"
  sobjShow.HTMLCR "//-->"
  sobjShow.HTMLCR "</script>"

  Response.Write "<TABLE width=100% border=0 CELLPADDING=0 CELLSPACING=0>"

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub EndTip()

  Response.Write "</TABLE>"

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub SetTipHeadClass(ByVal strClass As String)

  TitleClass = strClass

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub SetTipContentClass(ByVal strClass As String)

  ContentClass = strClass

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub SetTipOverColor(ByVal strColor As String)

  OverColor = strColor

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub SetTipNormalColor(ByVal strColor As String)

  OutColor = strColor

End Sub

Rem =========================================================================
Rem Adiciona um Item na Lista
Rem -------------------------------------------------------------------------
Public Sub AddTipBegin(ByVal strTitle As String, ByVal strBody As String)

  Response.Write "<TR><TD align=left ID=LF" & sintTipLevel & sintTipCounter

  If TitleClass > "" Then
      Response.Write " class=" & TitleClass
      
  End If
  
  Response.Write " onMouseOver=""makeCool(LF" & sintTipLevel & sintTipCounter & ");"" onMouseOut=""makeNormal(LF" & sintTipLevel & sintTipCounter & ");"" onClick=""OpenDesc(LT" & sintTipLevel & sintTipCounter & ");"">" & strTitle
  
  Response.Write "</TD></TR><TR><TD>"
    
  Response.Write "<TABLE border=" & Session("TableBorder") & " width=100% CELLPADDING=0 CELLSPACING=0>"

  Response.Write "<TR><TD>&nbsp;</TD><TD ID=LT" & sintTipLevel & sintTipCounter

  If ContentClass > "" Then
      Response.Write " class=" & ContentClass
      
  End If
  
  Response.Write " style=""display:'none';"">"

  If strBody > "" Then
    Response.Write strBody
    Response.Write "<hr color=white>"
    
  End If

  Response.Write "<TABLE border=" & Session("TableBorder") & " width=100% CELLPADDING=0 CELLSPACING=0>"

  sintTipCounter = sintTipCounter + 1

  sintTipLevel = sintTipLevel + 1

End Sub

Rem =========================================================================
Rem Muda para o nivel anterior
Rem -------------------------------------------------------------------------
Public Sub AddTipEnd()

  sintTipLevel = sintTipLevel - 1

  Response.Write "</TABLE></TD></TR></TABLE></TR></TD>"

End Sub
