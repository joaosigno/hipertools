VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Show"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =========================================================================
' Show.cls
' -------------------------------------------------------------------------
' Descricao   : Include de Shows
' Criaçao     : 11:23h 7/2/1998
' Local       : Brasilia/DF
' Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
' Alterado    : Flavio Costa <flavio@zevallos.com.br>
' Alteracao   : 24/02/2000
' Versao      : 2.0.0
' Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------
Option Explicit

Private ScriptContext As ScriptingContext

Private sstrComments As String, _
    sstrTitle As String, _
    sstrVersion As String, _
    sstrCompanyName As String, _
    sstrProductName As String, _
    sstrLegalCopyright As String, _
    sblnHTMLEncode As Boolean, _
    sblnCrLf As Boolean, _
    sblnShowComments As Boolean, _
    sblnCrLfBefore As Boolean, _
    sblnCrLfAfter As Boolean, _
    sblnAutoWrite As Boolean, _
    sstrListTag As String, _
    sstrErrorColor As String, _
    sstrAlertColor As String
    
Private Request As Request, _
    Response As Response, _
    Server As Server, _
    Session As Session, _
    Application As Application

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  TimeBombX xtbShow, 15656, 12934, 3080, 14624
 
  Set ScriptContext = ScriptContext
  
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application
  
  '------------------------------------------------------------------------------
  ' valores padrao
  '------------------------------------------------------------------------------
  sblnShowComments = True
  sblnCrLf = True
  sblnCrLfBefore = False
  sblnCrLfAfter = False
  sblnHTMLEncode = False
  sblnAutoWrite = True
  sstrErrorColor = "Red"
  sstrAlertColor = "Orange"

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

Public Sub OnEndPage()

  Set ScriptContext = Nothing
  
  Set Request = Nothing
  Set Response = Nothing
  Set Server = Nothing
  Set Session = Nothing
  Set Application = Nothing

End Sub

Public Property Get Comments() As String
  Comments = sstrComments
End Property

Public Property Get Title() As String
  Title = sstrTitle
End Property

Public Property Get Version() As String
  Version = sstrVersion
End Property

Public Property Get CompanyName() As String
  CompanyName = sstrCompanyName
End Property

Public Property Get ProductName() As String
  ProductName = sstrProductName
End Property

Public Property Get LegalCopyright() As String
  LegalCopyright = sstrLegalCopyright
End Property

Rem =========================================================================
Rem Codifica o código como HTML dependendo do valor de HTMLEncode
Rem -------------------------------------------------------------------------
Public Function EncodedHTML(Optional ByVal strHTML As String = "") As String
    If sblnHTMLEncode Then strHTML = Server.HTMLEncode(strHTML)
    EncodedHTML = strHTML
End Function

Rem =========================================================================
Rem Especifica se as apresentações utilizarão o Server.HTMLEncode
Rem -------------------------------------------------------------------------
Public Property Get HTMLEncode() As Boolean
    HTMLEncode = sblnHTMLEncode
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let HTMLEncode(ByVal blnParam As Boolean)
  sblnHTMLEncode = blnParam
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let MessageErrorColor(ByVal strColor As String)
  sstrErrorColor = strColor
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MessageErrorColor() As String
  If sstrErrorColor > "" Then
    MessageErrorColor = sstrErrorColor
  Else
    MessageErrorColor = "red"
  End If
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let MessageAlertColor(ByVal strColor As String)
  sstrAlertColor = strColor
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MessageAlertColor() As String
  If sstrErrorColor > "" Then
    MessageAlertColor = sstrErrorColor
  Else
    MessageAlertColor = "orange"
  End If
End Property


Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get AutoWrite() As Boolean
  AutoWrite = sblnAutoWrite
    
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let AutoWrite(ByVal blnAutoWrite As Boolean)
  sblnAutoWrite = blnAutoWrite
    
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get CrLfBefore() As Boolean
  CrLfBefore = sblnCrLfBefore
    
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let CrLfBefore(ByVal blnCrLf As Boolean)
  sblnCrLfBefore = blnCrLf
    
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get CrLfAfter() As Boolean
  CrLfAfter = sblnCrLfAfter
    
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let CrLfAfter(ByVal blnCrLf As Boolean)
  sblnCrLfAfter = blnCrLf
    
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get CrLf() As Boolean
  CrLf = sblnCrLf
    
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let CrLf(ByVal blnCrLf As Boolean)
  sblnCrLf = blnCrLf
    
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ShowComments() As String
  ShowComments = sblnShowComments
  
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let ShowComments(ByVal blnComment As String)
  sblnShowComments = blnComment
  
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function HTMLNoEncode(ByVal strHTML As String) As String
    Dim blnHTMLEncode As Boolean
    
    blnHTMLEncode = sblnHTMLEncode
    sblnHTMLEncode = False
    
    HTMLNoEncode = HTML(strHTML)
    
    sblnHTMLEncode = blnHTMLEncode
    
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function HTML(ByVal strHTML As String) As String
    
    strHTML = EncodedHTML(strHTML)
    
    If sblnCrLf Then
        If sblnCrLfBefore Then strHTML = vbCrLf & strHTML
        If sblnCrLfAfter Then strHTML = strHTML & vbCrLf
    
    End If
  
    If sblnAutoWrite Then Response.Write strHTML
    
    HTML = strHTML
    
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function HTMLCR(ByVal strHTML As String) As String
    Dim blnLocalCrLfBefore As Boolean, _
        blnLocalCrLfAfter As Boolean
  
    blnLocalCrLfBefore = sblnCrLfBefore
    blnLocalCrLfAfter = sblnCrLfAfter
    sblnCrLfBefore = False
    sblnCrLfAfter = True
    
    HTMLCR = HTML(strHTML)
    
    sblnCrLfBefore = blnLocalCrLfBefore
    sblnCrLfAfter = blnLocalCrLfAfter

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function CRHTML(ByVal strHTML As String) As String
    Dim blnLocalCrLfBefore As Boolean, _
        blnLocalCrLfAfter As Boolean
  
    blnLocalCrLfBefore = sblnCrLfBefore
    blnLocalCrLfAfter = sblnCrLfAfter
    sblnCrLfBefore = True
    sblnCrLfAfter = False
    
    CRHTML = HTML(strHTML)
    
    sblnCrLfBefore = blnLocalCrLfBefore
    sblnCrLfAfter = blnLocalCrLfAfter

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function CRHTMLCR(ByVal strHTML As String) As String
    Dim blnLocalCrLfBefore As Boolean, _
        blnLocalCrLfAfter As Boolean
  
    blnLocalCrLfBefore = sblnCrLfBefore
    blnLocalCrLfAfter = sblnCrLfAfter
    sblnCrLfBefore = True
    sblnCrLfAfter = True
    
    CRHTMLCR = HTML(strHTML)
    
    sblnCrLfBefore = blnLocalCrLfBefore
    sblnCrLfAfter = blnLocalCrLfAfter

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function HTMLComment(ByVal strHTML As String) As String
  If sblnShowComments Then HTMLComment = HTMLNoEncode("<!-- " & strHTML & " -->")
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Message(ByVal strMessage As String, Optional ByVal strColor As String = "") As String
    strMessage = EncodedHTML(strMessage)
    
    If strColor > "" Then
        strMessage = "<FONT COLOR=""" & strColor & """>" & strMessage & "</FONT>"
    
    End If
 
    strMessage = "<BR>" & strMessage
    Message = HTMLNoEncode(strMessage)
    
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function MessageError(ByVal strMessage As String)
  MessageError = Message(strMessage, sstrErrorColor)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function MessageAlert(ByVal strMessage As String)
  MessageAlert = Message(strMessage, sstrAlertColor)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function JSAlert(ByVal strMessage As String) As String
    Dim strReturn As String
    
    If Not strMessage > "" Then strMessage = "[vazio]"
        
    strReturn = "<SCRIPT TYPE=""text/javascript"" LANGUAGE=""JAVASCRIPT"">alert(""" & strMessage & """);</SCRIPT>"
    
    JSAlert = HTMLNoEncode(strReturn)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub AllServerVariables()
  Dim Key

  HTMLCR "<TABLE BORDER=1>"
  HTMLCR "<TR><TD VALIGN=TOP><B>Variable</B></TD><TD VALIGN=TOP><B>Value</B></TD></TR>"
  
  For Each Key In Request.ServerVariables
    HTMLCR "<TR><TD>(" & Key & "</TD>"
    HTMLCR "<TD>"
  
    If Request.ServerVariables(Key) = "" Then
      HTMLCR "&nbsp;"
      
    Else
      HTMLCR Request.ServerVariables(Key)
      
    End If
    HTMLCR "</TD></TR>"
  
  Next

  HTMLCR "</TABLE>"

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Link(ByVal strTag As String, _
                ByVal strName As String, _
                Optional ByVal strHRef As String = "", _
                Optional ByVal strTitle As String = "", _
                Optional ByVal strType As String = "") As String
  
    Dim strResult As String
  
    strResult = "<LINK"
  
    If strTag > "" Then
        strResult = strResult & " " & strTag
        If strName > "" Then strResult = strResult & "=""" & strName & """"
    
    End If

    If strHRef > "" Then strResult = strResult & " HREF=""" & strHRef & """"
    If strTitle > "" Then strResult = strResult & " TITLE=""" & strTitle & """"
    If strType > "" Then strResult = strResult & " TYPE=""" & strType & """"
    strResult = strResult & ">"

    Link = HTMLNoEncode(strResult)
  
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function LinkRev(ByVal strName As String, ByVal strHRef As String, ByVal strTitle As String) As String
    LinkRev = Link("REV", strName, strHRef, strTitle, "")
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function LinkRel(ByVal strName As String, ByVal strHRef As String, ByVal strTitle As String) As String
    LinkRel = Link("REL", strName, strHRef, strTitle, "")
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function LinkRelStyleSheet(ByVal strHRef As String) As String
    LinkRelStyleSheet = Link("REL", "STYLESHEET", strHRef, "", "text/css")
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Meta(ByVal strField As String, ByVal strName As String, ByVal strContent As String) As String
    Dim strResult As String
  
    strResult = "<META " & strField & "=""" & strName & """ CONTENT="

    If InStr("'", strContent) = 0 Then
        strResult = strResult & """" & strContent & """" & ">"
    
    Else
        strResult = strResult & "'" & strContent & "'" & ">"
    
    End If
  
    Meta = HTMLNoEncode(strResult)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function MetaName(ByVal strName As String, ByVal strContent As String) As String
    MetaName = Meta("NAME", strName, strContent)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function MetaHTTPEquiv(ByVal strName As String, ByVal strContent As String) As String
    MetaHTTPEquiv = Meta("HTTP-EQUIV", strName, strContent)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function ItemListBegin(Optional intType As Integer = shListTypeDisc) As String
    ItemListBegin = BeginItemList(intType)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function BeginItemList(Optional intType As Integer = shListTypeDisc) As String
    Dim strType As String, _
        strReturn As String

    Select Case intType
    
        Case shListTypeDisc
            sstrListTag = "UL"
            strType = "disc"
        
        Case shListTypeCircle
            sstrListTag = "UL"
            strType = "circle"
    
        Case shListTypeSquare
            sstrListTag = "UL"
            strType = "square"
    
        Case shListTypeNumber
            sstrListTag = "OL"
            strType = "1"
    
        Case shListTypeUChar
            sstrListTag = "OL"
            strType = "A"
    
        Case shListTypeLChar
            sstrListTag = "OL"
            strType = "a"
        
        Case shListTypeURoman
            sstrListTag = "OL"
            strType = "I"
    
        Case shListTypeLRoman
            sstrListTag = "OL"
            strType = "i"
    
    End Select

    strReturn = "<" & sstrListTag & " TYPE=""" & strType & """>"
    BeginItemList = HTMLNoEncode(strReturn)

End Function


Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function EndItemList() As String
  EndItemList = HTMLNoEncode("</" & sstrListTag & ">")
End Function


Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function ItemListEnd() As String
  ItemListEnd = HTMLNoEncode("</" & sstrListTag & ">")
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function ItemList(ByVal strTitle As String, Optional ByVal strContent As String, Optional ByVal strSeparator As String = ": ", Optional ByVal strColor As String = "") As String
    Dim strReturn As String

    strReturn = EncodedHTML(strTitle & strSeparator & strContent)
    
    If strColor > "" Then
      strReturn = "<FONT COLOR=""" & strColor & """>" & strReturn & "</FONT>"
    
    End If
    
    strReturn = "<LI>" & strReturn & "</LI>"
    
    ItemList = HTMLNoEncode(strReturn)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function ItemListError(ByVal strTitle As String, ByVal strContent As String) As String
  ItemListError = ItemList(strTitle, strContent, , sstrErrorColor)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Center() As String
  Center = HTMLNoEncode("<CENTER>")
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function CenterEnd() As String
  CenterEnd = HTMLNoEncode("</CENTER>")
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Nbsp(Optional ByVal intTimes As Integer = 1) As String
  Nbsp = HTMLNoEncode(HTReplicate("&nbsp;", intTimes))
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function BR(Optional ByVal intTimes As Integer = 1) As String
  BR = HTMLNoEncode(HTReplicate("<BR>", intTimes))
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Paragraph() As String
  Paragraph = HTMLNoEncode("<P>")
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function BoldText(ByVal strText As String) As String
  BoldText = Bold & HTML(strText) & BoldEnd
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function ItalicText(ByVal strText As String) As String
    ItalicText = Italic & HTML(strText) & ItalicEnd
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function UnderlineText(ByVal strText As String) As String
    UnderlineText = Underline & HTML(strText) & UnderlineEnd
End Function
                
Rem =========================================================================
Rem Alterado (Ridai Govinda)
Rem -------------------------------------------------------------------------
Public Function Bold() As String
  Bold = HTMLNoEncode("<B>")
End Function
                
Rem =========================================================================
Rem Alterado (Ridai Govinda)
Rem -------------------------------------------------------------------------
Public Function BoldEnd() As String
  BoldEnd = HTMLNoEncode("</B>")
End Function
                
Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Italic() As String
  Italic = HTMLNoEncode("<I>")
End Function
                
Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function ItalicEnd() As String
  ItalicEnd = HTMLNoEncode("</I>")
End Function
                
Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Underline() As String
  Underline = HTMLNoEncode("<U>")
End Function
                
Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function UnderlineEnd() As String
    UnderlineEnd = HTMLNoEncode("</U>")
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Font(Optional ByVal intSize As Integer = 0, Optional ByVal strColor As String = "", Optional ByVal strFace As String = "") As String
  Dim strFont As String

  strFont = "<FONT"
  If intSize > 0 Then strFont = strFont & " SIZE=""" & intSize & """"
  If strColor > "" Then strFont = strFont & " COLOR=""" & strColor & """"
  If strFace > "" Then strFont = strFont & " FACE=""" & strFace & """"
  strFont = strFont & ">"

  Font = HTMLNoEncode(strFont)

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function FontEnd() As String
  FontEnd = HTMLNoEncode("</FONT>")
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Image(ByVal strImagem As String, Optional ByVal strID As String = "", _
                 Optional ByVal strAlt As String = "", Optional ByVal strHRef As String = "", _
                 Optional ByVal strAlign As String = "", Optional ByVal intWidth As String = "", _
                 Optional ByVal intHeight As String = "", Optional ByVal intBorder As Integer = 0, _
                 Optional ByVal strTarget As String = "") As String
                
    Dim strImage As String
                
    If strHRef > "" Then
      strImage = "<A HREF=""" & strHRef & """"
      If strTarget > "" Then strImage = strImage & " TARGET=""" & strTarget & """"
      strImage = strImage & ">"
    End If
    
    strImage = strImage & "<IMG SRC=""" & strImagem & """ BORDER=""" & intBorder & """"
    If strAlt > "" Then strImage = strImage & " ALT=""" & strAlt & """"
    If strAlign > "" Then strImage = strImage & " ALIGN=""" & strAlign & """"
    If strID > "" Then strImage = strImage & " ID=""" & strID & """"
    If intWidth > "" Then strImage = strImage & " WIDTH=""" & intWidth & """"
    If intHeight > "" Then strImage = strImage & " HEIGHT=""" & intHeight & """"
    strImage = strImage & ">"
    If strHRef > "" Then strImage = strImage & "</A>"

    Image = HTMLNoEncode(strImage)

End Function

Rem =========================================================================
Rem Monta um parametro de uma TAG HTML, mas mostra somente se o value tiver algo
Rem -------------------------------------------------------------------------
Public Sub TagParameter(ByVal strParam As String, _
                        Optional ByVal strValue As String = "", _
                        Optional ByVal blnEqual As Boolean = True)
  
  strValue = Trim(strValue)
  
  If strValue > "" And blnEqual Then
    Response.Write " " & strParam

    If blnEqual Then
      
      HTML "=""" & strValue & """"
    End If
  ElseIf Not blnEqual Then
    HTML " " & strParam
  
  End If
End Sub


Rem =========================================================================
Rem  Monta uma Classe de estilo se os parametros tiverem algo
Rem -------------------------------------------------------------------------
Public Sub StyleClass(ByVal strClass As String, ByVal strParam As String)
  strParam = Trim(strParam)
  If strParam > "" Then HTML strClass & "{" & RemoveBrackets(strParam) & "}"
End Sub

Rem =========================================================================
Rem Remove as chaves dos estilos caso o usuário as tenha colocado.
Rem -------------------------------------------------------------------------
Public Function RemoveBrackets(ByVal strStyleKey As String) As String
  RemoveBrackets = HTRemoveChar(strStyleKey, "{", "}")
End Function

