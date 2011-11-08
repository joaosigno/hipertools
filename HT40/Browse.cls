VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Browse"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =========================================================================
' Browse.cls
' -------------------------------------------------------------------------
' Descricao   : Include do sistema de Browse
' Criacao     : 9/21/98 5:59AM
' Local       : Brasilia/DF
' Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 1998, 1999 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------

'    Description = "HiperTools.Browse - Sistema de Browse - 1/31/99 10:47AM"
'    ProgId = "HiperTools.Browse"
'    Version = "1.01"
'    ClassID = "{E0408FF0-B920-11d2-B163-00A0C923ECAE}"

Option Explicit

Rem =========================================================================
Rem Constantes de Navegacao
Rem -------------------------------------------------------------------------

Const conMaxLenght = 8
 
Const conOption = "htBO"
Const conTarget = "htBT"
Const conCurrentPage = "htBA"

Const conPage = "1"
Const conPageSize = "2"

Const conDefaultPS = 10
Const conBarPages = 10
  
Private sobjShow As HiperTools30.Show
   
Private Request As Request, _
        Response As Response, _
        Server As Server, _
        Session As Session, _
        Application As Application
    
Rem =========================================================================
Rem Pega os parametros passados
Rem -------------------------------------------------------------------------
Private sparOption As String
Private sparTarget As String
Private sparCurrentPage As String

Private sobjRS As Object

Rem =========================================================================
Rem Constantes Locais
Rem -------------------------------------------------------------------------
Private sintLine As Integer

Private ScriptContext As ScriptingContext

Private sstrComments As String, _
        sstrTitle As String, _
        sstrVersion As String, _
        sstrCompanyName As String, _
        sstrProductName As String, _
        sstrLegalCopyright As String

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  TimeBombX xtbBrowse, 15656, 12934, 3080, 14624
  
  Set ScriptContext = ScriptContext
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application
  Set sobjShow = Server.CreateObject("HiperTools30.Show")
  
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

Public Sub Class_Initialize()

    sstrComments = App.Comments
    sstrTitle = App.Title
    sstrVersion = App.Major & "." & _
                App.Minor & "." & _
                App.Revision
    sstrCompanyName = App.CompanyName
    sstrProductName = App.ProductName

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
Rem
Rem -------------------------------------------------------------------------
Public Property Get Page() As Integer
  
  Page = Session("BrowsePage" & Session("BrowseCurrent"))
  
  If Page < 1 Then
    Page = 1
    
  End If
  
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Page(ByVal intPage As Integer)
  
  If Not PageSize > 0 Then
     Session("BrowsePage" & Session("BrowseCurrent")) = conDefaultPS
  Else
     Session("BrowsePage" & Session("BrowseCurrent")) = intPage
  End If

End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get PageSize() As Integer
  PageSize = Session("BrowsePageSize" & Session("BrowseCurrent"))
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let PageSize(ByVal intPageSize As Integer)
  Session("BrowsePageSize" & Session("BrowseCurrent")) = intPageSize
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Actual() As String
  Actual = Current
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Actual(ByVal strCurrent As String)
  Current = strCurrent
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Current() As String
  Current = Session("BrowseCurrent")
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Current(ByVal strCurrent As String)
  
  Session("BrowseCurrent") = strCurrent
  Session("BrowseLast") = strCurrent

  If Not IsBrowse Then
    Status = False
  End If
  
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get QueryString() As String
  QueryString = Session("BrowseQueryString" & Session("BrowseCurrent"))
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let QueryString(ByVal strQueryString As String)
  
  Session("BrowseQueryString" & Session("BrowseCurrent")) = strQueryString
  Session("BrowseQueryStringStatus" & Session("BrowseCurrent")) = True

End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Status() As Boolean
  Status = Session("BrowseActive" & Session("BrowseCurrent"))
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Status(ByVal blnStatus As Boolean)
  Session("BrowseActive" & Session("BrowseCurrent")) = blnStatus
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get RecordSet() As ADODB.RecordSet
  Set RecordSet = sobjRS
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Set RecordSet(ByVal objRS As ADODB.RecordSet)
  Set sobjRS = objRS
  Status = True

  CountBegin

  If Request.ServerVariables("QUERY_STRING") > "" And Not sparOption > "" And Not sparCurrentPage > "" Then
     QueryString = Request.ServerVariables("QUERY_STRING")
  End If
  
  If IsBrowse Then
    
    Select Case sparOption
      Case CStr(conPage)
        Page = CInt(sparTarget)

      Case CStr(conPageSize)
        PageSize = CInt(sparTarget)
        Page = CInt(sparCurrentPage)

      Case Else

    End Select
  
    If Not sobjRS.BOF And Not sobjRS.EOF Then
      
      sobjRS.PageSize = PageSize
  
      If Page > sobjRS.PageCount Then
         Page = sobjRS.PageCount
      ElseIf Page < 1 Then
         Page = 1
      End If

      sobjRS.AbsolutePage = Page

      Status = True
  
    Else
      Status = False
    End If
  End If
  
  Session("PageCount") = sobjRS.PageCount
  Session("RecordCount") = sobjRS.RecordCount

End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get BarPages() As Integer
  
  If Session("BrowseBarPages") <= 1 Then
     Session("BrowseBarPages") = conBarPages
     BarPages = conBarPages
  Else
     BarPages = Session("BrowseBarPages")
  End If
  
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let BarPages(ByVal intBarPages As Integer)
  Session("BrowseBarPages") = intBarPages
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub IncLine()
  sintLine = sintLine + 1
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Line() As Integer
  Line = sintLine
End Property


Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub EndBrowse()
  
   ' Eu hein?
   
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub BeginBrowse(Optional ByVal objRS As ADODB.RecordSet = Nothing)
  
  If IsObject(objRS) Then
     Set RecordSet = objRS
  End If

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub Bar()

   Dim intLastPage As Integer, intFirstPage As Integer
   Dim intLessPages As Integer, intMorePages As Integer
   Dim strAHRef As String, strQuery As String
   Dim i As Integer

  Response.Write "<FONT SIZE=1>"
  
  If IsBrowse And Session("PageCount") > 1 Then
    
    intFirstPage = Int((Page - 1) / BarPages) * BarPages + 1
    intLastPage = intFirstPage + BarPages - 1

    If intLastPage >= Session("PageCount") Then
       intLastPage = Session("PageCount") - intLastPage + BarPages + intFirstPage - 1
    End If

    Rem ---------------------------------------------------------------------
    If Page + BarPages > Session("PageCount") Then
       intMorePages = Session("PageCount") - Page
    Else
       intMorePages = BarPages
    End If

    Rem ---------------------------------------------------------------------
    If Page - BarPages <= 0 Then
       intLessPages = intFirstPage
    Else
       intLessPages = BarPages
    End If

  End If

  Dim strSTail

  strSTail = ""

  If Line > 1 Then
     strSTail = "s"
  End If
  
  Response.Write "<BR>" & Line & " Linha" & strSTail & " apresentada" & strSTail
    
  If Page < Session("PageCount") Then
     Response.Write " de " & Session("RecordCount") & " linhas"
  End If

  strSTail = ""

  If Session("PageCount") > 1 Then
     strSTail = "s"
  End If
  
  If Session("PageCount") > 1 Then
     Response.Write "<BR>na página " & Page
     Response.Write " de " & Session("PageCount") & " página" & strSTail & "<BR>"
  End If
  
  strQuery = LCase(Request.ServerVariables("SCRIPT_NAME")) & "?"

  If IsBrowseQueryString Then
     strQuery = strQuery & QueryString & "&"
  End If

  Response.Write "<form method=POST action=" & strQuery & conOption & "=" & conPageSize & "&" & conCurrentPage & "=" & Page & ">"
  Response.Write "<b>Tamanho da página:</b> <input type=text name=str" & conPageSize & " size=3 value=" & PageSize & ">"
  Response.Write "<input type=submit Name=cmdbrowse value=Atualizar><BR>"
  sobjShow.HTMLCR "</form>"
  
  If IsBrowse And Session("PageCount") > 1 Then
    
    strAHRef = "<A HREF=" & strQuery & conOption & "=" & conPage & "&" & conCurrentPage & "=" & Page & "&" & conTarget & "="
 
    If Page > 1 Then
       AHRef strAHRef, Page - 1, "[Anterior]"
    Else
       sobjShow.HTMLCR "[Anterior]"
    End If

    If Page < Session("PageCount") Then
      AHRef strAHRef, Page + 1, "[Próxima]"
    Else
      sobjShow.HTMLCR "[Próxima]"
    End If

    sobjShow.HTMLCR "&nbsp;"

    If intFirstPage > 1 Then
       AHRef strAHRef, Page - intLessPages, "[Menos " & intLessPages & "]"
    Else
       sobjShow.HTMLCR "&nbsp;[Menos 10]"
    End If

    Response.Write "&nbsp;"

    For i = intFirstPage To Page - 1
      AHRef strAHRef, i, i
    Next

    sobjShow.HTMLCR "<FONT COLOR=Red><B>" & Page & "</B></FONT>"
  
    For i = Page + 1 To intLastPage
      AHRef strAHRef, i, i
    Next

    Response.Write "&nbsp;"
    
    If intLastPage < Session("PageCount") Then
       AHRef strAHRef, Page + intMorePages, "[Mais " & intMorePages & "]"
    Else
       sobjShow.HTMLCR "&nbsp;[Mais 10]"
    End If

    sobjShow.HTMLCR "&nbsp;"
  
    If Page > 1 Then
      AHRef strAHRef, 1, "[Primeira]"
    Else
      sobjShow.HTMLCR "[Primeira]"
    End If
  
    If Page < Session("PageCount") Then
      AHRef strAHRef, Session("PageCount"), "[Última]"
    Else
      sobjShow.HTMLCR "[Última]"
    End If
  
  End If

  Response.Write "</FONT>"

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub AHRef(ByVal strAHRef As String, _
                 ByVal strNumber As String, _
                 ByVal strText As String)
  sobjShow.HTMLCR strAHRef & strNumber & ">" & strText & "</A>"
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function IsBrowse() As Boolean
  
  If Not Current > "" Then
    Back
  End If
  
  IsBrowse = Status

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function IsBrowsePage() As Boolean
  IsBrowsePage = Not EndOfPage

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function EndOfPage() As Boolean
  EndOfPage = (Line >= PageSize)

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Private Sub Back()
  Current = Session("BrowseLast")
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Private Function IsBrowseQueryString() As Boolean
  
  If Not Session("BrowseQueryStringStatus" & Session("BrowseCurrent")) > "" Then
     IsBrowseQueryString = False
     Session("BrowseQueryStringStatus" & Session("BrowseCurrent")) = False
  Else
     IsBrowseQueryString = Session("BrowseQueryStringStatus" & Session("BrowseCurrent"))
  End If

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Private Sub CountBegin()
  
  sparOption = LCase(Request.QueryString(conOption))
  sparTarget = LCase(Request.QueryString(conTarget))
  sparCurrentPage = LCase(Request.QueryString(conCurrentPage))

  If sparOption <> conPageSize And PageSize = 0 Then
    PageSize = conDefaultPS
  End If

  If sparOption = conPageSize Then
    sparTarget = Request.Form("str" & conPageSize)
  End If
  
  sintLine = 0

End Sub

'Criado (Ridai Govinda)
Public Property Get parOption() As String
  parOption = Request.QueryString(conOption)
End Property

Public Property Get parTarget() As String
  parTarget = Request.QueryString(conTarget)
End Property

Public Property Get parCurrent() As String
  parCurrent = Request.QueryString(conCurrentPage)
End Property
'--------------------
