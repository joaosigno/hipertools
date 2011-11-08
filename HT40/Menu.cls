VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =========================================================================
' Menu.cls
' -------------------------------------------------------------------------
' Descricao   : Include de Edicao
' Criaçao     : 11:23h 7/2/1998
' Local       : Brasilia/DF
' Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------

Option Explicit

Private sblnVDHTML
Private sstrIMGSrc
Private sstrIMGAlt
Private sstrAHREF
Private sstrATarget
Private sintIMGWidth
Private sintIMGHeight
Private sstrIMGRoot

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
  
  TimeBombX xtbMenu, 15656, 12934, 3080, 14624
  
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
Rem Inicio do menu
Rem -------------------------------------------------------------------------
Public Sub MenuBegin()

  If sblnVDHTML Then
    Response.Write "<script language=""JavaScript"">"
    Response.Write "<" + "!--"
    
  End If
  
End Sub

Rem =========================================================================
Rem Inicio do menu
Rem -------------------------------------------------------------------------
Public Sub MenuEnd()

  If sblnVDHTML Then
    Response.Write "//--" + ">"
    Response.Write "</script>"
    
  End If
  
End Sub

Rem =========================================================================
Rem Inicializacao dos botoes
Rem -------------------------------------------------------------------------
Public Sub ButtonInit(ByVal strIMGExt As String, _
                      ByVal strIMGUp As String, _
                      ByVal strIMGOver As String, _
                      ByVal strIMGOut As String, _
                      ByVal strIMGDown As String)
  Dim strResult

  Session("MenuIMGExt") = strIMGExt
  Session("MenuIMGUp") = strIMGUp
  
  If sblnVDHTML Then
    strResult = ""
  
    If strIMGExt > "" Then strResult = """" & strIMGExt & """"
    If strIMGUp > "" Then strResult = strResult & ", """ & strIMGUp & """"
    If strIMGOver > "" Then strResult = strResult & ", """ & strIMGOver & """"
    If strIMGOut > "" Then strResult = strResult & ", """ & strIMGOut & """"
    If strIMGDown > "" Then strResult = strResult & ", """ & strIMGDown & """"

    Response.Write "InitButton(" & strResult & ");"
    
  End If
  
End Sub

Rem =========================================================================
Rem Inicializacao dos botoes
Rem -------------------------------------------------------------------------
Public Sub ButtonAdd(ByVal strIMGSrc As String, _
                     ByVal strIMGAlt As String, _
                     ByVal strAHRef As String, _
                     ByVal strATarget As String, _
                     ByVal intIMGWidth As Integer, _
                     ByVal intIMGHeight As Integer, _
                     ByVal strIMGRoot As String)
  Dim strResult
  
  strResult = ""

  If sblnVDHTML Then
    If strIMGSrc > "" Then strResult = """" & strIMGSrc & """"
    If strIMGAlt > "" Then strResult = strResult & ", """ & strIMGAlt & """"
    If strAHRef > "" Then strResult = strResult & ", """ & strAHRef & """"
    If strATarget > "" Then
      If LCase(strATarget) = "null" Then
        strResult = strResult & ", ""null"""
        
      Else
        strResult = strResult & ", """ & strATarget & """"
        
      End If
    End If

    If intIMGWidth > "" Then strResult = strResult & ", " & intIMGWidth
    If intIMGHeight > "" Then strResult = strResult & ", " & intIMGHeight
    If strIMGRoot > "" Then strResult = strResult & ", """ & strIMGRoot & """"

    Response.Write "AddNewButton(" & strResult & ");"

  Else
    If strIMGSrc > "" Then
      sstrIMGSrc = strIMGSrc
      
    Else
      strIMGSrc = sstrIMGSrc

    End If
      
    If strIMGAlt > "" Then
      sstrIMGAlt = strIMGAlt
      
    Else
      strIMGAlt = sstrIMGAlt
      
    End If

    If strAHRef > "" Then
      sstrAHREF = strAHRef
      
    Else
      strAHRef = sstrAHREF
      
    End If

    If strATarget > "" Then
      sstrATarget = strATarget

    Else
      strATarget = sstrATarget
      
    End If

    If intIMGWidth > "" Then
      sintIMGWidth = intIMGWidth
      
    Else
      intIMGWidth = sintIMGWidth
      
    End If

    If intIMGHeight > "" Then
      sintIMGHeight = intIMGHeight
      
    Else
      intIMGHeight = sintIMGHeight
      
    End If

    If strIMGRoot > "" Then
      sstrIMGRoot = strIMGRoot
      
    Else
      strIMGRoot = sstrIMGRoot
      
    End If

    If strIMGSrc > "" Then strIMGSrc = " src=""" & strIMGRoot & strIMGSrc & Session("MenuIMGUp") & Session("MenuIMGExt") & """"
    If strIMGAlt > "" Then strIMGAlt = " alt=""" & strIMGAlt & """"
    If strAHRef > "" Then strAHRef = " HREF=""" & strAHRef & """"
    If strATarget > "" Then strATarget = " TARGET=""" & strATarget & """"
    If intIMGWidth > "" Then intIMGWidth = " WIDTH=" & intIMGWidth
    If intIMGHeight > "" Then intIMGHeight = " HEIGHT=" & intIMGHeight

    strResult = "<A" & strAHRef & strATarget & ">"
    strResult = strResult & "<IMG" & strIMGSrc & intIMGWidth & intIMGHeight & " BORDER=0>"
    strResult = strResult & "</A>"

    Response.Write strResult
    
  End If
  
End Sub

Rem =========================================================================
Rem Inicializacao dos botoes
Rem -------------------------------------------------------------------------
Public Sub ButtonHTML(ByVal strText As String)

  If sblnVDHTML Then
    Response.Write "HTMLButton(""" & strText & """);"

  Else
    Response.Write strText
  
  End If
  
End Sub
