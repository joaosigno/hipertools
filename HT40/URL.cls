VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "URL"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------------------------------------------
' URL.cls
'-------------------------------------------------------------------------
' Descricao   : Include de URL
' Criacao     : 11:23h 7/2/1998
' Local       : Brasilia/DF
' Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
'-------------------------------------------------------------------------

Option Explicit

Private sstrAmpersand As String, _
        sstrURLLocation As String, _
        sobjParameters As HiperTools30.Pack, _
        sstrBeforeStr As String, _
        sstrAfterStr As String, _
        sstrBefCloseStr As String, _
        sstrURLTarget As String, _
        sblnURLEncode As Boolean, _
        sobjInitializer As HiperTools30.Initializer
'Criado (Ridai Govinda)
Public Hint As String
'---------------------
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
        Application As Application, ScriptContext As ScriptingContext

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  TimeBombX xtbURL, 15656, 12934, 3080, 14624
  
  Set ScriptContext = ScriptContext
  
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application
  
  Set sobjParameters = Server.CreateObject("HiperTools30.Pack")
  Set sobjInitializer = Server.CreateObject("HiperTools30.Initializer")
  
  Rem --- Alterado ---
  sstrBeforeStr = "<A "
  sstrBefCloseStr = ">"
  sstrAfterStr = "</A>"
  Hint = ""
  
  sblnURLEncode = False
  sstrURLLocation = sobjInitializer.ScriptName
End Sub

Public Sub OnEndPage()

  Set sobjInitializer = Nothing
  Set sobjParameters = Nothing
  
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
Public Property Let URLEncode(ByVal blnParam As Boolean)
  sblnURLEncode = blnParam
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get URLEncode() As Boolean
  URLEncode = sblnURLEncode
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Parameters() As HiperTools30.Pack
Attribute Parameters.VB_UserMemId = 0
  Set Parameters = sobjParameters
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Set Parameters(ByRef objParam As HiperTools30.Pack)
  Set sobjParameters = objParam
End Property

Rem -------------------------------------------------------------------------
Rem Monta URL para chamar o Edit
Rem -------------------------------------------------------------------------
Public Property Get Ampersand() As String
  'Alterado (Ridai)
  If Not sstrAmpersand > "" Then sstrAmpersand = "?"
  
  Ampersand = sstrAmpersand
  If sstrAmpersand = "?" Then
    sstrAmpersand = "&"

  End If
  
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Ampersand(ByVal strValue As String)
  
  If strValue = "?" Or strValue = "&" Then
    sstrAmpersand = strValue
  Else
    If Not strValue > "" Then sstrAmpersand = ""
  
  End If
  
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Location(ByVal strValue As String)
  If Not strValue > "" Then
    strValue = sobjInitializer.ScriptName
  End If
  
  'Alterado (Ridai)
  
  sstrURLLocation = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Location() As String
  Location = sstrURLLocation
End Property

Rem --- Alterado/Criado --- (Ridai Govinda)
Rem -------------------------------------------------------------------------
Rem Especifica o target para o a URL (<A HREF= TARGET=) usado
Rem -------------------------------------------------------------------------
Public Property Let Target(ByVal strString As String)
  Dim arrAux As Variant
  
  If InStr(strString, "=") > 0 Then
    strString = Replace(strString, """", vbNullString)
    
    arrAux = Split(strString, "=")
    
    If IsArray(arrAux) Then
      strString = arrAux(1)
    End If
  End If
  
  sstrURLTarget = strString
  
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Target() As String
  Target = sstrURLTarget
End Property

Rem -------------------------------------------------------------------------
Rem Especifica que Tag fica antes da URL
Rem -------------------------------------------------------------------------
Public Property Let BeforeURL(ByVal strString As String)
    If InStr(strString, "<") > 0 Then sstrBeforeStr = strString
End Property

Rem --- Alterado/Criado --- (Ridai Govinda)
Rem -------------------------------------------------------------------------
Rem Lê a tag antes da URL Lista
Rem -------------------------------------------------------------------------
Public Property Get BeforeURL() As String
    BeforeURL = sstrBeforeStr
End Property

Rem --- Alterado/Criado --- (Ridai Govinda)
Rem -------------------------------------------------------------------------
Rem Especifica a tag após à URL
Rem -------------------------------------------------------------------------
Public Property Let AfterURL(ByVal strString As String)
    If InStr(strString, "</") > 0 Then sstrAfterStr = strString
End Property

Rem --- Alterado/Criado --- (Ridai Govinda)
Rem -------------------------------------------------------------------------
Rem Lê a tag após a URL
Rem -------------------------------------------------------------------------
Public Property Get AfterURL() As String
    AfterURL = sstrAfterStr
End Property

Rem --- Alterado/Criado --- (Ridai Govinda)
Rem -------------------------------------------------------------------------
Rem Especifica como fechar a Tag de URL
Rem -------------------------------------------------------------------------
Public Property Let CloseBeforeURL(ByVal strString As String)
    If InStr(strString, ">") > 0 Then sstrBefCloseStr = strString
End Property

Rem --- Alterado/Criado --- (Ridai Govinda)
Rem -------------------------------------------------------------------------
Rem Lê o que foi adicionado após da URL Lista
Rem -------------------------------------------------------------------------
Public Property Get CloseBeforeURL() As String
    
    Dim strTarget As String
    
    strTarget = ""
    If Target > "" Then
      strTarget = " TARGET=""" & Target & """ "
    End If
    
    CloseBeforeURL = strTarget & sstrBefCloseStr

End Property

Rem -------------------------------------------------------------------------
Rem Inicia o URL
Rem -------------------------------------------------------------------------
Public Sub BeginURL(Optional ByVal strLocation As String = "")

  'Alterado (Ridai)
  Ampersand = ""

  If strLocation > "" Then
    Location = strLocation
  End If

  'sintURLCounter = 0

End Sub

Rem -------------------------------------------------------------------------
Rem Termina o URL
Rem -------------------------------------------------------------------------
Public Sub EndURL()

  Rem --- Alterado--- (Ridai Govinda)
  'sintURLCounter = 0

  Ampersand = ""
  Parameters.Clear

End Sub

Rem --- Alterado --- (Ridai Govinda)
Rem -------------------------------------------------------------------------
Rem Adiciona a URL na Lista
Rem -------------------------------------------------------------------------
Public Sub Add(ByVal strOption As String, _
               ByVal strValue As String)

  Rem --- Alterado --- (Ridai Govinda)
  'sintURLCounter = sintURLCounter + 1
  'sarrURL(sintURLCounter) = strOption & "=" & Server.URLEncode(strValue)
  
  If Not Parameters.Find(strOption) Then
    If InStr(strOption, "=") > 0 Then
      Err.Raise 300, "HiperTools30.URL", "O parâmetro contém o simbolo de igual (""="")"
    Else
      Parameters.Add strOption, strValue
    End If
  Else
    Err.Raise 300, "HiperTools30.URL", "Este parâmetro já existe"
  End If
  Rem ---

End Sub

Rem --- Alterado/Criado --- (Ridai Govinda)
Rem -------------------------------------------------------------------------
Rem Retorna só a URL na Lista
Rem -------------------------------------------------------------------------
Public Function GetOnlyURL(Optional ByVal strOption As String = "") As String
  
  Dim strResult As String, i As Integer, arrAux As Variant
  Dim objPackItem As Object
  
  strResult = Location
  
  If InStr(strResult, "?") = 0 Then Ampersand = ""
  
  If InStr(strOption, "=") = 0 Then
    If strOption > "" Then _
    Err.Raise 300, "HiperTools30.URL", "Não foi há o símbolo de igual (""="") no parâmetro passado"
  End If
  
  If Left(strOption, 1) = "&" Or Left(strOption, 1) = "?" Then
    strOption = Mid(strOption, 2)
    
  End If
  
  If URLEncode And strOption > "" Then
    arrAux = Split(strOption, "=")
    
    If IsArray(arrAux) Then
      arrAux(1) = Server.URLEncode(arrAux(1))
      strOption = Join(arrAux, "=")
    End If
  End If
  
  If strOption > "" Or Parameters.Count > 0 Then
    
    If strOption > "" Then
      strResult = strResult & Ampersand
      strResult = strResult & strOption
      
    End If
  End If

  If Parameters.Count > 0 Then
    
    'If strOption > "" Then
    '  strResult = strResult & Ampersand
      
    'End If
    
    For Each objPackItem In Parameters
      'strResult = strResult & Ampersand

      'strResult = strResult & CStr(objPackItem.Key) & "="
      
      'If URLEncode Then
        'strResult = strResult & Server.URLEncode(CStr(objPackItem))
      'Else
        'strResult = strResult & CStr(objPackItem)
      'End If
      strResult = strResult & Equal(objPackItem.Key, objPackItem.Value)
    Next
    
  End If
  
  GetOnlyURL = strResult
  
End Function

Rem -------------------------------------------------------------------------
Rem Pega URL com tags
Rem -------------------------------------------------------------------------
Public Function GetURL(ByVal strDescription As String, _
                       Optional ByVal strOption As String = "") As String
  
  Const lconOpenHREF = "HREF="""
  Const lconCloseHREF = """"
  
  Dim strResult As String
  
  If Left(strOption, 1) = "&" Then
    strOption = Mid(strOption, 2)
    
  End If

  Rem --- Alterado/Criado --- (Ridai Govinda)
  strResult = BeforeURL & lconOpenHREF & GetOnlyURL(strOption) & lconCloseHREF
  
  If Hint > "" Then strResult = strResult & " TITLE=""" & Hint & """"
  
  strDescription = strDescription
  
  If Not strDescription > "" Then
    strDescription = "_"
    
  End If

  Rem --- Alterado --- (Ridai Govinda)
  strResult = strResult & CloseBeforeURL & strDescription & AfterURL
  
  GetURL = strResult
  Rem ---
End Function

Rem -------------------------------------------------------------------------
Rem Monta URL para chamar o Edit
Rem -------------------------------------------------------------------------
Public Sub Show(ByVal strDescription As String, _
                Optional ByVal strOption As String = "")
  Rem Ampersand = ""
  
  Response.Write GetURL(strDescription, strOption)
End Sub

Rem -------------------------------------------------------------------------
Rem Monta URL para chamar o Edit
Rem -------------------------------------------------------------------------
Public Function Equal(ByVal strOption As String, _
                      ByVal strValue As String) As String
  Rem -- Alterado --- (Ridai Govinda Pombo)
  If URLEncode Then strValue = Server.URLEncode(strValue)
  Equal = Ampersand & strOption & "=" & strValue
  Rem ---
  
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub AddEditForm(Optional ByVal strEditable As String = "", _
                       Optional ByVal strAction As String = "", _
                       Optional ByVal strWhat As String = "", _
                       Optional ByVal strFilter As String = "", _
                       Optional ByVal strDefaults As String = "", _
                       Optional ByVal strList As String = "")
  
  Const ur_efQueryStrEditable = "EE"
  Const ur_efQueryStrAction = "EA"
  Const ur_efQueryStrWhat = "EW"
  Const ur_efQueryStrFilter = "ER"
  Const ur_efQueryStrDefaults = "ED"
  Const ur_efQueryStrMove = "EM"
  Const ur_efQueryStrList = "EL"
  
  If strEditable > "" Then Add ur_efQueryStrEditable, strEditable
    
  If strAction > "" Then Add ur_efQueryStrAction, strAction
  
  If strWhat > "" Then Add ur_efQueryStrWhat, strWhat
  
  If strFilter > "" Then Add ur_efQueryStrFilter, strFilter
  
  If strDefaults > "" Then Add ur_efQueryStrDefaults, strDefaults
  
  If strList > "" Then Add ur_efQueryStrList, strList
  
End Sub
