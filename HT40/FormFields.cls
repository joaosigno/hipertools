VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "FormFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"FormField"
Attribute VB_Ext_KEY = "Member0" ,"FormField"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' =========================================================================
' FormFields.cls
' -------------------------------------------------------------------------
' Descricao   : Classe de Collection de Fields p/ o Form.ShowField... etc
' Criaçao     : 10/02/2000 12:47PM
' Local       : Brasilia/DF
' Elaborado   : Ridai Govinda Pombo <ridai@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 97-2000 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------

  Option Explicit
  
  'Propriedades complexas:
  Private sstrLastValidation As String
  
  Private scolFields As Collection
  
  Private ScriptContext As ScriptingContext

  Private Request As Request, _
          Response As Response, _
          Server As Server, _
          Session As Session, _
          Application As Application

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  Set ScriptContext = ScriptContext
  
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application
  
  'Propriedades complexas:
  sstrLastValidation = ""

End Sub

Public Sub OnEndPage()
  
  Set ScriptContext = Nothing
  
  Set Request = Nothing
  Set Response = Nothing
  Set Server = Nothing
  Set Session = Nothing
  Set Application = Nothing

End Sub

'Procedimentos típicos de Collection:

Private Sub Class_Initialize()
    Set scolFields = New Collection
End Sub

Private Sub Class_Terminate()
    Set scolFields = Nothing
End Sub

Public Function Add(ByVal strName As String) As HiperTools30.FormField
    
  Dim objFormField As HiperTools30.FormField
  Set objFormField = Server.CreateObject("HiperTools30.FormField")

  scolFields.Add objFormField, strName

  Set Add = objFormField
  Set objFormField = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As HiperTools30.FormField
Attribute Item.VB_UserMemId = 0
  Set Item = scolFields(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = scolFields.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    scolFields.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = scolFields.[_NewEnum]
End Property

'Propriedades
Public Property Get LastValidation() As String
  LastValidation = sstrLastValidation
End Property

Public Property Let LastValidation(ByVal strParam As String)
  sstrLastValidation = strParam
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Find(ByVal varIndexKey As Variant, _
                     Optional ByVal blnRemove As Boolean = False) As Boolean
  Dim varAux As HiperTools30.FormField
  
  Find = False
  
  On Error Resume Next
  Set varAux = scolFields(varIndexKey)
  
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
  Else
    Find = True
    
    If blnRemove Then
      Remove varIndexKey
    End If
    
    On Error GoTo 0
  End If
  
End Function


