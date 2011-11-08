VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "UploadItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"UploadPackItem"
Attribute VB_Ext_KEY = "Member0" ,"UploadPackItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' =========================================================================
'  /HiperTools30.UploadItems
' -------------------------------------------------------------------------
'  Descricao   : Lista de itens do Upload
'  Criaçao     : 11/01/2000 12:18
'  Local       : Brasilia/DF
'  Elaborado   : Ridai Govinda Pombo <ridai@zevallos.com.br>
'                Kelson da Costa Medeiros <kelson@zevallos.com.br>
'  Versao      : 2.0.0
'  Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------
Option Explicit

Const ERR_DONT_HAVE_UPLOAD_FILE As Long = vbObjectError + 105

Private clctPack As Collection, _
        sintPosition As Long

Private sobjFirstFile As HiperTools30.UploadItem

Private ScriptContext As ASPTypeLibrary.ScriptingContext, _
        Response As ASPTypeLibrary.Response, _
        Server As ASPTypeLibrary.Server

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  TimeBombX xtbUpload, 15656, 12934, 3080, 14624
  
  Set ScriptContext = ScriptContext
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  
End Sub

Public Sub OnEndPage()
  
  Set ScriptContext = Nothing
  Set Response = Nothing
  Set Server = Nothing

End Sub

Private Sub Class_Initialize()
  
  Set clctPack = New Collection
    
  sintPosition = 1
  
End Sub

Private Sub Class_Terminate()
  Clear
  Set clctPack = Nothing
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Add(ByVal strName As String, Optional ByRef varValue As Variant = "") As UploadItem

  Dim objNewMember As UploadItem
  Set objNewMember = Server.CreateObject("HiperTools30.UploadItem")

  If IsObject(varValue) Then
    Set objNewMember.Value = varValue
  Else
    objNewMember.Value = varValue
  End If
  
  clctPack.Add objNewMember, strName
  
  sintPosition = sintPosition + 1
  objNewMember.Name = strName
  objNewMember.Position = sintPosition
  
  Set Add = objNewMember
  Set objNewMember = Nothing

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Item(ByVal varIndexName As Variant) As UploadItem
Attribute Item.VB_UserMemId = 0
  Set Item = clctPack(varIndexName)
End Property

Public Property Get FirstFile() As HiperTools30.UploadItem
  If Not sobjFirstFile Is Nothing Then
    Set FirstFile = sobjFirstFile
  Else
    Err.Raise ERR_DONT_HAVE_UPLOAD_FILE, , "Não existe um arquivo para Upload."
  End If
End Property

Friend Property Set FirstFile(objParam As HiperTools30.UploadItem)
  Set sobjFirstFile = objParam
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Count() As Long
    Count = clctPack.Count
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub Remove(ByVal varIndexName As Variant)
    clctPack.Remove varIndexName
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = clctPack.[_NewEnum]
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function Find(ByVal varIndexName As Variant, _
                     Optional ByVal blnRemove As Boolean = False) As Boolean
  Dim varAux As UploadItem
  
  Find = False
  
  On Error Resume Next
  Set varAux = clctPack(varIndexName)
  
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
  Else
    Find = True
    
    If blnRemove Then
      Remove varIndexName
    End If
    
    On Error GoTo 0
  End If
  
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub Clear()
  Do While clctPack.Count > 0
    Remove 1
  Loop
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Value(ByVal varIndexName As Variant) As Variant
  Value = clctPack(varIndexName).Value
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Value(ByVal varIndexName As Variant, varValue As Variant)
  clctPack(varIndexName).Value = varValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Name(ByVal varIndexName As Variant) As String
  Name = clctPack(varIndexName).Name
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Name(ByVal varIndexName As Variant, strNewName As String)
  Dim varLastValue As Variant
  
  varLastValue = clctPack(varIndexName).Value
  
  Remove varIndexName
  Add strNewName, varLastValue
  
End Property
