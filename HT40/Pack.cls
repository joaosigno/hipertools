VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"PackItem"
Attribute VB_Ext_KEY = "Member0" ,"PackItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' =========================================================================
'  /HiperTools30.Pack
' -------------------------------------------------------------------------
'  Descricao   : Pack, clone do "Scripting.Dictionary"
'  Criaçao     : 11/01/2000 12:18
'  Local       : Brasilia/DF
'  Elaborado   : Ridai Govinda Pombo <ridai@zevallos.com.br>
'                Eduardo Alves Cortes <edualves@zevallos.com.br>
'  Versao      : 2.0.0
'  Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------
Private clctPack As Collection
Private sintPosition As Long

Private Response As ASPTypeLibrary.Response, _
        Server As ASPTypeLibrary.Server

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  'TimeBombX xtbPack, 15656, 12934, 3080, 14624
  
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
  sstrComments = App.Comments
  sstrClassTitle = App.Title
  sstrVersion = App.Major & "." & _
                App.Minor & "." & _
                App.Revision
  sstrCompanyName = App.CompanyName
  sstrProductName = App.ProductName

    Set clctPack = New Collection
    
    sintPosition = 1
End Sub

Private Sub Class_Terminate()
    Clear
    Set clctPack = Nothing
End Sub

Public Property Get Comments() As String
  Comments = sstrComments
End Property

Public Property Get ClassTitle() As String
  ClassTitle = sstrClassTitle
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
Public Function Add(ByVal strKey As String, Optional varValue As Variant = "") As PackItem
  Dim objNewMember As PackItem
  Set objNewMember = New PackItem

  If IsObject(varValue) Then
    Set objNewMember.Value = varValue
  Else
    objNewMember.Value = varValue
  End If
  
  clctPack.Add objNewMember, strKey
  
  sintPosition = sintPosition + 1
  objNewMember.Key = strKey
  objNewMember.Position = sintPosition
  
  Set Add = objNewMember
  Set objNewMember = Nothing

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Item(ByVal varIndexKey As Variant) As PackItem
Attribute Item.VB_UserMemId = 0
  Set Item = clctPack(varIndexKey)
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
Public Sub Remove(ByVal varIndexKey As Variant)
    Set clctPack(varIndexKey).Pack = Nothing
    clctPack.Remove varIndexKey
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
Public Function Find(ByVal varIndexKey As Variant, _
                     Optional ByVal blnRemove As Boolean = False) As Boolean
  Dim varAux As PackItem
  
  Find = False
  
  On Error Resume Next
  Set varAux = clctPack(varIndexKey)
  
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
Public Property Get Value(ByVal varIndexKey As Variant) As Variant
  Value = clctPack(varIndexKey).Value
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Value(ByVal varIndexKey As Variant, varValue As Variant)
  clctPack(varIndexKey).Value = varValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Key(ByVal varIndexKey As Variant) As String
  Key = clctPack(varIndexKey).Key
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Key(ByVal varIndexKey As Variant, strNewKey As String)
  Dim varLastValue As Variant
  
  varLastValue = clctPack(varIndexKey).Value
  
  Remove varIndexKey
  Add strNewKey, varLastValue
  
End Property
