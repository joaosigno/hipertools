VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PackItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
' =========================================================================
'  /HiperTools30.PackItem
' -------------------------------------------------------------------------
'  Descricao   : Pack, clone do "Scripting.Dictionary"
'  Criaçao     : 11/01/2000 12:18
'  Local       : Brasilia/DF
'  Elaborado   : Ridai Govinda Pombo <ridai@zevallos.com.br>
'                Eduardo Alves Cortes <edualves@zevallos.com.br>
'  Versao      : 2.0.0
'  Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------
Option Explicit

Private sstrKey As String
Private sintPosition As Long
Private sobjPack As HiperTools30.Pack

'
Public Value As Variant
Attribute Value.VB_VarUserMemId = 0

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Key() As String
  Key = sstrKey
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Friend Property Let Key(ByVal strKey As String)
  sstrKey = strKey
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Position() As Long
  Position = sintPosition
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Friend Property Let Position(ByVal intPosition As Long)
  sintPosition = intPosition
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Pack() As HiperTools30.Pack
  If sobjPack Is Nothing Then
    Set sobjPack = New HiperTools30.Pack
  End If
  
  Set Pack = sobjPack
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Set Pack(objPack As Object)
  If TypeName(objPack) = "HiperTools30.Pack" Or objPack Is Nothing Then
    If Not sobjPack Is Nothing Then
      sobjPack.Clear
      Set sobjPack = Nothing
    End If
    
    Set sobjPack = objPack
  End If
End Property

Private Sub Class_Initialize()
  sintPosition = 1
End Sub

Private Sub Class_Terminate()
  If Not sobjPack Is Nothing Then
    Set sobjPack = Nothing
  End If
End Sub
