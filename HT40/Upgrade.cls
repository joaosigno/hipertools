VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Upgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

Private slngCode As Long
Private sdtmStarting As Date
Private sdtmEnding As Date
Private sobjStartingVersion As Variant
Private sobjEndingVersion As Variant

Public Property Set EndingVersion(objVer As Version)
  Set sobjEndingVersion = objVer
End Property

Public Property Get EndingVersion() As Version
  Set EndingVersion = sobjEndingVersion
End Property

Public Property Set StartingVersion(objVer As Version)
  Set sobjStartingVersion = objVer
End Property

Public Property Get StartingVersion() As Variant
  Set StartingVersion = sobjStartingVersion
End Property

Public Property Let EndingDate(ByVal dtmEnd As Date)
  sdtmEnding = dtmEnd
End Property

Public Property Get EndingDate() As Date
  EndDate = sdtmEnding
End Property

Public Property Let StartingDate(ByVal dtmStart As Date)
  sdtmStarting = dtmStart
End Property

Public Property Get StartingDate() As Date
  StartingDate = sdtmStarting
End Property

Public Property Let Code(ByVal lngCode As Long)
  slngCode = lngCode
End Property

Public Property Get Code() As Long
  Code = slngCode
End Property
