VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Upgrades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Coleção de Updates realizados no HiperTools"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Update"
Attribute VB_Ext_KEY = "Member0" ,"Update"
Option Explicit

Private sobjUpgrades As Collection

Public Function Add(ByVal lngCode As Long, ByVal dtmStart As Date, ByVal dtmEnd As Date, _
                    ByVal objStartVer As HiperTools30.Version, ByVal objEndVer As HiperTools30.Version) As Upgrade
  Dim objNewMember As HiperTools30.Upgrade
  Set objNewMember = New HiperTools30.Upgrade

  objNewMember.Key = "Upgrade" & lngCode
  objNewMember.Code = lngCode
  objNewMember.StartingDate = dtmStart
  objNewMember.EndingDate = dtmEnd
  Set objNewMember.StartingVersion = objStartVer
  Set objNewMember.EndingVersion = EndingVersion
  sobjUpgrades.Add objNewMember, objNewMember.Key

  Set Add = objNewMember
  Set objNewMember = Nothing

End Function

Public Property Get Item(ByVal vntIndexKey As Variant) As HiperTools30.Upgrade
Attribute Item.VB_UserMemId = 0
  Set Item = sobjUpgrades(vntIndexKey)
End Property

Public Property Get Count() As Long
  Count = sobjUpgrades.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
  sobjUpgrades.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = sobjUpgrades.[_NewEnum]
End Property

Private Sub Class_Initialize()
  Set sobjUpgrades = New Collection
End Sub

Private Sub Class_Terminate()
  Set sobjUpgrades = Nothing
End Sub
