VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "RegistryValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'local variable(s) to hold property value(s)
Private mvarName As String 'local copy
Private mvarType As Long 'local copy
Private objParent As Registry

Public Property Let DataType(ByVal lngData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Type = 5
    mvarType = lngData
End Property

Public Property Get DataType() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Type
    DataType = mvarType
End Property

Public Property Let Name(ByVal strData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Name = 5
    mvarName = strData
End Property

Public Property Get Name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Name
    Name = mvarName
End Property


Friend Property Set Parent(ByRef objRegistry As Registry)
    Set objParent = objRegistry
End Property

Friend Property Get Parent() As Registry
    Parent = objParent
End Property

Private Sub Class_Initialize()
  TimeBombX xtbRegistry, 15656, 12934, 3080, 14624
End Sub

Private Sub Class_Terminate()
    Set objParent = Nothing
End Sub
