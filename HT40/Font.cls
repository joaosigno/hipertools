VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Font"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Face As String
Public Color As String
Public Size As Integer
Public Bold As Boolean
Public Italic As Boolean
Public Underline As Boolean

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  Size = 2
  Bold = False
  Italic = False
  Underline = False

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function BeginFont()
  Dim strAux As String
  
  If Size > 0 Or Face > "" Or Color > "" Then
    strAux = "<FONT"
    
    If Face > "" Then strAux = strAux & " FACE=""" & Face & """"
    If Color > "" Then strAux = strAux & " COLOR=""" & Color & """"
    If Size > 0 Then strAux = strAux & " SIZE=" & Size
    
    strAux = strAux & ">"
  End If
  
  If Bold Then strAux = "<B>" & strAux
  If Italic Then strAux = "<I>" & strAux
  If Underline Then strAux = "<U>" & strAux
  
  BeginFont = strAux

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function EndFont()
  Dim strAux As String
  
  If Size > 0 Or Face > "" Or Color > "" Then strAux = "</FONT>"
  If Bold Then strAux = strAux & "</B>"
  If Italic Then strAux = strAux & "</I>"
  If Underline Then strAux = strAux & "</U>"
  
  EndFont = strAux

End Function

