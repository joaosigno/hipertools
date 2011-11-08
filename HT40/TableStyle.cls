VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "TableStyle"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Bordas externas
Public ExternalBorder As HiperTools30.Border

' Bordas do cabecalho
Public HeaderBorder As HiperTools30.Border

' Bordas internas
Public InternalBorder As HiperTools30.Border

' Font do título
Public TitleFont As HiperTools30.Font

' Font do cabecalho
Public HeaderFont As HiperTools30.Font

' Font interna
Public InternalFont As HiperTools30.Font

' Cor do título
Public TitleColor As String

' Formato da cor
  ' tbStFormatHeader                 - Cabecalho cor 1, resto cor 2
  ' tbStFormatHeaderAlternateColumns - Alterna colunas
  ' tbStFormatAlternateColumns       - Cabecalho fixo, alterna colunas
  ' tbStFormatHearderAlternateLines  - Alterna linhas
  ' tbStFormatAlternateLines         - Cabecalho fixo, alterna linhas
  ' tbStFormatNothing                - Todo da cor 1(menos o Título)
  ' tbStFormatTitle                  - Todo da cor 1(inclusive Título)
Public ColorFormat As Integer

' Criado (Ridai Govinda)
  ' tbStFormatAlternateColumns       - Cabecalho fixo, alterna colunas
  ' tbStFormatAlternateLines         - Cabecalho fixo, alterna linhas
  ' tbStFormatNothing                - Sem configuração
Public BackgroundFormat As Integer
'----------------------

'Alterado/Criado (Ridai Govinda)
'Removida as variáveis públicas de cor1, 2 e 3

Private sstrBaseColor As String
Private sstrHeaderColor As String
Private sstrFirstAltColor As String
Private sstrAlternateColor As String
Private sstrLastColor As String

Private sstrColor1 As String
Private sstrColor2 As String
Private sstrColor3 As String

'Background passado p/ objeto Style
Public CellBackground As String
Public RowBackground As String
Public TableBackground As String
Public TitleBackground As String

  '[Cabeçalho:
Public HeaderCellBackground As String 'Coloca background só nas colunas do cabeçalho
Public HeaderRowBackground As String 'Coloca background na linha de cabeçalho
  '[Alternações]
Public FirstAltBackground As String
Public SecondAltBackground As String
'------

' Estilo de bordas:
  ' tbBdFormatCompleteGrid    - Grade completa
  ' tbBdFormatOnlyLines       - Apenas linhas
  ' tbBdFormatOutterGridLines - Grade externa + linhas
  ' tbBdFormatInvisible       - Invisível
Public BorderFormat As Integer

Private Server As Server

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  TimeBombX xtbTableStyle, 15656, 12934, 3080, 14624
  
  'Criado (Ridai Govinda)
  CellBackground = ""
  RowBackground = ""
  TableBackground = ""
  TitleBackground = ""
  HeadColsBackground = ""
  HeadRowBackground = ""
  FirstAltBackground = ""
  SecondAltBackground = ""
  BackgroundFormat = tbStFormatNothing
  '---------------------
  
  Set Server = ScriptContext.Server
  Set ExternalBorder = Server.CreateObject("HiperTools30.Border")
  Set HeaderBorder = Server.CreateObject("HiperTools30.Border")
  Set InternalBorder = Server.CreateObject("HiperTools30.Border")
  Set TitleFont = Server.CreateObject("HiperTools30.Font")
  Set HeaderFont = Server.CreateObject("HiperTools30.Font")
  Set InternalFont = Server.CreateObject("HiperTools30.Font")

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get NeedTwoTables() As Boolean
' (Retorna Se for necessário colocar uma tabela dentro da outra na hora de montar o table, por
' motivo de construção de bordas)
  NeedTwoTables = (BorderFormat = 2)
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let BorderWidth(bytWidth As Byte)
  ExternalBorder.Width = bytWidth
  HeaderBorder.Width = bytWidth
  InternalBorder.Width = bytWidth
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let BorderColor(strColor As String)
  ExternalBorder.Color = strColor
  HeaderBorder.Color = strColor
  InternalBorder.Color = strColor
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let FontSize(intSize As Integer)
  HeaderFont.Size = intSize
  InternalFont.Size = intSize
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let FontColor(strColor As String)
  HeaderFont.Color = strColor
  InternalFont.Color = strColor
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let FontBold(blnBold As Boolean)
  HeaderFont.Bold = blnBold
  InternalFont.Bold = blnBold
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let FontItalic(blnItalic As Boolean)
  HeaderFont.Italic = blnItalic
  InternalFont.Italic = blnItalic
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let FontUnderline(blnUnderline As Boolean)
  HeaderFont.Underline = blnUnderline
  InternalFont.Underline = blnUnderline
End Property
 
 'Criado (Ridai Govinda)
Rem =========================================================================
Rem Cor base, cor de cabeçalho e primeira cor de alternação:
Rem -------------------------------------------------------------------------
Public Property Get BaseColor() As String
  BaseColor = sstrBaseColor
End Property

Public Property Let BaseColor(ByRef strValue As String)
 sstrBaseColor = strValue
End Property

Public Property Get HeaderColor() As String
  HeaderColor = sstrHeaderColor
End Property

Public Property Let HeaderColor(ByRef strValue As String)
 sstrHeaderColor = strValue
End Property

Public Property Get FirstAltColor() As String
  FirstAltColor = sstrFirstAltColor
End Property

Public Property Let FirstAltColor(ByRef strValue As String)
 sstrFirstAltColor = strValue
End Property

Rem =========================================================================
Rem Segunda cor de alternação
Rem -------------------------------------------------------------------------
Public Property Get AlternateColor() As String
  AlternateColor = sstrAlternateColor
End Property

Public Property Let AlternateColor(ByRef strValue As String)
 sstrAlternateColor = strValue
End Property

Rem =========================================================================
Rem última cor (de alternação)
Rem -------------------------------------------------------------------------
Public Property Get LastColor() As String
  LastColor = sstrLastColor
End Property

Public Property Let LastColor(ByRef strValue As String)
 sstrLastColor = strValue
End Property
'--------

'Criado (Ridai Govinda)
' Cores do ColorFormat
Rem Cor base, cor de cabeçalho e primeira cor de alternação:
Public Property Get Color1() As String
  Color1 = sstrColor1
End Property

Public Property Let Color1(ByVal strColor As String)
  sstrColor1 = strColor
  BaseColor = strColor
  HeaderColor = strColor
  FirstAltColor = strColor
End Property

Public Property Get Color2() As String
  Color2 = sstrColor2
End Property

Public Property Let Color2(ByVal strColor As String)
  sstrColor2 = strColor
  AlternateColor = strColor
End Property

Public Property Get Color3() As String
  Color3 = sstrColor3
End Property

Public Property Let Color3(ByVal strColor As String)
  sstrColor3 = strColor
  LastColor = strColor
End Property
'-----------------------

'---------------------------------------------------------------------
' criado Flavio Costa
'---------------------------------------------------------------------
Public Property Get ColumnBackGround() As String
  ColumnBackGround = CellBackground
End Property

Public Property Let ColumnBackGround(ByVal strBackground As String)
  CellBackground = strBackground
End Property

Public Property Get HeadColsBackground() As String
  HeadColsBackground = HeaderCellBackground
End Property

Public Property Let HeadColsBackground(ByVal strBackground As String)
  HeaderCellBackground = strBackground
End Property

Public Property Get HeadRowBackground() As String
  HeadRowBackground = HeaderRowBackground
End Property

Public Property Let HeadRowBackground(ByVal strBackground As String)
  HeaderRowBackground = strBackground
End Property

