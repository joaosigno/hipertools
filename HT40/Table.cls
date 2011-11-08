VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =========================================================================
' Table.cls
' -------------------------------------------------------------------------
' Descricao : Objeto da Gerencia as Tabelas
' Criacao   : 1/30/99 3:34AM
' Local     : Brasilia/DF
' Elaborado : Ruben Zevallos Jr. - zevallos@zevallos.com.br
' Versao    : 1.0.0
' Copyright : 1999 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------

Option Explicit

Private Const conPixelImage = "/img/pixel.gif"

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
        Application As Application, _
        ScriptContext As ScriptingContext

'Criado (Ridai Govinda)
Private sobjTableStyle As HiperTools30.TableStyle
Private sblnTableOpened As Boolean
'Private sclctGColors As Collection
'----------------------

'Criado (Flavio Costa)
Public CellWidth As String
Public CellHeight As String
Public CellAlign As String
Public CellRowSpan As String
Public CellColSpan As String
Public CellVAlign As String
Public CellColor As String
Public CellNoWrap As Boolean
Public Padding As String
Public ShowBorder As Boolean
'----------------------

'Criados (Ridai Govinda)
Public CellGraphTotalWidth As Integer
Public CellGraphMeterWidth As Integer
Public CellGraphMeterRightDivider As String
Public CellGraphMeterLeftDivider As String
Public CellGraphMeterCenterDivider As String
Public CellGraphHeight As Integer
Public CellGraphImage As String
'---------------------

Private clctTables As New Collection, _
        sintGraphNum As Integer, sintGraphActual As Integer, _
        sarrMax(5) As Long, sarrTotal(5) As Double, sarrMin(5) As Double, _
        sblnNoCellColors As Boolean, _
        sblnNoBorderColors As Boolean 'Criado (Ridai Govinda)

'Alterado (Ridai)
'Removido e colocado como Property mesmo
'Public ColumnBackGround As String
'
'Public RowBackGround As String
'
'Public TableBackGround As String
'
'Public Style As HiperTools30.TableStyle

Public Sub OnStartPage(ScriptContext As ScriptingContext)

  TimeBombX xtbTable, 15656, 12934, 3080, 14624

  Set ScriptContext = ScriptContext
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application
  Set sobjTableStyle = Server.CreateObject("HiperTools30.TableStyle")

  sintGraphNum = 0
  Padding = 3

' (Formatação padrão do table)
  Style.TitleFont.Size = 3
  Style.TitleFont.Bold = True
  Style.TitleColor = "white"
  Style.Color1 = "#CCCCFF"
  'Alterado (Ridai Govinda)
  Style.AlternateColor = "#ECECFF"
  Rem Style.Color2 = "#ECECFF"
  Style.LastColor = "#FFFFFF"
  Rem Style.Color3 = "#FFFFFF"
  '-----------------------
  Style.ColorFormat = tbStFormatHeader
  Style.BorderFormat = tbBdFormatCompleteGrid
  Style.BorderColor = ""

  ShowBorder = False
  'Criado (Ridai Govinda)
  sblnNoCellColors = False
  sblnNoBorderColors = False
  CellGraphTotalWidth = 80
  CellGraphMeterWidth = 120
  CellGraphMeterDivider = "|"
  CellGraphHeight = 15
  CellGraphImage = Session("htDefaultHiperToolsWorkPath") & conPixelImage
  '---------------------

End Sub

Public Sub OnEndPage()

  'Criado (Ridai Govinda)
  If sblnTableOpened Then EndTable
  '----------------------

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

Public Property Get Comments() As String
  Comments = sstrComments
End Property

Public Property Get Title() As String
  Title = sstrTitle
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
Public Property Get BorderColor() As String
  BorderColor = Style.ExternalBorder.Color
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let BorderColor(ByRef strValue As String)
  Style.BorderColor = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get HeadRowColor() As String
  'Alterado (Ridai Govinda)
  HeadRowColor = Style.HeaderColor
  Rem HeadRowColor = Style.Color1
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let HeadRowColor(ByRef strValue As String)
  'Alterado (Ridai Govinda)
  Style.HeaderColor = strValue
  Rem Style.Color1 = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get RowColor() As String

  Select Case Style.ColorFormat
    Case tbStFormatHeader, tbStFormatAlternateColumns
      'Alterado (Ridai Govinda)
      RowColor = Style.AlternateColor
      Rem RowColor = Style.Color2
      '-----------------------

    Case tbStFormatHeaderAlternateColumns
      'Alterado (Ridai Govinda)
      RowColor = Style.FirstAltColor
      Rem RowColor = Style.Color1
      '----

    Case tbStFormatHearderAlternateLines
      If (RowNumber Mod 2 = 1) Then
        'Alterado (Ridai Govinda)
        RowColor = Style.FirstAltColor
        Rem RowColor = Style.Color1
        '----

      Else
        'Alterado (Ridai Govinda)
        RowColor = Style.AlternateColor
        Rem RowColor = Style.Color2
        '-----------------------

      End If

    Case tbStFormatAlternateLines
      If (RowNumber Mod 2 = 0) Then
        'Alterado (Ridai Govinda)
        RowColor = Style.AlternateColor
        Rem RowColor = Style.Color2
        '-----------------------

      Else
        'Alterado (Ridai Govinda)
        RowColor = Style.LastColor
        Rem RowColor = Style.Color3
        '-----------------------

      End If

    Case tbStFormatNothing, tbStFormatTitle
      'Alterado (Ridai Govinda)
      RowColor = Style.BaseColor
      Rem RowColor = Style.Color1
      '----

  End Select

End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let RowColor(ByRef strValue As String)

  Select Case Style.ColorFormat
    Case tbStFormatHeader, tbStFormatAlternateColumns
      'Alterado (Ridai Govinda)
      Style.AlternateColor = strValue
      Rem Style.Color2 = strValue
      '-----------------------

    Case tbStFormatHeaderAlternateColumns
      'Alterado (Ridai Govinda)
      Style.FirstAltColor = strValue
      Rem Style.Color1 = strValue
      '----

    Case tbStFormatHearderAlternateLines
      If (RowNumber Mod 2 = 1) Then
        'Alterado (Ridai Govinda)
        Style.FirstAltColor = strValue
        Rem Style.Color1 = strValue
        '----

      Else
        'Alterado (Ridai Govinda)
        Style.AlternateColor = strValue
        Rem Style.Color2 = strValue
        '-----------------------

      End If

    Case tbStFormatAlternateLines
      If (RowNumber Mod 2 = 0) Then
        'Alterado (Ridai Govinda)
        Style.AlternateColor = strValue
        Rem Style.Color2 = strValue
        '-----------------------

      Else
        'Alterado (Ridai Govinda)
        Style.LastColor = strValue
        Rem Style.Color3 = strValue
        '-----------------------

      End If

    Case tbStFormatNothing, tbStFormatTitle
      'Alterado (Ridai Govinda)
      Style.BaseColor = strValue
      Rem Style.Color1 = strValue
      '----

  End Select

End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get RowFontSize() As String
  RowFontSize = Style.InternalFont.Size
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let RowFontSize(ByRef strValue As String)
  Style.FontSize = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Border() As String
  Border = Spacing
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Border(ByRef strValue As String)
  Style.BorderWidth = strValue
End Property


Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ForPrintBorder() As Boolean
  ForPrintBorder = ShowBorder
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let ForPrintBorder(ByRef blnValue As Boolean)
  ShowBorder = blnValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ColumnWidth() As String
  ColumnWidth = CellWidth
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let ColumnWidth(ByRef strValue As String)
  CellWidth = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ColumnHeight() As String
  ColumnHeight = CellHeight
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let ColumnHeight(ByRef strValue As String)
  CellHeight = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ColumnAlign() As String
  ColumnAlign = CellAlign
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let ColumnAlign(ByRef strValue As String)
  CellAlign = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ColumnColSpan() As String
  ColumnColSpan = CellColSpan
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let ColumnColSpan(ByRef strValue As String)
  CellColSpan = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ColumnRowSpan() As String
  ColumnRowSpan = CellRowSpan
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let ColumnRowSpan(ByRef strValue As String)
  CellRowSpan = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ColumnVAlign() As String
  ColumnVAlign = CellVAlign
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let ColumnVAlign(ByRef strValue As String)
  CellVAlign = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ColumnColor() As String
  ColumnColor = CellColor
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let ColumnColor(ByRef strValue As String)
  CellColor = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ColumnNoWrap() As Boolean
  ColumnNoWrap = CellNoWrap
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let ColumnNoWrap(ByRef strValue As Boolean)
  CellNoWrap = strValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Spacing() As String
  Spacing = Style.ExternalBorder.Width
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Spacing(ByRef strValue As String)
  Style.ExternalBorder.Width = strValue
End Property

Public Property Get GraphTotal() As Double
  GraphTotal = sarrTotal(sintGraphActual)
End Property

Public Property Let GraphTotal(ByVal dblValue As Double)
  sarrTotal(sintGraphActual) = dblValue
End Property

Public Property Get GraphMax() As Double
  GraphMax = sarrMax(sintGraphActual)
End Property

Public Property Let GraphMax(ByVal dblValue As Double)
  If dblValue <= sarrMin(sintGraphActual) Then
    ShowError 125, "Table", "GraphMin", "Máximo definido é menor que o mínimo definido."
  Else
    sarrMax(sintGraphActual) = dblValue
  End If
End Property

Public Property Get GraphMin() As Double
  GraphMin = sarrMin(sintGraphActual)
End Property

Public Property Let GraphMin(ByVal dblValue As Double)
  If dblValue >= sarrMax(sintGraphActual) Then
    ShowError 125, "Table", "GraphMin", "Mínimo maior que o máximo"
  Else
    sarrMin(sintGraphActual) = dblValue
  End If
End Property

Private Property Get ColumnNumber() As Integer
  ColumnNumber = clctTables(clctTables.Count).ColumnNumber
End Property

Private Property Let ColumnNumber(ByVal Value As Integer)

  Dim Stack As HiperTools30.TableStack

  Set Stack = clctTables(clctTables.Count)
  Stack.ColumnNumber = Value
  clctTables.Remove clctTables.Count
  clctTables.Add Stack

End Property

Private Property Get RowNumber() As Integer
  RowNumber = clctTables(clctTables.Count).RowNumber
End Property

Private Property Let RowNumber(ByVal Value As Integer)

  Dim Stack As HiperTools30.TableStack

  Set Stack = clctTables(clctTables.Count)
  Stack.RowNumber = Value
  clctTables.Remove clctTables.Count
  clctTables.Add Stack

End Property

Private Property Get ColumnOpened() As Boolean
  ColumnOpened = clctTables(clctTables.Count).ColumnOpened
End Property

Private Property Let ColumnOpened(ByVal Value As Boolean)

  Dim Stack As HiperTools30.TableStack

  Set Stack = clctTables(clctTables.Count)
  Stack.ColumnOpened = Value
  clctTables.Remove clctTables.Count
  clctTables.Add Stack

End Property

Private Property Get RowOpened() As Boolean
  RowOpened = clctTables(clctTables.Count).RowOpened
End Property

Private Property Let RowOpened(ByVal Value As Boolean)

  Dim Stack As HiperTools30.TableStack

  Set Stack = clctTables(clctTables.Count)
  Stack.RowOpened = Value
  clctTables.Remove clctTables.Count
  clctTables.Add Stack

End Property

Private Property Get HaveSecondTable() As Boolean
  HaveSecondTable = clctTables(clctTables.Count).HaveSecondTable
End Property

Private Property Let HaveSecondTable(ByVal Value As Boolean)

  Dim Stack As HiperTools30.TableStack

  Set Stack = clctTables(clctTables.Count)
  Stack.HaveSecondTable = Value
  clctTables.Remove clctTables.Count
  clctTables.Add Stack

End Property

Private Property Get TableColumns() As Integer
  TableColumns = clctTables(clctTables.Count).TableColumns
End Property

Private Property Let TableColumns(ByVal Value As Integer)

  Dim Stack As HiperTools30.TableStack

  Set Stack = clctTables(clctTables.Count)
  Stack.TableColumns = Value
  clctTables.Remove clctTables.Count
  clctTables.Add Stack

End Property

Private Property Get IsHeaderRow() As Boolean
  IsHeaderRow = clctTables(clctTables.Count).HeaderRow
End Property

Private Property Let IsHeaderRow(ByVal Value As Boolean)

  Dim Stack As HiperTools30.TableStack

  Set Stack = clctTables(clctTables.Count)
  Stack.HeaderRow = Value
  clctTables.Remove clctTables.Count
  clctTables.Add Stack

End Property

Private Property Get LastRow() As Integer
  LastRow = clctTables(clctTables.Count).LastRow
End Property

Private Property Let LastRow(ByVal Value As Integer)

  Dim Stack As HiperTools30.TableStack

  Set Stack = clctTables(clctTables.Count)
  Stack.LastRow = Value
  clctTables.Remove clctTables.Count
  clctTables.Add Stack

End Property

Private Property Get HaveBorder() As Boolean
  HaveBorder = clctTables(clctTables.Count).HaveBorder
End Property

Private Property Let HaveBorder(ByVal Value As Boolean)

  Dim Stack As HiperTools30.TableStack

  Set Stack = clctTables(clctTables.Count)
  Stack.HaveBorder = Value
  clctTables.Remove clctTables.Count
  clctTables.Add Stack

End Property

Private Property Get HaveTitle() As Boolean
  HaveTitle = clctTables(clctTables.Count).HaveTitle
End Property

Private Property Let HaveTitle(ByVal Value As Boolean)

  Dim Stack As HiperTools30.TableStack

  Set Stack = clctTables(clctTables.Count)
  Stack.HaveTitle = Value
  clctTables.Remove clctTables.Count
  clctTables.Add Stack

End Property

Private Sub CloseCell()

  If ColumnOpened Then
    EndCell
  End If

End Sub

Private Sub CloseLine()

  If RowOpened Then
    EndRow
  End If

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub Clear()

  ColumnBackGround = ""
  ColumnWidth = ""
  ColumnHeight = ""
  ColumnAlign = ""
  ColumnColSpan = ""
  ColumnRowSpan = ""
  ColumnVAlign = ""
  ColumnColor = ""
  ColumnNoWrap = ""
  Spacing = ""
  Padding = ""

End Sub

Rem =========================================================================
Rem Inicio da tabela
Rem -------------------------------------------------------------------------
Public Sub BeginTable(Optional ByVal strTableWidth As String = "", _
                      Optional ByVal strMessage As String = "", _
                      Optional ByVal intColumns As Integer = 1, _
                      Optional ByVal blnWithBorder As Boolean = False, _
                      Optional ByVal strTableHeight As String = "")

  Dim strBorderColor As String, strBorderWidth As String, strSpacing As String
  Dim strHeight As String, strLineRow As String, strWidth As String

  If strTableWidth > "" Then
     strWidth = " WIDTH=""" & strTableWidth & """"
  End If

  If strTableHeight > "" Then
     strHeight = " HEIGHT=""" & strTableHeight & """"
  End If

  clctTables.Add Server.CreateObject("HiperTools30.TableStack")

  'Alterado/Criado (Ridai Govinda Pombo)
  sblnTableOpened = True
  '------------------------------------

  TableColumns = intColumns
  HaveBorder = blnWithBorder
  HaveTitle = (strMessage > "")

  strSpacing = " CELLSPACING=" & Spacing

  If blnWithBorder Then

    If Style.NeedTwoTables Then

      HaveSecondTable = True
      strSpacing = " CELLSPACING=0"
      'Alterado (Ridai Govinda)
      Response.Write "<TABLE BORDER=""0"""
      If Not NoBorderColors Then
        Response.Write " BGCOLOR=""" & Style.ExternalBorder.Color & """"
      End If
      Response.Write " CELLSPACING = " & (Spacing - 1)
      '-----------------------

      If strWidth > "" Then
        Response.Write " WIDTH=""" & strTableWidth & """"
        strWidth = " WIDTH=""100%"""
      End If

      If strHeight > "" Then
        Response.Write " HEIGHT=""" & strTableHeight & """"
        strHeight = " HEIGHT=""100%"""
      End If

      Response.Write "><TR><TD>"

    Else

      If Style.BorderFormat = tbBdFormatOnlyLines Then

        If Spacing > 0 And Style.ExternalBorder.Color > "" Then

          'Alterado (Ridai Govinda)
          strLineRow = "<TR"
          If Not NoBorderColors Then
            strLineRow = strLineRow & " BGCOLOR=""" & Style.ExternalBorder.Color & """"
          End If
          strLineRow = strLineRow & "><TD HEIGHT=" & Spacing
          '------------------------

          If TableColumns > 1 Then
            strLineRow = strLineRow & " COLSPAN=" & TableColumns
          End If

          strLineRow = strLineRow & "></TD></TR>"

        End If

      End If

      'Alterado (Ridai Govinda)
      If Style.ExternalBorder.Color > "" And Style.BorderFormat = tbBdFormatCompleteGrid _
         And Not NoBorderColors Then '<----------
        strBorderColor = " BGCOLOR=""" & Style.ExternalBorder.Color & """"
      End If
      '-----------------------

    End If

    'Alterado (Ridai Govinda)
    If Not strBorderColor > "" And Style.BorderFormat = tbBdFormatCompleteGrid _
       And Not NoBorderColors Then '<----------
      strBorderColor = " BGCOLOR=""#000000"""
    End If
    '-----------------------

  End If

  If Session("TableBorder") > "" Then
    strBorderWidth = " BORDER=""" & Session("TableBorder") & """"
  Else
    If ShowBorder Then
      strBorderWidth = " BORDER=""1"""
    Else
      strBorderWidth = " BORDER=""0"""
    End If
  End If

  Response.Write "<TABLE" & strBorderColor & strBorderWidth

  If Padding > "" Then
    Response.Write " CELLPADDING=""" & Padding & """"
  End If

  If TableBackground > "" Then
    Response.Write " BACKGROUND=""" & TableBackground & """"
  End If

  Response.Write strSpacing & strWidth & strHeight & ">" & strLineRow

  If HaveTitle Then

    Response.Write "<TR"

    'Alterado/Criado (Ridai Govinda)
    If Style.ColorFormat = tbStFormatTitle And Not NoCellColors Then
      Response.Write " BGColor=""" & Style.BaseColor & """"
      Rem Response.Write " BGColor=""" & Style.Color1 & """"
    ElseIf Style.TitleColor > "" And Not NoCellColors Then
      Response.Write " BGCOLOR=""" & Style.TitleColor & """"
    End If

    If Style.TitleBackground > "" Then
      Response.Write " BACKGROUND=""" & Style.TitleBackground & """"
    End If
    '-----------------------

    Response.Write "><TD ALIGN=""CENTER"""

    If TableColumns > 1 Then
      Response.Write " COLSPAN=" & TableColumns
    End If

    Response.Write ">" & Style.TitleFont.BeginFont & strMessage & Style.TitleFont.EndFont & "</TD></TR>"

  End If

  ColumnBackGround = ""
  ColumnAlign = ""
  ColumnVAlign = ""
  ColumnColor = ""
  ColumnColSpan = ""
  ColumnWidth = ""
  ColumnHeight = ""

End Sub

Rem =========================================================================
Rem Fim da tabela
Rem -------------------------------------------------------------------------
Public Sub EndTable()

  Dim strLineRow As String

  'Criado (Ridai Govinda)
  sblnTableOpened = False
  '---------------------

  CloseCell
  CloseLine

  If Style.BorderFormat = tbBdFormatOnlyLines And Spacing > 0 And HaveBorder Then

    If Spacing > 0 And Style.ExternalBorder.Color > "" Then

      'Alterado (Ridai Govinda)
      strLineRow = "<TR"
      If Not NoBorderColors Then
        strLineRow = strLineRow & " BGCOLOR=""" & Style.ExternalBorder.Color & """"
      End If
      strLineRow = strLineRow & "><TD HEIGHT=" & Spacing
      '-----------------------

      If TableColumns > 1 Then
        strLineRow = strLineRow & " COLSPAN=" & TableColumns
      End If

      strLineRow = strLineRow & "></TD></TR>"

    End If

    Response.Write strLineRow

  End If

  If HaveSecondTable Then
    Response.Write "</TD></TR></TABLE>"
  End If

  clctTables.Remove clctTables.Count
  Response.Write "</TABLE>"

End Sub

'Criado (Ridai Govinda)
Public Sub BeginRowColumn(Optional ByVal intFontSize As Integer = 0, _
                          Optional ByVal blnHeadRow As Boolean = False)
  BeginRow intFontSize, blnHeadRow
  BeginColumn

End Sub
'---------------------

Rem =========================================================================
Rem Inicio de uma linha
Rem -------------------------------------------------------------------------
Public Sub BeginRow(Optional ByVal intFontSize As Integer = 0, _
                    Optional ByVal blnHeadRow As Boolean = False)
  Dim strColor, strLine, strBackground As String

  'Criado (Ridai Govinda)
  'If Not sblnTableOpened Then
  '  BeginTable
  'End If
  '----------------------

  CloseCell
  CloseLine
  ColumnNumber = 0
  RowNumber = RowNumber + 1
  RowOpened = True
  IsHeaderRow = blnHeadRow
  'Criado (Ridai Govinda)
  strBackground = ""
  '---------------------

  If IsHeaderRow Then
    Style.HeaderFont.Size = intFontSize
    'Alterado (Ridai Govinda)
    strColor = " BGColor=""" & Style.HeaderColor & """"
    Rem strColor = " BGColor=""" & Style.Color1 & """"
    '---------------
  Else

    'Alterado (Ridai Govinda)
    If intFontSize > 0 Then Style.InternalFont.Size = intFontSize
    '-----------------------

    Select Case Style.ColorFormat

      Case tbStFormatHeader, tbStFormatAlternateColumns
        'Alterado (Ridai Govinda)
        strColor = " BGColor=""" & Style.AlternateColor & """"
        Rem strColor = " BGColor=""" & Style.Color2 & """"
        '-----------------------

      Case tbStFormatHeaderAlternateColumns
        'Alterado (Ridai Govinda)
        strColor = " BGColor=""" & Style.FirstAltColor & """"
        Rem strColor = " BGColor=""" & Style.Color1 & """"
        '-----------------------

      Case tbStFormatHearderAlternateLines
        If (RowNumber Mod 2 = 1) Then
          'Alterado (Ridai Govinda)
          strColor = " BGColor=""" & Style.FirstAltColor & """"
          Rem strColor = " BGColor=""" & Style.Color1 & """"
          '-----------------------
        Else
          'Alterado (Ridai Govinda)
          strColor = " BGColor=""" & Style.AlternateColor & """"
          Rem strColor = " BGColor=""" & Style.Color2 & """"
          '-----------------------
        End If

      Case tbStFormatAlternateLines
        If (RowNumber Mod 2 = 0) Then
          'Alterado (Ridai Govinda)
          strColor = " BGColor=""" & Style.AlternateColor & """"
          Rem strColor = " BGColor=""" & Style.Color2 & """"
          '-----------------------
        Else
          'Alterado (Ridai Govinda)
          strColor = " BGColor=""" & Style.LastColor & """"
          Rem strColor = " BGColor=""" & Style.Color3 & """"
          '-----------------------
        End If

      Case tbStFormatNothing, tbStFormatTitle
        'Alterado (Ridai Govinda)
        strColor = " BGColor=""" & Style.BaseColor & """"
        Rem strColor = " BGColor=""" & Style.Color1 & """"
        '-----------------------

    End Select

  End If

  ' (Estilo de bordas:
  ' 0 - Grade completa
  ' 1 - Apenas linhas
  ' 2 - Grade externa + linhas)
  ' else - sem grade (fundo da página)
  ' Verificar:Não funciona
  ' 3 - Apenas colunas
  ' 4 - Grade externa + colunas)
  If ((Style.BorderFormat = tbBdFormatOnlyLines) Or (Style.BorderFormat = tbBdFormatOutterGridLines)) And HaveBorder Then

    If (LastRow > 0) Or (IsHeaderRow And HaveTitle) Then

      If (LastRow = 1 Or IsHeaderRow) And Style.HeaderBorder.Width > 0 And Style.HeaderBorder.Color > "" Then

        strLine = "<TR BGCOLOR=""" & Style.HeaderBorder.Color & """><TD HEIGHT=" & Style.HeaderBorder.Width

        If TableColumns > 1 Then
          strLine = strLine & " COLSPAN=" & TableColumns
        End If

        strLine = strLine & "></TD></TR>"

      ElseIf LastRow = 2 And Style.InternalBorder.Width > 0 And Style.InternalBorder.Color > "" Then

        strLine = "<TR BGCOLOR=""" & Style.InternalBorder.Color & """><TD HEIGHT=" & Style.InternalBorder.Width

        If TableColumns > 1 Then
          strLine = strLine & " COLSPAN=" & TableColumns
        End If

        strLine = strLine & "></TD></TR>"

      End If

    End If

    If IsHeaderRow Then
      LastRow = 1
    Else
      LastRow = 2
    End If

  End If

  'Alterado (Ridai)
  If Not IsHeaderRow Then
    If RowBackground > "" Then strBackground = " Background=" & RowBackground
  Else
    If Style.HeadRowBackground > "" Then strBackground = " Background=" & Style.HeadRowBackground
  End If

  If NoCellColors Then strColor = ""

  Response.Write strLine & "<TR" & strColor & strBackground & ">"
  '---------------

End Sub

Rem =========================================================================
Rem Fim de uma linha
Rem -------------------------------------------------------------------------
Public Sub EndRow()

  CloseCell
  RowOpened = False

  Response.Write "</TR>"

End Sub

Rem =========================================================================
Rem Inicio de uma celula
Rem -------------------------------------------------------------------------
Public Sub BeginColumn()
  BeginCell
End Sub


Rem =========================================================================
Rem Inicio de uma coluna
Rem -------------------------------------------------------------------------
Public Sub BeginCell()

  Dim strFont As String, strColor As String, strColumn As String, strBackground As String

  CloseCell
  ColumnNumber = ColumnNumber + 1
  ColumnOpened = True
  'Criado (Ridai Govinda)
  strBackground = ""
  '---------------------

  If IsHeaderRow Then
    strFont = Style.HeaderFont.BeginFont
  Else
    strFont = Style.InternalFont.BeginFont
  End If

  If ColumnColor > "" Then

    strColor = " BGCOLOR=""" & ColumnColor & """"

  Else

    If IsHeaderRow Then

      Select Case Style.ColorFormat
        Case tbStFormatHeaderAlternateColumns
          If (ColumnNumber Mod 2 = 0) Then
            'Alterado (Ridai Govinda)
            strColor = " BGColor=""" & Style.AlternateColor & """"
            Rem strColor = " BGColor=""" & Style.Color2 & """"
            '-----------------------
          End If
      End Select

    Else

      Select Case Style.ColorFormat

        Case tbStFormatHeaderAlternateColumns

          If (ColumnNumber Mod 2 = 0) Then
            'Alterado (Ridai Govinda)
            strColor = " BGColor=""" & Style.AlternateColor & """"
            Rem strColor = " BGColor=""" & Style.Color2 & """"
            '-----------------------
          End If

        Case tbStFormatAlternateColumns

          If (ColumnNumber Mod 2 = 0) Then
            'Alterado (Ridai Govinda)
            strColor = " BGColor=""" & Style.LastColor & """"
            Rem strColor = " BGColor=""" & Style.Color3 & """"
            '-----------------------
          End If

      End Select

    End If

  End If

  'Criado (Ridai Govinda)
  Select Case Style.BackgroundFormat
  Case tbStFormatAlternateColumns
    If (ColumnNumber Mod 2 = 0) Then
      strBackground = Style.FirstAltBackground

    Else
      strBackground = Style.SecondAltBackground

    End If

  Case tbStFormatAlternateLines
    If (RowNumber Mod 2 = 0) Then
      strBackground = Style.FirstAltBackground

    Else
      strBackground = Style.SecondAltBackground

    End If

  End Select

  If IsHeaderRow Then
    Response.Write strColumn & "<TH"
  Else
    Response.Write strColumn & "<TD"
  End If

  'Alterado (Ridai Govinda)
  If IsHeaderRow Then
    If Style.HeadColsBackground > "" Then
      Response.Write " BACKGROUND=""" & Style.HeadColsBackground & """"
    End If
  Else
    If ColumnBackGround > "" Then
      strBackground = ColumnBackGround
    End If

    If strBackground > "" Then Response.Write " BACKGROUND=""" & strBackground & """"

  End If
  '-----------------------

  If ColumnHeight > "" Then
    Response.Write " HEIGHT=""" & ColumnHeight & """"
  End If

  If ColumnWidth > "" Then
    Response.Write " WIDTH=""" & ColumnWidth & """"
  End If

  If ColumnAlign > "" Then
    Response.Write " ALIGN=""" & ColumnAlign & """"
  End If

  If ColumnColSpan > "" Then
    Response.Write " COLSPAN=""" & ColumnColSpan & """"
  End If

  If ColumnRowSpan > "" Then
    Response.Write " ROWSPAN=""" & ColumnRowSpan & """"
  End If

  If ColumnVAlign > "" Then
    Response.Write " VALIGN=""" & ColumnVAlign & """"
  End If

  If ColumnNoWrap Then
    Response.Write " NOWRAP=""NOWRAP"""
  End If

  'Alterado (Ridai Govinda)
  If NoCellColors Then strColor = ""
  Response.Write strColor & ">" & strFont
  '---------------------------------

  'Removido (Ridai Govinda)
  'ColumnBackGround = ""
  'ColumnColor = ""
  '-----------------------

End Sub


Rem =========================================================================
Rem Fim da celula (obsoleto)
Rem -------------------------------------------------------------------------
Public Sub EndColumn()
  EndCell
End Sub

Rem =========================================================================
Rem Fim da celula
Rem -------------------------------------------------------------------------
Public Sub EndCell()

  Dim strFont As String

  ColumnOpened = False

  If IsHeaderRow Then
    strFont = Style.HeaderFont.EndFont
  Else
    strFont = Style.InternalFont.EndFont
  End If

  'Alterado (Ridai)
  If IsHeaderRow Then
    Response.Write strFont & "</TH>"
  Else
    Response.Write strFont & "</TD>"
  End If
  '----------------------

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub ColumnGraph(ByVal intParc As Double, ByVal strLink As String)
  CellGraph intParc, strLink
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub CellGraph(ByVal intParc As Double, Optional ByVal strLink As String = "")

  Dim intLargura As Integer, intWidth As Integer
  'Criado (Ridai Govinda)
  Dim blnNoCellColors As Boolean, blnNoBorderColors As Boolean
  Dim strCellAlign As String, strCellWidth As String, strCellHeight As String
  Dim strPadding As String, strSpacing As String
  Dim intDiv As Long

  strCellAlign = CellAlign
  strCellWidth = CellWidth
  strCellHeight = CellHeight
  strPadding = Padding
  strSpacing = Spacing
  
  BeginCell

  blnNoCellColors = sblnNoCellColors
  blnNoBorderColors = sblnNoBorderColors

  sblnNoCellColors = True
  sblnNoBorderColors = True
  '-------------

  Padding = 0
  Spacing = 0
  BeginTable

  CellWidth = CellGraphTotalWidth
  If GraphTotal > 0 Then CellWidth = (CellGraphTotalWidth \ 3) * 2
  
  CellAlign = "Right"
  BeginRow 1

  If strLink > "" Then
    Cell strLink & "&nbsp;"
  Else
    Cell "&nbsp;" & Round(intParc, 2) & "&nbsp;"
  End If

  CellWidth = CellGraphTotalWidth \ 3
  
  If intParc < GraphMin Then intParc = GraphMin
  intParc = intParc - GraphMin

  intDiv = (GraphMax - GraphMin)
  If intDiv = 0 Then intDiv = 1
  
  intLargura = Round((CellGraphMeterWidth * intParc) / intDiv, 0)

  If intLargura < 1 Then
    intLargura = 1
  End If
  
  If GraphTotal > 0 Then
    If GraphTotal = 0 Then GraphTotal = 1
    
    Cell "&nbsp;" & CInt(100 * intParc / GraphTotal) & "%" & "&nbsp;"
  End If

  'Criado (Ridai Govinda)
  sblnNoCellColors = False
  '----------------------

  CellHeight = CellGraphHeight
  CellColor = GraphColor
  CellWidth = intLargura
  Cell "<IMG SRC=""" & CellGraphImage & """ WIDTH=""" & intLargura & """ HEIGHT=""" & CellGraphHeight - 1 & """>"
  'Criado (Ridai Govinda)
  CellColor = ""
  sblnNoCellColors = True
  '---------------------
  EndRow

  EndTable

  EndCell

  'Alterado (Ridai Govinda)
  sblnNoCellColors = blnNoCellColors
  sblnNoBorderColors = blnNoBorderColors
  '----------------------

  CellAlign = strCellAlign
  CellHeight = strCellHeight
  CellWidth = strCellWidth
  Padding = strPadding
  Spacing = strSpacing

  If sintGraphActual = sintGraphNum Then
    sintGraphActual = 1
  Else
    sintGraphActual = sintGraphActual + 1
  End If

End Sub

'Criado (Ridai Govinda)
Public Property Let CellGraphMeterDivider(ByVal strParam As String)
  CellGraphMeterRightDivider = strParam
  CellGraphMeterLeftDivider = strParam
  CellGraphMeterCenterDivider = strParam
End Property
'---------------------

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub LegendInfo(ByVal dblMax As Double, ByVal dblTotal As Double, Optional ByVal dblMin As Double = 0)

  sintGraphNum = sintGraphNum + 1
  sintGraphActual = sintGraphNum
  GraphTotal = dblTotal
  GraphMax = dblMax
  GraphMin = dblMin

  If Session(sintGraphActual & "sintCor") = "" Then
     Session(sintGraphActual & "sintCor") = -1
     Session(sintGraphActual & "sintR") = 255
     Session(sintGraphActual & "sintG") = -85
     Session(sintGraphActual & "sintB") = 0
  End If

Rem  Response.Write "sintGraphNum = " & sintGraphNum
Rem  Response.Write "<BR>sintGraphActual = " & sintGraphActual
Rem  Response.Write "<BR>Session(sintGraphActual & 'sintCor') = " & Session(sintGraphActual & "sintCor")
Rem  Response.Write "<BR>GraphTotal = " & GraphTotal
Rem  Response.Write "<BR>GraphMax = " & GraphMax

End Sub

Rem =========================================================================
Rem Desenha a Legenda do Grafico de Barras
Rem Revisado por: Ridai Govinda Pombo
Rem   Utilizando propriedades para alterar o comportamento default;
Rem -------------------------------------------------------------------------
Public Sub GraphLegend(ByVal dblMax As Double, ByVal dblTotal As Double, Optional ByVal dblMin As Double = 0)
  Dim i As Long
  Dim strCellAlign As String, strCellWidth As String
  Dim strPadding As String, strSpacing As String

  LegendInfo dblMax, dblTotal, dblMin

  strCellAlign = CellAlign
  strCellWidth = CellWidth
  strPadding = Padding
  strSpacing = Spacing
  
  CellAlign = "Left"
  BeginCell

  Padding = 0
  Spacing = 0
  BeginTable

  BeginRow 1

  CellAlign = "Left"

  CellWidth = CellGraphTotalWidth
  Cell IIf(GraphTotal > 0, "Total", "&nbsp;")

  CellWidth = CellGraphMeterWidth \ 3
  
  CellAlign = "LEFT"
  Cell GraphMin
  
  CellAlign = "CENTER"
  Cell (GraphMax + GraphMin) \ 2
  
  CellAlign = "RIGHT"
  Cell GraphMax
  
  BeginRow 1
  
  CellWidth = CellGraphTotalWidth
  CellAlign = "CENTER"
  Cell IIf(GraphTotal > 0, "&nbsp;<B>" & HTFormatInt(GraphTotal) & "<B>&nbsp;", "&nbsp;")
  
  CellWidth = CellGraphMeterWidth \ 3
  
  CellAlign = "LEFT"
  Cell CellGraphMeterLeftDivider
  
  CellAlign = "CENTER"
  Cell CellGraphMeterCenterDivider
  
  CellAlign = "RIGHT"
  Cell CellGraphMeterRightDivider
  
  EndRow

  EndTable

  EndCell

  CellAlign = strCellAlign
  CellWidth = strCellWidth
  Padding = strPadding
  Spacing = strSpacing

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub Column(ByVal strText As String)
  Cell strText
End Sub
Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub Cell(ByVal strText As String)
  BeginCell
  Response.Write strText
  EndCell

End Sub

Rem =========================================================================
Rem Adiciona um linha na tabela com o conteúdo do array informado separado por colunas
Rem -------------------------------------------------------------------------
Public Sub Row(ParamArray strColumns() As Variant)

  Dim i As Integer

  BeginRow

    For i = 0 To UBound(strColumns)
      Column strColumns(i)
    Next

    If IsMissing(strColumns) Then
      Column "&nbsp;"
    End If

  EndRow

End Sub

'Criado (Ridai Govinda)
Rem =========================================================================
Rem Adiciona uma linha de cabeçalho na tabela com o conteúdo do
Rem array informado separado por colunas
Rem -------------------------------------------------------------------------
Public Sub HeaderRow(ParamArray strColumns() As Variant)

  Dim i As Integer

  BeginRow , True

    For i = 0 To UBound(strColumns)
      Column strColumns(i)
    Next

    If IsMissing(strColumns) Then
      Column "&nbsp;"
    End If

  EndRow

End Sub
'------------

Private Function LeadingZeroes(ByVal strNumber As String, ByVal lngPlaces As Long) As String
  strNumber = Trim(strNumber)
  LeadingZeroes = String(lngPlaces - Len(strNumber), "0") & strNumber

End Function

Rem =========================================================================
Rem Retorna cores sequenciais
Rem -------------------------------------------------------------------------
Private Function GraphColor()

 If (Session(sintGraphActual & "sintCor") \ 3) = 6 Then
   Session(sintGraphActual & "sintCor") = 0

 End If

 Select Case (Session(sintGraphActual & "sintCor") \ 3)
   Case 0
     Session(sintGraphActual & "sintG") = Session(sintGraphActual & "sintG") + 85

   Case 1
     Session(sintGraphActual & "sintR") = Session(sintGraphActual & "sintR") - 85

   Case 2
     Session(sintGraphActual & "sintB") = Session(sintGraphActual & "sintB") + 85

   Case 3
     Session(sintGraphActual & "sintG") = Session(sintGraphActual & "sintG") - 85

   Case 4
     Session(sintGraphActual & "sintR") = Session(sintGraphActual & "sintR") + 85

   Case 5
     Session(sintGraphActual & "sintB") = Session(sintGraphActual & "sintB") - 85

 End Select
 Session(sintGraphActual & "sintCor") = Session(sintGraphActual & "sintCor") + 1

 GraphColor = "#" & LeadingZeroes(Hex(Session(sintGraphActual & "sintR")), 2) & _
                    LeadingZeroes(Hex(Session(sintGraphActual & "sintG")), 2) & _
                    LeadingZeroes(Hex(Session(sintGraphActual & "sintB")), 2)

End Function

'Criado (Ridai Govinda)
'A criar:
'Public Property Get GraphColors() As Collection
'  If sclctGColors Is Nothing Then Set sclctGColors = New Collection
'  Set GraphColors = sclctGColors
'End Property

Public Property Get ColumnBackGround() As String
  ColumnBackGround = Style.ColumnBackGround
End Property

Public Property Let ColumnBackGround(ByVal strBG As String)
  Style.ColumnBackGround = strBG
End Property

Public Property Get RowBackground() As String
  RowBackground = Style.RowBackground
End Property

Public Property Let RowBackground(ByVal strBG As String)
  Style.RowBackground = strBG
End Property

Public Property Get TableBackground() As String
  TableBackground = Style.TableBackground
End Property

Public Property Let TableBackground(ByVal strBG As String)
  Style.TableBackground = strBG
End Property

Public Property Get Style() As HiperTools30.TableStyle
  Set Style = sobjTableStyle
End Property

Public Property Set Style(objStyle As HiperTools30.TableStyle)
  Set sobjTableStyle = objStyle
End Property

Public Property Let NoCellColors(ByVal blnParam As Boolean)
  sblnNoCellColors = blnParam
End Property

Public Property Get NoCellColors() As Boolean
  NoCellColors = sblnNoCellColors
End Property

Public Property Let NoBorderColors(ByVal blnParam As Boolean)
  sblnNoBorderColors = blnParam
End Property

Public Property Get NoBorderColors() As Boolean
  NoBorderColors = sblnNoBorderColors
End Property
'-------------------
