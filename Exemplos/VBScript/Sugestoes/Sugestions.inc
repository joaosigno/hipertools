<%
REM *************************************************************************
REM  /Sugestions.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Procedimentos e funções auxiliares à execução do sistema.
REM  Criação     : 2000/04/27 12:18AM
REM  Local       : Brasília/DF
REM  Elaborado   : Kelson da Costa Medeiros
REM  Versao      : 1.0.0
REM  Copyright   : 1999 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------


REM =========================================================================
REM Configura as tabelas
REM -------------------------------------------------------------------------
Private Sub FormatTable()

  If ( Request.QueryString(efQueryStrAction) = efQSActionList ) Then
     Edit.Style.ColorFormat        = tbStFormatNothing
  Else
     Edit.Style.ColorFormat        = tbStFormatAlternateLines
  End If

  Table.Style.ColorFormat          = tbStFormatAlternateColumns
  Table.Style.BorderFormat         = tbBdFormatOnlyLines
  Table.Style.Color1               = "#ECECEC"
  Table.Style.Color2               = "#FFFFFF"
  Table.Style.Color3               = "#ECECEC"
  Table.Style.BorderColor          = "Orange"
  Table.Style.ExternalBorder.Width = 3
  Table.Style.HeaderBorder.Width   = 1
  Table.Style.InternalBorder.Width = 1

  Set Edit.Style = Table.Style

End Sub
REM -------------------------------------------------------------------------
REM Fim do FortatTable
REM =========================================================================

REM =========================================================================
REM Monta a Frame
REM -------------------------------------------------------------------------
Public Sub ShowFrames

  Show.HTMLCR "<html>"
  Show.HTMLCR " <head>"
  Show.HTMLCR "  <title>Sugestões para o HiperTools 30</title>"
  Show.HTMLCR " </head>"
  Show.HTMLCR " <frameset rows=""100%"" cols=""190,*"" framespacing=0>"
  Show.HTMLCR "  <frame name=""Menu"" src=""Sugestions_Menu.asp"" scrolling=""auto"" marginheight=""0"" marginwidth=""0"" frameborder=""0"" Resize=""on"">"
  Show.HTMLCR "  <frame name=""Body"" src=""Sugestions.asp"" scrolling=""auto"" marginheight=""0"" marginwidth=""0"" frameborder=""0"" >"
  Show.HTMLCR " </frameset>"
  Show.HTMLCR "</html>"

End Sub
REM -------------------------------------------------------------------------
REM Fim do Frame
REM =========================================================================

REM =========================================================================
REM Configura um link para acesso direto as telas do EditForm
REM -------------------------------------------------------------------------
Public Function ConfigEditForm( ByVal strLocalization, ByVal strWhat, ByVal strAction, ByVal strEditable, ByVal strFind, ByVal strFilter, ByVal strList, ByVal strDefaults, ByVal strTab )

  If Not strLocalization > "" Then
     Exit Function
  End If

  URL.BeginURL strLocalization

    If strWhat > "" Then
       URL.Add efQueryStrWhat, strWhat
    End If

    If strAction > "" Then
     URL.Add efQueryStrAction, strAction
    End If

    If strEditable > "" Then
      URL.Add efQueryStrEditable, strEditable
    End If

    If strFind > "" Then
       URL.Add efQueryStrFind, strFind
    End If

    If strFilter > "" Then
       URL.Add efQueryStrFilter, strFilter
    End If

    If strList > "" Then
       URL.Add efQueryStrList, strList
    End If

    If strDefaults > "" Then
       URL.Add efQueryStrDefaults, strDefaults
    End If

    If strTab > "" Then
       URL.Add efQueryStrTab, strTab
    End If

    ConfigEditForm = URL.GetOnlyURL

  URL.EndURL

End Function
REM -------------------------------------------------------------------------
REM Fim do ConfigEditForm
REM =========================================================================

REM -------------------------------------------------------------------------
REM Fim do /Sugestions.inc
REM *************************************************************************
%>  