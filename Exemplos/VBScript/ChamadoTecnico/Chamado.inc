<%
REM =========================================================================
REM Altera o estilo do Objeto Table
REM -------------------------------------------------------------------------
Private Sub FormatTable()

  REM Estilos de cores:
  REM -------------------------------------------------------------------------
  REM tbStFormatHeader                 - Cabecalho cor 1, resto cor 2
  REM tbStFormatHeaderAlternateColumns - Alterna colunas
  REM tbStFormatAlternateColumns       - Cabecalho fixo, alterna colunas
  REM tbStFormatHearderAlternateLines  - Alterna linhas
  REM tbStFormatAlternateLines         - Cabecalho fixo, alterna linhas
  REM tbStFormatNothing                - Todo da cor 1(menos o T�tulo)
  REM tbStFormatTitle                  - Todo da cor 1(inclusive T�tulo)

  REM Estilo de bordas:
  REM -------------------------------------------------------------------------
  REM tbBdFormatCompleteGrid    - Grade completa
  REM tbBdFormatOnlyLines       - Apenas linhas
  REM tbBdFormatOutterGridLines - Grade externa + linhas
  REM tbBdFormatInvisible       - Invis�vel

  REM Estilos de background(s)
  REM -------------------------------------------------------------------------
  REM tbStFormatAlternateColumns       - Cabecalho fixo, alterna colunas
  REM tbStFormatAlternateLines         - Cabecalho fixo, alterna linhas
  REM tbStFormatNothing                - Sem configura��o

  Table.Style.BackGroundFormat     = tbStFormatNothing
  Table.Style.ColorFormat          = tbStFormatHeader
  Table.Style.BorderFormat         = tbBdFormatOnlyLines

  Table.Style.HeadColsBackGround   = "img\cellbg.jpg"

  Table.Style.Color1               = "Orange"
  Table.Style.BaseColor            = "#ECECEC"
  Table.Style.HeaderColor          = ""
  Table.Style.FirstAltColor        = "#ECECEC"
  Table.Style.AlternateColor       = "#FFFFFF"
  Table.Style.LastColor            = "#ECECEC"
  Table.Style.BorderColor          = "Orange"
  Table.Style.ExternalBorder.Width = 3
  Table.Style.HeaderBorder.Width   = 1
  Table.Style.InternalBorder.Width = 1

  Set Edit.Style = Table.Style

End Sub
REM -------------------------------------------------------------------------
REM Fim do FormatTable
REM =========================================================================
%>