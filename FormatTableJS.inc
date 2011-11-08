<%
// =========================================================================
// Altera o estilo do Objeto Table
// -------------------------------------------------------------------------
function FormatTable(){

  // Estilos de cores:
  // -------------------------------------------------------------------------
  // tbStFormatHeader                 - Cabecalho cor 1, resto cor 2
  // tbStFormatHeaderAlternateColumns - Alterna colunas
  // tbStFormatAlternateColumns       - Cabecalho fixo, alterna colunas
  // tbStFormatHearderAlternateLines  - Alterna linhas
  // tbStFormatAlternateLines         - Cabecalho fixo, alterna linhas
  // tbStFormatNothing                - Todo da cor 1(menos o Título)
  // tbStFormatTitle                  - Todo da cor 1(inclusive Título)

  // Estilo de bordas:
  // -------------------------------------------------------------------------
  // tbBdFormatCompleteGrid    - Grade completa
  // tbBdFormatOnlyLines       - Apenas linhas
  // tbBdFormatOutterGridLines - Grade externa + linhas
  // tbBdFormatInvisible       - Invisível

  // Estilos de background(s)
  // -------------------------------------------------------------------------
  // tbStFormatAlternateColumns       - Cabecalho fixo, alterna colunas
  // tbStFormatAlternateLines         - Cabecalho fixo, alterna linhas
  // tbStFormatNothing                - Sem configuração

  Table.Style.BackGroundFormat     = tbStFormatAlternateColumns;
  Table.Style.ColorFormat          = tbStFormatAlternateLines;
  Table.Style.BorderFormat         = tbBdFormatOnlyLines;

  Table.Style.HeadColsBackGround   = "cellbg.jpg";

  Table.Style.BaseColor            = "#ECECEC";
  Table.Style.HeaderColor          = "#ECECEC";
  Table.Style.FirstAltColor        = "#ECECEC";
  Table.Style.AlternateColor       = "#FFFFFF";
  Table.Style.LastColor            = "#ECECEC";
  Table.Style.BorderColor          = "Orange";
  Table.Style.ExternalBorder.Width = 3;
  Table.Style.HeaderBorder.Width   = 1;
  Table.Style.InternalBorder.Width = 1;

  Edit.Style = Table.Style;
}
// -------------------------------------------------------------------------
// Fim do FormatTable
// =========================================================================
%>