<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE VIRTUAL="/HiperTools/HiperTools30.inc"-->
<!--#INCLUDE VIRTUAL="/HiperTools/Objects.inc"-->
<%
REM =========================================================================
REM /Novidades.asp
REM -------------------------------------------------------------------------
REM Nome     : Novidades
REM Descricao: Sistema de novidades do Site zevallos
REM Home     : www.zevallos2.com.br/Novidades
REM Criacao  : 1/26/00 5:53:49 PM
REM Autor    : Eduardo Gonçalves (Zeus III)
REM          : Ruben Zevallos (Estruturas de Dados)
REM          : Fernando Aquino (Desenvolvimento e Estruturas de Dados)
REM Versao   : 1
REM Local    :  - DF
REM Companhia: Zevallos
REM -------------------------------------------------------------------------

  Const conScriptTimeout     = 15
  Const conSessionTimeout    = 300
  Const conWhatNovidades     = "1"
  Const conWhatTipoNovidades = "2"

  Main

REM =========================================================================
REM Rotina Principal do Sistema
REM -------------------------------------------------------------------------
Private Sub Main

  Server.ScriptTimeOut = conScriptTimeout
  Session.TimeOut      = conSessionTimeout

  MainBody

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Main
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela "TipoNovidades"
REM -------------------------------------------------------------------------
Public Sub dataTipoNovidades

  Edit.DataBegin "zsnTipoNovidades"

  Edit.DataAddField "tnvCodigo", efDataTypeInt, 4, efRequired
  Edit.DataAddField "tnvSigla", efDataTypeVarChar, 20, efRequired
  Edit.DataAddField "tnvNome", efDataTypeVarChar, 100, efRequired
  Edit.DataAddField "tnvDescricao", efDataTypeText, 6000, efNotRequired

  Edit.DataAddPrimaryKey "tnvCodigo"

  If ( Not Edit.IsTable( "zsnTipoNovidades" ) ) Then
      Edit.TableCreate "zsnTipoNovidades"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim do dataTipoNovidades
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela "Novidades"
REM -------------------------------------------------------------------------
Public Sub dataNovidades

  Edit.DataBegin "zsnNovidades"

  Edit.DataAddField "novCodigo", efDataTypeInt, 4, efRequired
  Edit.DataAddField "novTipoNovidade", efDataTypeInt, 4, efRequired
  Edit.DataAddField "novCliente", efDataTypeInt, 4, efNotRequired
  Edit.DataAddField "novProduto", efDataTypeInt, 4, efNotRequired
  Edit.DataAddField "novLayout", efDataTypeInt, 4, efNotRequired
  Edit.DataAddField "novAcessos", efDataTypeInt, 10, efNotRequired
  Edit.DataAddField "novDataCriacao", efDataTypeDateTime, 10, efRequired
  Edit.DataAddField "novDatUltmAcess", efDataTypeDateTime, 10, efNotRequired
  Edit.DataAddField "novDatAlteracao", efDataTypeDateTime, 10, efNotRequired
  Edit.DataAddField "novIdUsuarioAlterac", efDataTypeInt, 4, efNotRequired
  Edit.DataAddField "novIdUsuarioCriador", efDataTypeInt, 4, efNotRequired
  Edit.DataAddField "novTitulo", efDataTypeVarChar, 100, efRequired
  Edit.DataAddField "novSubTitulo", efDataTypeVarChar, 100, efNotRequired
  Edit.DataAddField "novAutor", efDataTypeVarChar, 100, efNotRequired
  Edit.DataAddField "novReferencia", efDataTypeVarChar, 255, efNotRequired
  Edit.DataAddField "novMailAutor", efDataTypeChar, 100, efNotRequired
  Edit.DataAddField "novUrlReferencia", efDataTypeChar, 100, efNotRequired
  Edit.DataAddField "novSaibamaisUrl", efDataTypeChar, 100, efNotRequired
  Edit.DataAddField "novImagem", efDataTypeVarChar, 250, efNotRequired
  Edit.DataAddField "novPalavrasChaves", efDataTypeVarChar, 60, efNotRequired
  Edit.DataAddField "novTexto", efDataTypeText, 6000, efNotRequired
  Edit.DataAddField "novResumo", efDataTypeText, 6000, efNotRequired

  Edit.DataAddPrimaryKey "novCodigo"

  If ( Not Edit.IsTable( "zsnNovidades" ) ) Then
      Edit.TableCreate "zsnNovidades"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim do dataNovidades
REM =========================================================================

REM =========================================================================
REM Inicio do procedimento que desenha o formulário de edição da tabela
REM "TipoNovidades"
REM -------------------------------------------------------------------------
Public Sub formTipoNovidades

  dataTipoNovidades

  Edit.FormBegin  "zsnTipoNovidades", "Tipo de Novidades", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "tnvSigla;tnvNome;tnvDescricao"
    Edit.FormFind "tnvCodigo,tnvSigla,tnvNome,tnvDescricao"
    Edit.FormList "tnvCodigo,tnvSigla,tnvNome,tnvDescricao"

      Edit.AddField "tnvCodigo",    "&Código",    , , ,efNext
      Edit.AddField "tnvSigla",     "&Sigla"
      Edit.AddField "tnvNome",      "&Nome"
      Edit.AddField "tnvDescricao", "&Descrição", efFldTypeTextArea

      Edit.FieldInternalLink "tnvNome", Edit.parWhat, "tnvCodigo"

  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do "TipoNovidades
REM =========================================================================

REM =========================================================================
REM Inicio do procedimento que desenha o formulário de edição da tabela
REM "Novidades"
REM -------------------------------------------------------------------------
Public Sub formNovidades

  dataNovidades

  Edit.FormBegin  "zsnNovidades", "Novidades", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "novTipoNovidade,novCliente;novTitulo;novSubTitulo;novTexto;novResumo;novPalavrasChaves;novImagem;novAutor;novReferencia;novMailAutor;novUrlReferencia;novSaibamaisUrl"
    Edit.FormFind "novCodigo,novTipoNovidade,novCliente,novProduto,novLayout,novAcessos,novDataCriacao,novDatUltmAcess,novDatAlteracao,novIdUsuarioAlterac,novIdUsuarioCriador,novTitulo,novSubTitulo,novAutor,novReferencia,novMailAutor,novUrlReferencia,novSaibamaisUrl,novImagem,novPalavrasChaves,novTexto,novResumo"
    Edit.FormList "novCodigo,novTipoNovidade,novCliente,novProduto,novLayout,novAcessos,novDataCriacao,novDatUltmAcess,novDatAlteracao,novIdUsuarioAlterac,novIdUsuarioCriador,novTitulo,novSubTitulo,novAutor,novReferencia,novMailAutor,novUrlReferencia,novSaibamaisUrl,novImagem,novPalavrasChaves,novTexto,novResumo"

    REM -------------------------------------------------------------------------
    REM Campos reservados para futura implementação:
    REM novProduto,novLayout
    REM -------------------------------------------------------------------------

      Edit.AddField "novCodigo",           "&Código",               , , ,efNext
      Edit.AddField "novTipoNovidade",     "&Típo de Novidade",     efFldTypeLookup
      Edit.AddField "novCliente",          "&Cliente",              efFldTypeLookup
      Edit.AddField "novProduto",          "&Código do Produto"
      Edit.AddField "novLayout",           "&Layout"
      Edit.AddField "novAcessos",          "&Acessos"
      Edit.AddField "novDataCriacao",      "&Data Criação",         , efValOptDate, ,Strings.LongDate(Now)
      Edit.AddField "novDatUltmAcess",     "&Data Último Acesso",   , efValOptDate
      Edit.AddField "novDatAlteracao",     "&Data Alteração",       , efValOptDate, ,Strings.LongDate(Now)
      Edit.AddField "novIdUsuarioAlterac", "&ID do Usuario Alteração"
      Edit.AddField "novIdUsuarioCriador", "&ID do Usuário Criador"
      Edit.AddField "novTitulo",           "&Título"
      Edit.AddField "novSubTitulo",        "&Sub-Título"
      Edit.AddField "novAutor",            "&Autor"
      Edit.AddField "novReferencia",       "&Referência (Máximo de 255 Caracteres)", efFldTypeTextArea
      Edit.AddField "novMailAutor",        "&E-mail do Autor"
      Edit.AddField "novUrlReferencia",    "&URL da Referência",    efFldTypeHTTP
      Edit.AddField "novSaibamaisUrl",     "&Saiba mais",           efFldTypeHTTP
      Edit.AddField "novImagem",           "&Imagem",               efFldTypeImage
      Edit.AddField "novPalavrasChaves",   "&Palavras Chaves"
      Edit.AddField "novTexto",            "&Texto",                efFldTypeTextArea
      Edit.AddField "novResumo",           "&Resumo",               efFldTypeTextArea

      Edit.FieldInternalLink "novCodigo", Edit.parWhat, "novTitulo"

      Edit.FieldImage "novImagem", "Novidades\Img", 50, "Novidade"

      Edit.FieldLookUp "novCliente",      "zcsClientes",      "cliCodigo", "cliSigla"
      Edit.FieldLookUp "novTipoNovidade", "zsnTipoNovidades", "tnvCodigo", "tnvSigla"

  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do "Novidades
REM =========================================================================

REM =========================================================================
REM Altera o estilo do Objeto Table
REM -------------------------------------------------------------------------
Private Sub FormatTable

  Table.Style.BackgroundFormat     = tbStFormatNothing
  Table.Style.BorderFormat         = tbBdFormatInvisible
'  Table.Style.ColorFormat          = tbStFormatTitle
  Table.Style.ColorFormat          = tbStFormatNothing
  Table.Style.AlternateColor       = "#0079BD"
'  Table.Style.BaseColor            = "#0079BD"
  Table.Style.BaseColor            = ""
  Table.Style.BorderColor          = "black"
  Table.Style.HeaderFont.Color     = "yellow"
  Table.Style.InternalBorder.Color = "green"

  Set Edit.Style = Table.Style

End Sub
REM -------------------------------------------------------------------------
REM Fim do FormatTable
REM =========================================================================

REM =========================================================================
REM Procedimento que retorna o caminho da URL completo sem o nome do arquivo
REM apenas
REM -------------------------------------------------------------------------
Private Function TranslateSiteRoot

  Dim strReverse
  strReverse = Strings.Reverse( Initializer.ScriptURL )
  TranslateSiteRoot = Strings.Reverse( Mid( strReverse, InStr( strReverse, "/" ) ) )

End Function
REM -------------------------------------------------------------------------
REM Fim do TranslateSiteRoot
REM =========================================================================

REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub MainBody

  FormatTable

  Default.BodyBGColor         = "#0079BD"
  Default.BodyBackground      = TranslateSiteRoot & "/img/bgnovidades.gif"
  Default.BodyText            = "#000000"
  Default.BodyLink            = "#0000FF"
  Default.BodyVLink           = "darkblue"
  Default.BodyALink           = "red"
  Default.LinkStyleSheetHRef  = ""
  Default.BodyTopMargin       = 0

  Edit.OpenConnection

  Select Case Edit.parWhat

    Case conWhatNovidades
         formNovidades

    Case conWhatTipoNovidades
         formTipoNovidades

    Case Else
         Response.Redirect "News_Admin.asp?EA=h01&EE=1&EW=" & conWhatNovidades
  End Select

  Edit.RedirectActions


  Default.BeginHTML
  Default.BeginBody
  Default.WriteHead

    Show.Center
    Table.Style.BaseColor = "Orange" '"#0079BD"
    Table.BeginTable "100%"
      Table.CellWidth = "50%"
      Table.CellAlign = "center"
      URL.BeginURL Initializer.ScriptName
        URL.Add efQueryStrAction,   efQSActionEditor
        URL.Add efQueryStrEditable, "1"
        Table.Row URL.GetURL( "Novidades", efQueryStrWhat & "=" & conWhatNovidades ), _
                  URL.GetURL( "Tipos de Novidades", efQueryStrWhat & "=" & conWhatTipoNovidades )
      URL.EndURL
      Table.EndRow
      Table.CellWidth = ""
    Table.EndTable
    Show.CenterEnd

    Edit.ShowUnitUpperBar = False

    Table.Style.BaseColor = ""
    Edit.IsMyAction

  Default.PageFooterDefault
  Default.EndBody
  Default.EndHTML

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================

REM -------------------------------------------------------------------------
REM Fim do ZeusIII.asp
REM =========================================================================
%>