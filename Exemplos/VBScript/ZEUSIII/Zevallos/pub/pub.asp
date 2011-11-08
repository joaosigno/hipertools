<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE VIRTUAL="/HiperTools/HiperTools30.inc"-->
<!--#INCLUDE VIRTUAL="/Objects.inc"-->
<%
REM =========================================================================
REM /pub.asp
REM -------------------------------------------------------------------------
REM Nome     : Abin
REM Descricao: 
REM Home     : localhost
REM Criacao  : 7/12/01 11:24:55 AM
REM Autor    :  - 
REM Versao   : 1.0
REM Local    :  - DF
REM Companhia: Zevallos
REM -------------------------------------------------------------------------

  Const conScriptTimeout  = 15
  Const conSessionTimeout = 300

  Const conWhatDocumentos                = "12"
  Const conWhatDocumentosTipo            = "13"
  Const conWhatNoticiasCategoria         = "15"
  Const conWhatNoticias                  = "14"
  Const conWhatPaginas                   = "16"
  Const conWhatPaginasLayout             = "17"

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
REM Procedimento que mostra a primeira página                                
REM -------------------------------------------------------------------------
Private Sub ShowFirstPage

  Table.Spacing = 0

  URL.BeginURL Initializer.ScriptName
  URL.Add efQueryStrAction, efQSActionEditor
  URL.Add efQueryStrEditable, "1"

    Table.BeginTable "45%", "Painel de Controle", 2, True
      Table.CellAlign = "center"
      Table.Row URL.GetURL( "Documentos",                efQueryStrWhat & "=" & conWhatDocumentos )
      Table.Row URL.GetURL( "DocumentosTipo",            efQueryStrWhat & "=" & conWhatDocumentosTipo )
      Table.Row URL.GetURL( "NoticiasCategoria",         efQueryStrWhat & "=" & conWhatNoticiasCategoria )
      Table.Row URL.GetURL( "Noticias",                  efQueryStrWhat & "=" & conWhatNoticias )
      Table.Row URL.GetURL( "Paginas",                   efQueryStrWhat & "=" & conWhatPaginas )
      Table.Row URL.GetURL( "PaginasLayout",             efQueryStrWhat & "=" & conWhatPaginasLayout )
    Table.EndTable

  URL.EndURL

End Sub
REM -------------------------------------------------------------------------
REM Fim do ShowFirstPage                                                     
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela "Documentos"
REM -------------------------------------------------------------------------
Public Sub dataDocumentos

  Edit.DataTable "pubDocumentos"

  Edit.DataAddField "docCodigo",                  efDataTypeInt,        4, efRequired
  Edit.DataAddField "docNome",                    efDataTypeVarChar,  200, efNotRequired
  Edit.DataAddField "docResumo",                  efDataTypeVarChar, 1000, efNotRequired
  Edit.DataAddField "docDescricao",               efDataTypeText,    6000, efNotRequired
  Edit.DataAddField "docURL",                     efDataTypeVarChar,  200, efNotRequired
  Edit.DataAddField "docTipoDocumento",           efDataTypeInt,        4, efNotRequired
  Edit.DataAddField "docInclusao",                efDataTypeDateTime,   10, efNotRequired
  Edit.DataAddField "docAlteracao",               efDataTypeDateTime,   10, efNotRequired
  Edit.DataAddField "docAtivo",                   efDataTypeTinyInt,    1, efNotRequired

  Show.Message "Não foi definido um campo primário para esta tabela!"

  If ( Not Edit.HasTable( "pubDocumentos" ) ) Then
      Edit.CreateTable "pubDocumentos"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim do dataDocumentos
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela "DocumentosTipo"
REM -------------------------------------------------------------------------
Public Sub dataDocumentosTipo

  Edit.DataTable "pubDocumentosTipo"

  Edit.DataAddField "dotCodigo",                  efDataTypeInt,        4, efRequired
  Edit.DataAddField "dotNome",                    efDataTypeVarChar,  200, efNotRequired
  Edit.DataAddField "dotDescricao",               efDataTypeVarChar,  500, efNotRequired
  Edit.DataAddField "dotInclusao",                efDataTypeDateTime,   10, efRequired
  Edit.DataAddField "dotAlteracao",               efDataTypeDateTime,   10, efNotRequired
  Edit.DataAddField "dotAtivo",                   efDataTypeTinyInt,    1, efNotRequired

  Show.Message "Não foi definido um campo primário para esta tabela!"

  If ( Not Edit.HasTable( "pubDocumentosTipo" ) ) Then
      Edit.CreateTable "pubDocumentosTipo"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim do dataDocumentosTipo
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela "Noticias"
REM -------------------------------------------------------------------------
Public Sub dataNoticias

  Edit.DataTable "pubNoticias"

  Edit.DataAddField "notCodigo",                  efDataTypeInt,        4, efRequired
  Edit.DataAddField "notCategoria",               efDataTypeInt,        4, efNotRequired
  Edit.DataAddField "notNome",                    efDataTypeVarChar,  200, efNotRequired
  Edit.DataAddField "notResumo",                  efDataTypeVarChar, 1000, efNotRequired
  Edit.DataAddField "notDescricao",               efDataTypeText,    6000, efNotRequired
  Edit.DataAddField "notReferencia",              efDataTypeVarChar, 1000, efNotRequired
  Edit.DataAddField "notInclusao",                efDataTypeDateTime,   10, efNotRequired
  Edit.DataAddField "notAlteracao",               efDataTypeDateTime,   10, efNotRequired
  Edit.DataAddField "notAtivo",                   efDataTypeTinyInt,    1, efNotRequired

  Show.Message "Não foi definido um campo primário para esta tabela!"

  If ( Not Edit.HasTable( "pubNoticias" ) ) Then
      Edit.CreateTable "pubNoticias"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim do dataNoticias
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela "NoticiasCategoria"
REM -------------------------------------------------------------------------
Public Sub dataNoticiasCategoria

  Edit.DataTable "pubNoticiasCategoria"

  Edit.DataAddField "nocCodigo",                  efDataTypeInt,        4, efRequired
  Edit.DataAddField "nocNome",                    efDataTypeVarChar,  200, efNotRequired
  Edit.DataAddField "nocDescricao",               efDataTypeVarChar,  500, efNotRequired
  Edit.DataAddField "nocInclusao",                efDataTypeDateTime,   10, efNotRequired
  Edit.DataAddField "nocAlteracao",               efDataTypeDateTime,   10, efNotRequired
  Edit.DataAddField "nocAtivo",                   efDataTypeTinyInt,    1, efNotRequired

  Show.Message "Não foi definido um campo primário para esta tabela!"

  If ( Not Edit.HasTable( "pubNoticiasCategoria" ) ) Then
      Edit.CreateTable "pubNoticiasCategoria"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim do dataNoticiasCategoria
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela "Paginas"
REM -------------------------------------------------------------------------
Public Sub dataPaginas

  Edit.DataTable "pubPaginas"

  Edit.DataAddField "pagCodigo",                  efDataTypeInt,        4, efRequired
  Edit.DataAddField "pagPai",                     efDataTypeInt,        4, efNotRequired
  Edit.DataAddField "pagNome",                    efDataTypeVarChar,  200, efNotRequired
  Edit.DataAddField "pagDescricao",               efDataTypeText,    6000, efNotRequired
  Edit.DataAddField "pagReferencia",              efDataTypeVarChar,  500, efNotRequired
  Edit.DataAddField "pagLayout",                  efDataTypeInt,        4, efNotRequired
  Edit.DataAddField "pagInclusao",                efDataTypeDateTime,   10, efNotRequired
  Edit.DataAddField "pagAlteracao",               efDataTypeDateTime,   10, efNotRequired
  Edit.DataAddField "pagAtivo",                   efDataTypeTinyInt,    1, efNotRequired

  Show.Message "Não foi definido um campo primário para esta tabela!"

  If ( Not Edit.HasTable( "pubPaginas" ) ) Then
      Edit.CreateTable "pubPaginas"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim do dataPaginas
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela "PaginasLayout"
REM -------------------------------------------------------------------------
Public Sub dataPaginasLayout

  Edit.DataTable "pubPaginasLayout"

  Edit.DataAddField "palCodigo",                  efDataTypeInt,        4, efRequired
  Edit.DataAddField "palNome",                    efDataTypeVarChar,  200, efNotRequired
  Edit.DataAddField "palImagem",                  efDataTypeVarChar,  100, efNotRequired
  Edit.DataAddField "palImagemBackground",        efDataTypeVarChar,  100, efNotRequired
  Edit.DataAddField "palBackgroundStye",          efDataTypeVarChar,  500, efNotRequired
  Edit.DataAddField "palInclusao",                efDataTypeDateTime,   10, efNotRequired
  Edit.DataAddField "palAlteracao",               efDataTypeDateTime,   10, efNotRequired
  Edit.DataAddField "palAtivo",                   efDataTypeTinyInt,    1, efNotRequired

  Show.Message "Não foi definido um campo primário para esta tabela!"

  If ( Not Edit.HasTable( "pubPaginasLayout" ) ) Then
      Edit.CreateTable "pubPaginasLayout"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim do dataPaginasLayout
REM =========================================================================

REM =========================================================================
REM Inicio do procedimento que desenha o formulário de edição da tabela
REM "Documentos"
REM -------------------------------------------------------------------------
Public Sub formDocumentos

  dataDocumentos

  Edit.FormBegin  "pubDocumentos", "Documentos", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "docCodigo;docNome;docResumo;docDescricao;docURL;docTipoDocumento;docInclusao;docAlteracao;docAtivo"
    Edit.FormFind "docCodigo,docNome,docResumo,docDescricao,docURL,docTipoDocumento,docInclusao,docAlteracao,docAtivo"
    Edit.FormList "docCodigo,docNome,docResumo,docDescricao,docURL,docTipoDocumento,docInclusao,docAlteracao,docAtivo"

      Edit.AddField "docCodigo", "Codigo"
      Edit.AddField "docNome", "Nome"
      Edit.AddField "docResumo", "Resumo"
      Edit.AddField "docDescricao", "Descricao"
      Edit.AddField "docURL", "URL"
      Edit.AddField "docTipoDocumento", "TipoDocumento"
      Edit.AddField "docInclusao", "Inclusao"
      Edit.AddField "docAlteracao", "Alteracao"
      Edit.AddField "docAtivo", "Ativo"

  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do "Documentos
REM =========================================================================

REM =========================================================================
REM Inicio do procedimento que desenha o formulário de edição da tabela
REM "DocumentosTipo"
REM -------------------------------------------------------------------------
Public Sub formDocumentosTipo

  dataDocumentosTipo

  Edit.FormBegin  "pubDocumentosTipo", "DocumentosTipo", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "dotCodigo;dotNome;dotDescricao;dotInclusao;dotAlteracao;dotAtivo"
    Edit.FormFind "dotCodigo,dotNome,dotDescricao,dotInclusao,dotAlteracao,dotAtivo"
    Edit.FormList "dotCodigo,dotNome,dotDescricao,dotInclusao,dotAlteracao,dotAtivo"

      Edit.AddField "dotCodigo", "Codigo"
      Edit.AddField "dotNome", "Nome"
      Edit.AddField "dotDescricao", "Descricao"
      Edit.AddField "dotInclusao", "Inclusao"
      Edit.AddField "dotAlteracao", "Alteracao"
      Edit.AddField "dotAtivo", "Ativo"

  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do "DocumentosTipo
REM =========================================================================

REM =========================================================================
REM Inicio do procedimento que desenha o formulário de edição da tabela
REM "Noticias"
REM -------------------------------------------------------------------------
Public Sub formNoticias

  dataNoticias

  Edit.FormBegin  "pubNoticias", "Noticias", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "notCodigo;notCategoria;notNome;notResumo;notDescricao;notReferencia;notInclusao;notAlteracao;notAtivo"
    Edit.FormFind "notCodigo,notCategoria,notNome,notResumo,notDescricao,notReferencia,notInclusao,notAlteracao,notAtivo"
    Edit.FormList "notCodigo,notCategoria,notNome,notResumo,notDescricao,notReferencia,notInclusao,notAlteracao,notAtivo"

      Edit.AddField "notCodigo", "Codigo"
      Edit.AddField "notCategoria", "Categoria"
      Edit.AddField "notNome", "Nome"
      Edit.AddField "notResumo", "Resumo"
      Edit.AddField "notDescricao", "Descricao"
      Edit.AddField "notReferencia", "Referencia"
      Edit.AddField "notInclusao", "Inclusao"
      Edit.AddField "notAlteracao", "Alteracao"
      Edit.AddField "notAtivo", "Ativo"

  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do "Noticias
REM =========================================================================

REM =========================================================================
REM Inicio do procedimento que desenha o formulário de edição da tabela
REM "NoticiasCategoria"
REM -------------------------------------------------------------------------
Public Sub formNoticiasCategoria

  dataNoticiasCategoria

  Edit.FormBegin  "pubNoticiasCategoria", "NoticiasCategoria", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "nocCodigo;nocNome;nocDescricao;nocInclusao;nocAlteracao;nocAtivo"
    Edit.FormFind "nocCodigo,nocNome,nocDescricao,nocInclusao,nocAlteracao,nocAtivo"
    Edit.FormList "nocCodigo,nocNome,nocDescricao,nocInclusao,nocAlteracao,nocAtivo"

      Edit.AddField "nocCodigo", "Codigo"
      Edit.AddField "nocNome", "Nome"
      Edit.AddField "nocDescricao", "Descricao"
      Edit.AddField "nocInclusao", "Inclusao"
      Edit.AddField "nocAlteracao", "Alteracao"
      Edit.AddField "nocAtivo", "Ativo"

  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do "NoticiasCategoria
REM =========================================================================

REM =========================================================================
REM Inicio do procedimento que desenha o formulário de edição da tabela
REM "Paginas"
REM -------------------------------------------------------------------------
Public Sub formPaginas

  dataPaginas

  Edit.FormBegin  "pubPaginas", "Paginas", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "pagCodigo;pagPai;pagNome;pagDescricao;pagReferencia;pagLayout;pagInclusao;pagAlteracao;pagAtivo"
    Edit.FormFind "pagCodigo,pagPai,pagNome,pagDescricao,pagReferencia,pagLayout,pagInclusao,pagAlteracao,pagAtivo"
    Edit.FormList "pagCodigo,pagPai,pagNome,pagDescricao,pagReferencia,pagLayout,pagInclusao,pagAlteracao,pagAtivo"

      Edit.AddField "pagCodigo", "Codigo"
      Edit.AddField "pagPai", "Pai"
      Edit.AddField "pagNome", "Nome"
      Edit.AddField "pagDescricao", "Descricao"
      Edit.AddField "pagReferencia", "Referencia"
      Edit.AddField "pagLayout", "Layout"
      Edit.AddField "pagInclusao", "Inclusao"
      Edit.AddField "pagAlteracao", "Alteracao"
      Edit.AddField "pagAtivo", "Ativo"

  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do "Paginas
REM =========================================================================

REM =========================================================================
REM Inicio do procedimento que desenha o formulário de edição da tabela
REM "PaginasLayout"
REM -------------------------------------------------------------------------
Public Sub formPaginasLayout

  dataPaginasLayout

  Edit.FormBegin  "pubPaginasLayout", "PaginasLayout", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "palCodigo;palNome;palImagem;palImagemBackground;palBackgroundStye;palInclusao;palAlteracao;palAtivo"
    Edit.FormFind "palCodigo,palNome,palImagem,palImagemBackground,palBackgroundStye,palInclusao,palAlteracao,palAtivo"
    Edit.FormList "palCodigo,palNome,palImagem,palImagemBackground,palBackgroundStye,palInclusao,palAlteracao,palAtivo"

      Edit.AddField "palCodigo", "Codigo"
      Edit.AddField "palNome", "Nome"
      Edit.AddField "palImagem", "Imagem"
      Edit.AddField "palImagemBackground", "ImagemBackground"
      Edit.AddField "palBackgroundStye", "BackgroundStye"
      Edit.AddField "palInclusao", "Inclusao"
      Edit.AddField "palAlteracao", "Alteracao"
      Edit.AddField "palAtivo", "Ativo"

  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do "PaginasLayout
REM =========================================================================

REM =========================================================================
REM Altera o estilo do Objeto Table                                          
REM -------------------------------------------------------------------------
Private Sub FormatTable

  Table.Style.BaseColor            = ""
  Table.Style.HeaderColor          = "Orange"
  Table.Style.FirstAltColor        = "#ECECEC"
  Table.Style.AlternateColor       = "#ECECEC"
  Table.Style.LastColor            = "#ECECEC"
  Table.Style.BorderColor          = "Orange"
  Table.Style.BorderFormat         = tbBdFormatOnlyLines
  Table.Style.ColorFormat          = tbStFormatTitle
  Table.Style.ExternalBorder.Width = 3
  Table.Style.HeaderBorder.Width   = 2
  Table.Style.InternalBorder.Width = 1

  Set Edit.Style = Table.Style

End Sub
REM -------------------------------------------------------------------------
REM Fim do FormatTable                                                       
REM =========================================================================

REM =========================================================================
REM Corpo Principal do sistema                                               
REM -------------------------------------------------------------------------
Private Sub MainBody

  FormatTable

  Edit.ConnectionString = Session( "ConnectionString" )

  Edit.OpenConnection

  Select Case Edit.parWhat

    Case conWhatDocumentos
         formDocumentos

    Case conWhatDocumentosTipo
         formDocumentosTipo

    Case conWhatNoticiasCategoria
         formNoticiasCategoria

    Case conWhatNoticias
         formNoticias

    Case conWhatPaginas
         formPaginas

    Case conWhatPaginasLayout
         formPaginasLayout

    Case Else

  End Select

  Edit.RedirectActions

  Default.BeginHTML
  Default.BeginBody
  Default.BeginBody

	   If ( Edit.IsMyAction ) Then
	      ShowFirstPage
	   End If

  Default.PageFooterDefault
  Default.EndBody
  Default.EndHTML

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody                                                    
REM =========================================================================

REM -------------------------------------------------------------------------
REM Fim do pub.asp
REM =========================================================================
%>
