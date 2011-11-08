<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE VIRTUAL="/HiperTools/HiperTools30.inc"-->
<!--#INCLUDE VIRTUAL="/HiperTools/Objects.inc"-->
<!--#INCLUDE FILE="Constants.inc"-->
<!--#INCLUDE FILE="MakeSystem.inc"-->
<!--#INCLUDE FILE="Import.inc"-->
<%
REM =========================================================================
REM  /ZeusIII.asp
REM ------------------------------------------------------------------------
REM Nome     : Gerador de Sistemas ZEUS III
REM Descricao: Menu para o Sistema ZEUS III
REM Home     : http://www.hipertools.com.br/
REM Criacao  : 2/12/0 5:14PM
REM Autor    : Zevallos Tecnologia em Informacao
REM Versao   : 1.1.0.0
REM Local    : Brasília - DF
REM Copyright: 97-2000 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

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
REM Altera o estilo do Objeto Table
REM -------------------------------------------------------------------------
Private Sub FormatTable

  Table.Style.BaseColor            = ""
  Table.Style.HeaderColor          = "Blue"
  Table.Style.FirstAltColor        = "#ECECEC"
  Table.Style.AlternateColor       = "#ECECEC"
  Table.Style.LastColor            = "#ECECEC"
  Table.Style.BorderColor          = "Orange"
  Table.Style.ExternalBorder.Width = 3
  Table.Style.HeaderBorder.Width   = 2
  Table.Style.InternalBorder.Width = 1
  Table.Style.BorderFormat         = 1
  Table.Style.ColorFormat          = 6

  Set Edit.Style = Table.Style

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormatTable
REM =========================================================================

REM =========================================================================
REM Monta a Frame
REM -------------------------------------------------------------------------
Public Sub ShowFrames

  Default.BodyWidth = 620

  Show.HTML "<html>"
  Show.HTML "<head>"
  Show.HTML "	<title>ZEUS III v1.1</title>"
  Show.HTML "</head>"
  Show.HTML "<frameset cols=""180,*"">"
  Show.HTML "	<frame name=""Menu"" src=""ZEUSIIIMenu.asp"" scrolling=""auto"" marginheight=""0"" marginwidth=""0"">"
  Show.HTML "	<frame name=""Body"" src=""ZeusIII.asp?EE=1&EA=h02&EW=1"" scrolling=""auto"" marginheight=""0"" marginwidth=""0"">"
  Show.HTML "</frameset>"
  Show.HTML "</html>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Frame
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela de "Empresas"
REM -------------------------------------------------------------------------
Public Sub tableZeuEmpresas

  Edit.DataTable "zeuEmpresas"

  Edit.DataAddField "empCodigo",       efDataTypeInt,      4, efRequired
  Edit.DataAddField "empNome",         efDataTypeVarChar, 50, efRequired
  Edit.DataAddField "empRazao_Social", efDataTypeVarChar, 80, efRequired
  Edit.DataAddField "empEndereco",     efDataTypeVarChar, 80, efNotRequired
  Edit.DataAddField "empCidade",       efDataTypeVarChar, 30, efNotRequired
  Edit.DataAddField "empBairro",       efDataTypeVarChar, 50, efNotRequired
  Edit.DataAddField "empUF",           efDataTypeVarChar, 02, efNotRequired
  Edit.DataAddField "empCEP",          efDataTypeVarChar, 09, efNotRequired
  Edit.DataAddField "empTelefone",     efDataTypeVarChar, 17, efNotRequired
  Edit.DataAddField "empContato",      efDataTypeVarChar, 30, efNotRequired
  Edit.DataAddField "empHomePage",     efDataTypeVarChar, 60, efNotRequired
  Edit.DataAddField "empEmail",        efDataTypeVarChar, 30, efNotRequired

  Edit.DataAddPrimaryKey "empCodigo"

  Edit.DataAddIndex "idxNome", "empNome", True

  If ( Not Edit.HasTable( "zeuEmpresas" ) ) Then
     Edit.CreateTable "zeuEmpresas"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final do procedimento que define a estrutura de dados tabela de >Empresas
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela de "Sistemas"
REM -------------------------------------------------------------------------
Public Sub tableZeuSistemas

  If ( Not Edit.HasTable( "zeuEmpresas" ) ) Then
     tableZeuEmpresas
  End If

  Edit.DataTable "zeuSistemas"

  Edit.DataAddField "sisEmpresa",     efDataTypeInt,       4, efRequired
  Edit.DataAddField "sisCodigo",      efDataTypeInt,       4, efRequired
  Edit.DataAddField "sisSigla",       efDataTypeVarChar,   3, efRequired
  Edit.DataAddField "sisNome",        efDataTypeVarChar,  80, efRequired
  Edit.DataAddField "sisVersao",      efDataTypeVarChar,  10, efRequired
  Edit.DataAddField "sisHome",        efDataTypeVarChar, 100, efRequired
  Edit.DataAddField "sisResponsavel", efDataTypeVarChar,  30, efNotRequired
  Edit.DataAddField "sisTelefone",    efDataTypeVarChar,  17, efNotRequired
  Edit.DataAddField "sisDescricao",   efDataTypeVarChar, 200, efNotRequired

  Edit.DataAddPrimaryKey "sisCodigo"

  Edit.DataAddIndex "idxNome",  "sisEmpresa,sisNome", True
  Edit.DataAddIndex "idxSigla", "sisSigla", True

  If ( Not Edit.HasTable( "zeuSistemas" ) ) Then
     Edit.CreateTable "zeuSistemas"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final do procedimento que define a estrutura de dados tabela de >Sistemas
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela de "tabelas"
REM -------------------------------------------------------------------------
Public Sub tableZeuTabelas

  If ( Not Edit.HasTable( "zeuSistemas" ) ) Then
     tableZeuSistemas
  End If

  Edit.DataTable "zeuTabelas"

  Edit.DataAddField "tabSistema",   efDataTypeInt,       4, efRequired
  Edit.DataAddField "tabCodigo",    efDataTypeInt,       4, efRequired
  Edit.DataAddField "tabSigla",     efDataTypeVarChar,   3, efRequired
  Edit.DataAddField "tabNome",      efDataTypeVarChar,  50, efRequired
  Edit.DataAddField "tabDescricao", efDataTypeVarChar, 200, efNotRequired

  Edit.DataAddPrimaryKey "tabCodigo"

  Edit.DataAddIndex "idxNome",  "tabSistema,tabNome",  True
  Edit.DataAddIndex "idxSigla", "tabSistema,tabSigla", True

  If ( Not Edit.HasTable( "zeuTabelas" ) ) Then
     Edit.CreateTable "zeuTabelas"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final do procedimento que define a estrutura de dados tabela de "tabelas"
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela de "campos"
REM -------------------------------------------------------------------------
Public Sub tableZeuCampos

  If ( Not Edit.HasTable( "zeuTabelas" ) ) Then
     tableZeuTabelas
  End If

  Edit.DataTable "zeuCampos"

  Edit.DataAddField "camTabela",          efDataTypeInt,       4, efRequired
  Edit.DataAddField "camCodigo",          efDataTypeInt,       4, efRequired
  Edit.DataAddField "camNome",            efDataTypeVarChar,  50, efRequired
  Edit.DataAddField "camTipo",            efDataTypeTinyInt,   2, efRequired
  Edit.DataAddField "camTamanho",         efDataTypeInt,       2, efRequired
  Edit.DataAddField "camRequerido",       efDataTypeTinyInt,   2, efNotRequired
  Edit.DataAddField "camPrimario",        efDataTypeTinyInt,   2, efNotRequired
  Edit.DataAddField "camOrdem",           efDataTypeTinyInt,   2, efRequired
  Edit.DataAddField "camRotulo",          efDataTypeVarChar, 200, efRequired
  Edit.DataAddField "camDelimitador",     efDataTypeVarChar,   1, efNotRequired
  Edit.DataAddField "camTexto",           efDataTypeVarChar, 200, efNotRequired
  Edit.DataAddField "camTipoEdicao",      efDataTypeTinyInt,   2, efNotRequired
  Edit.DataAddField "camTipoValidacao",   efDataTypeTinyInt,   2, efNotRequired
  Edit.DataAddField "camListagem",        efDataTypeTinyInt,   2, efNotRequired
  Edit.DataAddField "camMostrar",         efDataTypeTinyInt,   2, efNotRequired
  Edit.DataAddField "camLocalizacao",     efDataTypeTinyInt,   2, efNotRequired
  Edit.DataAddField "camMaximo",          efDataTypeInt,       2, efNotRequired
  Edit.DataAddField "camMsgObrigatorio",  efDataTypeVarChar, 200, efNotRequired
  Edit.DataAddField "camPadrao",          efDataTypeVarChar, 200, efNotRequired
  Edit.DataAddField "camDescricao",       efDataTypeVarChar, 200, efNotRequired

  Edit.DataAddPrimaryKey "camCodigo"

  Edit.DataAddIndex "idxNome",  "camTabela,camNome",  True
  Edit.DataAddIndex "idxOrdem", "camTabela,camOrdem", True

  If ( Not Edit.HasTable( "zeuCampos" ) ) Then
     Edit.CreateTable "zeuCampos"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final do procedimento que define a estrutura de dados tabela de "campos"
REM =========================================================================

REM =========================================================================
REM Procedimento que desenha o formulário de edição de "empresas"
REM -------------------------------------------------------------------------
Public Sub GetEmpresas

  tableZeuEmpresas

  Edit.BeginForm  "zeuEmpresas", "Cadastro de Empresas", 1, Edit.parWhat, _
                  efValLocClient
    Edit.FormUnit "empNome;empRazao_Social;empEndereco;empCidade," & _
                  "empBairro,empUF,empCEP;empTelefone,empContato," & _
                  "empEMail, ;empHomePage"
    Edit.FormFind "empNome,empRazao_Social,empCidade,empBairro,empUF," & _
                  "empCEP,empContato"
    Edit.FormList "empCodigo,empNome,empRazao_Social,empCidade," & _
                  "empBairro,empUF,empCEP,empContato"
      Edit.AddField "empCodigo",        "",           efFldTypeText, , , efNext
      Edit.AddField "empNome",         "&Nome"
      Edit.AddField "empRazao_Social", "&Razão Social"
      Edit.AddField "empEndereco",     "&Endereço"
      Edit.AddField "empBairro",       "&Bairro"
      Edit.AddField "empCidade",       "C&idade"
      Edit.AddField "empUF",           "&UF",         efFldTypeUF
      Edit.AddField "empCEP",          "&CEP",        efFldTypeText, efValOptCEP
      Edit.AddField "empTelefone",     "&Telefone"
      Edit.AddField "empContato",      "C&ontato"
      Edit.AddField "empHomePage",     "&HomePage",   efFldTypeHTTP
      Edit.AddField "empEMail",        "E-&Mail",     efFldTypeEmail, efValOptEmail

        Edit.FieldShowSize "empBairro", 20

    Edit.FieldInternalLink "empNome", Edit.parWhat, "empCodigo"
    Edit.AddReport "empUF,#Total(gc,1=Total)", "Relatório por unidade federativa", conListEmpresaUF

  Edit.EndForm

End Sub
REM -------------------------------------------------------------------------
REM Fim do procedimento que desenha o formulário de edição de "empresas"
REM =========================================================================

REM =========================================================================
REM Fim do procedimento que desenha o formulário de edição de "sistemas"
REM -------------------------------------------------------------------------
Public Sub GetSistemas

  tableZeuSistemas

  Edit.BeginForm "zeuSistemas", "Cadastro de Sistemas", 1, Edit.parWhat, efValLocClient
    Edit.FormUnit   "sisSigla,sisNome,sisVersao;sisHome;sisResponsavel,sisTelefone;sisDescricao"
    Edit.FormFind   "sisSigla,sisNome,sisResponsavel,sisTelefone,sisDescricao"
    Edit.FormList   "sisSigla,sisNome,sisResponsavel,sisTelefone,sisDescricao"

      Edit.AddField "sisEmpresa", ""
      Edit.AddField "sisCodigo",       "", , , , efNext
      Edit.AddField "sisSigla",        "&Sigla"
      Edit.AddField "sisNome",         "&Nome"
      Edit.AddField "sisVersao",       "&Versao"
      Edit.AddField "sisHome",         "&Home"
      Edit.AddField "sisResponsavel",  "&Responsável"
      Edit.AddField "sisTelefone",     "&Telefone"
      Edit.AddField "sisDescricao",    "&Descrição", efFldTypeTextArea

        Edit.FieldCharCase "sisSigla",       efCharCaseLower
        Edit.FieldShowSize "sisSigla",       5
        Edit.FieldShowSize "sisNome",        71
        Edit.FieldShowSize "sisResponsavel", 80
        Edit.FieldShowSize "sisDescricao",   102

    Edit.FieldInternalLink "sisNome", Edit.parWhat, "sisCodigo"

  Edit.EndForm

End Sub
REM -------------------------------------------------------------------------
REM Fim do procedimento que desenha o formulário de edição de "sistemas"
REM =========================================================================

REM =========================================================================
REM Procedimento que desenha o formulário de edição de "tabelas"
REM -------------------------------------------------------------------------
Public Sub GetTabelas

  tableZeuTabelas

  Edit.BeginForm    "zeuTabelas", "Cadastro de Tabelas", 1, Edit.parWhat, _
                    efValLocClient
    Edit.FormUnit   "tabSigla,tabNome;tabDescricao"
    Edit.FormList   "tabSigla,tabNome,tabDescricao"
    Edit.FormFind   "tabSigla,tabNome,tabDescricao"

     Edit.AddField "tabSistema", ""
     Edit.AddField "tabCodigo",    "",           , , , efNext
     Edit.AddField "tabSigla",     "&Sigla"
     Edit.AddField "tabNome",      "&Nome"
     Edit.AddField "tabDescricao", "&Descrição", efFldTypeTextArea

        Edit.FieldCharCase "tabSigla",     efCharCaseLower
        Edit.FieldShowSize "tabNome",      82
        Edit.FieldShowSize "tabDescricao", 94

    Edit.FieldInternalLink "tabNome", Edit.parWhat, "tabCodigo"

  Edit.EndForm

End Sub
REM -------------------------------------------------------------------------
REM Fim do procedimento que desenha o formulário de edição de "tabelas"
REM =========================================================================

REM =========================================================================
REM Procedimento que desenha o formulário de edição de "campos"
REM -------------------------------------------------------------------------
Public Sub GetCampos

  tableZeuCampos

  Edit.BeginForm    "zeuCampos", "Cadastro de Campos", 1, Edit.parWhat, _
                    efValLocClient
    Edit.FormList   "camOrdem,camNome,camRotulo,camTipo,camTamanho,camRequerido"
    Edit.FormUnit   "camNome,camTipo,camTamanho,camRequerido,camPrimario=EditForm;camOrdem,camRotulo,camDelimitador," & _
                    "camTexto;camListagem,camMostrar,camLocalizacao, ;camTipoEdicao,camTipoValidacao,camMaximo, ;" & _
                    "camMsgObrigatorio,camPadrao;camDescricao"
    Edit.FormFind   "camNome,camRotulo,camTipo,camTamanho,camRequerido,camDescricao"
      Edit.AddField "camTabela", ""
      Edit.AddField "camCodigo", "", , , , efNext
      Edit.AddField "camNome", "&Nome"
      Edit.AddField "camTipo", "&Tipo", efFldTypeSelect, , , 1
        Edit.FieldAddValue "camTipo",  0, "Texto"
        Edit.FieldAddValue "camTipo",  1, "Texto variável"
        Edit.FieldAddValue "camTipo",  2, "Data e hora"
        Edit.FieldAddValue "camTipo",  3, "Memorando"
        Edit.FieldAddValue "camTipo",  4, "Inteiro longo"
        Edit.FieldAddValue "camTipo",  5, "Inteiro"
        Edit.FieldAddValue "camTipo",  6, "Inteiro curto"
        Edit.FieldAddValue "camTipo",  7, "Ponto flutuante"
        Edit.FieldAddValue "camTipo",  8, "Real"
        Edit.FieldAddValue "camTipo",  9, "Monetário"
        Edit.FieldAddValue "camTipo", 10, "Lógico"
        Edit.FieldAddValue "camTipo", 11, "Incremental"
      Edit.AddField "camTamanho",     "&Tamanho"
      Edit.AddField "camRequerido",   "&Requerido", efFldTypeCheck, , , 0
        Edit.FieldAddValue "camRequerido", 1, "Requerido"
        Edit.FieldAddValue "camRequerido", 0, ""
      Edit.AddField "camPrimario",    "&Primário", efFldTypeCheck, , , 0
        Edit.FieldAddValue "camPrimario", 1, 1
        Edit.FieldAddValue "camPrimario", 0, 0
      Edit.AddField "camOrdem",       "&Ordem"
      Edit.AddField "camRotulo",      "&Rótulo"
			Edit.AddField "camDelimitador", "Delimitador", efFldTypeSelect, , , ";"
        Edit.FieldAddValue "camDelimitador", ",", ","
        Edit.FieldAddValue "camDelimitador", ";", ";"
        Edit.FieldAddValue "camDelimitador", "-", "-"
        Edit.FieldAddValue "camDelimitador", "=", "="
        Edit.FieldAddValue "camDelimitador", "|", "|"
			Edit.AddField "camTexto",       "Texto"
			Edit.AddField "camListagem",    "Listagem", efFldTypeCheck, , , 1
        Edit.FieldAddValue "camListagem", 1, 1
        Edit.FieldAddValue "camListagem", 0, 0
			Edit.AddField "camMostrar",     "Mostrar", efFldTypeCheck, , , 1
        Edit.FieldAddValue "camMostrar", 1, 1
        Edit.FieldAddValue "camMostrar", 0, 0
			Edit.AddField "camLocalizacao", "Localizacao", efFldTypeCheck, , , 1
        Edit.FieldAddValue "camLocalizacao", 1, 1
        Edit.FieldAddValue "camLocalizacao", 0, 0
			Edit.AddField "camMaximo",         "Maximo"
			Edit.AddField "camMsgObrigatorio", "Mensagem &Obrigatório"
			Edit.AddField "camPadrao",         "Padrao"
			Edit.AddField "camTipoEdicao",     "TipoEdicao", efFldTypeSelect
        Edit.FieldAddValue "camTipoEdicao",  1, "Texto"
        Edit.FieldAddValue "camTipoEdicao",  2, "UF"
        Edit.FieldAddValue "camTipoEdicao",  3, "Lookup"
        Edit.FieldAddValue "camTipoEdicao",  4, "Check"
        Edit.FieldAddValue "camTipoEdicao",  5, "Memorando"
        Edit.FieldAddValue "camTipoEdicao",  6, "Radio"
        Edit.FieldAddValue "camTipoEdicao",  7, "Select"
        Edit.FieldAddValue "camTipoEdicao",  8, "Senha"
        Edit.FieldAddValue "camTipoEdicao",  9, "HTTP"
        Edit.FieldAddValue "camTipoEdicao", 10, "EMail"
        Edit.FieldAddValue "camTipoEdicao", 11, "Data separada"
        Edit.FieldAddValue "camTipoEdicao", 12, "Imagem"
        Edit.FieldAddValue "camTipoEdicao", 13, "Arquivo"
        Edit.FieldAddValue "camTipoEdicao", 14, "Atualização"
        Edit.FieldAddValue "camTipoEdicao", 15, "Cor"
			Edit.AddField "camTipoValidacao",  "TipoValidacao", efFldTypeSelect
        Edit.FieldAddValue "camTipoValidacao",  1, "Nenhum"
        Edit.FieldAddValue "camTipoValidacao",  2, "CGC"
        Edit.FieldAddValue "camTipoValidacao",  3, "CPF"
        Edit.FieldAddValue "camTipoValidacao",  4, "Data"
        Edit.FieldAddValue "camTipoValidacao",  5, "Data separada"
        Edit.FieldAddValue "camTipoValidacao",  6, "Data maior que hoje"
        Edit.FieldAddValue "camTipoValidacao",  7, "Data separada maior que hoje"
        Edit.FieldAddValue "camTipoValidacao",  8, "Hora"
        Edit.FieldAddValue "camTipoValidacao",  9, "Email"
        Edit.FieldAddValue "camTipoValidacao", 10, "Compara datas"
        Edit.FieldAddValue "camTipoValidacao", 11, "CEP"
      Edit.AddField "camDescricao",      "&Descrição", efFldTypeTextArea
        Edit.FieldShowSize "camNome",           45
        Edit.FieldShowSize "camTamanho",        8
        Edit.FieldShowSize "camDescricao",      94
        Edit.FieldShowSize "camMsgObrigatorio", 60
        Edit.FieldShowSize "camPadrao",         60
        Edit.FieldShowSize "camRotulo",         45
        Edit.FieldShowSize "camTexto",          45
    Edit.FieldInternalLink "camNome", Edit.parWhat, "camCodigo"

    Edit.AddReport "camTipo,#Total(gc,1=Total)",        "Relatórios por tipos de campos",       conListCampoTipo
    Edit.AddReport "camDelimitador,#Total(gc,1=Total)", "Relatório por tipos de delimitadores", conListCampoDelimitador

  Edit.EndForm

End Sub
REM -------------------------------------------------------------------------
REM Fim do procedimento que desenha o formulário de edição de "campos"
REM =========================================================================

REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub MainBody

  If Request.QueryString( conOptions ) = conOptionTable Then
     Response.Cookies( "ServerName"   ) =  Request.Form( "Server"       )
     Response.Cookies( "UserName"     ) =  Request.Form( "Userid"       )
     Response.Cookies( "DatabaseName" ) =  Request.Form( "Database"     )
     Response.Cookies( "DatabaseType" ) =  Request.Form( "DatabaseType" )
  End If

  FormatTable

  Edit.ConnectionString = Session( "ConnectionString" )

  Edit.OpenConnection

  Select Case Edit.parWhat

    Case conWhatEmpresa
         GetEmpresas

    Case conWhatSistema
         GetSistemas

    Case conWhatTabela
         GetTabelas

    Case conWhatCampo
         GetCampos

    Case Else
         If ( Not Request.QueryString( conToMakeSystem ) > "" ) And ( Not Request.QueryString( conOptions ) > "" ) Then
            ShowFrames
            Response.Redirect "/ZeusIII.asp?EE=1&EA=h02&EW=1"
         End If

  End Select

  Edit.RedirectActions

  Default.BodyText  = "navy"
  Default.BodyLink  = "navy"
  Default.BodyVLink = "#0B8D94"
  Default.BodyALink = "#0B8D94"

  Default.BeginHTML
  Default.HeadAll "Zeus III"
  Default.BeginBody

  	If ( Edit.IsMyAction ) Then

       If ( Request.QueryString( conToMakeSystem ) > "" ) Then
          DoMakeSystem Request.QueryString( conToMakeSystem )
       Else

         Show.Center

         Select Case Request.QueryString( conOptions )

           Case conOptionInitImport
                ShowImportLogon

           Case conOptionTable
                ShowTables

           Case conOptionRename
                ShowRenameTables

           Case conOptionImport
                ImportTables

         End Select

       End If

    End If

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