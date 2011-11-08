<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE VIRTUAL="/HiperTools/HiperTools30.inc"-->
<!--#INCLUDE VIRTUAL="/HiperTools/Objects.inc"-->
<!--#INCLUDE FILE="Constants.inc"-->
<%
REM =========================================================================
REM  Entregas.asp
REM -------------------------------------------------------------------------
REM  Descricao   : Sistema de Entregas
REM  Cria‡ao     : 4/21/0 8:41PM
REM  Local       : Brasilia/DF
REM  Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM              : Eduardo Alves Cortes <edualves@zevallos.com.br>
REM  Versao      : 2.0.0
REM  Copyright   : 2000 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

Main

REM =========================================================================
REM Rotina Principal do Sistema
REM -------------------------------------------------------------------------
Private Sub Main
  Server.ScriptTimeOut = conScriptTimeout
  MainBody
  Server.ScriptTimeOut = Session("ScriptTimeOut")

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Main
REM =========================================================================

REM =========================================================================
REM Altera o estilo do Objeto Table
REM -------------------------------------------------------------------------
Private Sub FormatTable

  Table.Style.BaseColor            = ""
  Table.Style.HeaderColor          = "LightBlue"
  Table.Style.FirstAltColor        = "#ECECEC"
  Table.Style.AlternateColor       = "#ECECEC"
  Table.Style.LastColor            = "#ECECEC"
  Table.Style.BorderColor          = "Blue"
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

  Session("BodyWidth") = "620"

  Show.Html "<html>"
  Show.Html "<head>"
  Show.Html "	<title>Entregas</title>"
  Show.Html "</head>"
  Show.Html "<frameset cols=""180,*"">"
  Show.Html "	<frame name=""Menu"" src=""EntregasMenu.asp"" scrolling=""auto"" marginheight=""0"" marginwidth=""0"">"
  Show.Html "	<frame name=""Body"" src=""Entregas.asp?EE=1&EA=h02&EW=6"" scrolling=""auto"" marginheight=""0"" marginwidth=""0"">"
  Show.Html "</frameset>"
  Show.Html "</html>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Frame
REM =========================================================================

REM =========================================================================
REM Mostra a Primeira Pagina
REM -------------------------------------------------------------------------
Private Sub ShowFirstPage

  Table.Spacing = 4
  Table.Padding = 4
  Table.ColumnNoWrap = True
  Table.BeginTable "40%"

  Table.BeginRow 3, True
  Table.Column "Formulários"
  Table.EndRow

  Table.BeginRow 2
  Table.ColumnVAlign = "Top"
  Table.BeginColumn

  EditCreateURL "Funções", conOptionFuncoes
  Show.BR

  EditCreateURL "Índices", conOptionIndices
  Show.BR

  EditCreateURL "Valores", conOptionValores
  Show.BR

  EditCreateURL "Produtos", conOptionProdutos
  Show.BR

  EditCreateURL "Fornecedores", conOptionFornecedores
  Show.BR

  EditCreateURL "Funcionários", conOptionFuncionarios
  Show.BR

  EditCreateURL "Clientes", conOptionClientes
  Show.BR

  Table.EndRow
  Table.EndTable

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FistPage
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela de "Indices"
REM -------------------------------------------------------------------------
Public Sub TableTesIndices

  Edit.DataBegin "tesIndices"

  Edit.DataAddField "indCodigo",    efDataTypeInt,      4, efRequired
  Edit.DataAddField "indDescricao", efDataTypeVarChar, 25, efNotRequired

  Edit.DataAddPrimaryKey "indCodigo"
  Edit.DataIndexClustered "Codigo", "indCodigo"

  If ( Not Edit.IsTable( "tesIndices" ) ) Then
     Edit.TableCreate "tesIndices"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final do procedimento que define a estrutura de dados tabela de >Empresas
REM =========================================================================

REM =========================================================================
REM Procedimento que desenha o formulário de edição de "Indices"
REM -------------------------------------------------------------------------
Public Sub GetIndices

  TableTesIndices
  
  Edit.FormBegin "tesIndices", "Indices", 1, conOptionIndices, efValLocClient

    Edit.FormUnit "indDescricao"
    Edit.FormList "indDescricao"
  
    Edit.AddField "indCodigo",    "&Código", , , , efNext
    Edit.AddField "indDescricao", "&Descrição", , , , , "A descrição do índice deve ser preenchida"
  
    Edit.FieldInternalLink "indDescricao", conOptionIndices, "indCodigo"
  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do procedimento que desenha o formulário de edição de "Indices"
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela de "Funcoes"
REM -------------------------------------------------------------------------
Public Sub TableTesFuncoes

  Edit.DataBegin "tesFuncoes"

  Edit.DataAddField "fncCodigo",    efDataTypeInt,     4 , efRequired
  Edit.DataAddField "fncDescricao", efDataTypeVarchar, 30, efNotRequired

  Edit.DataAddPrimaryKey "fncCodigo"
  Edit.DataIndexClustered "Codigo", "fncCodigo"

  If ( Not Edit.IsTable( "tesFuncoes" ) ) Then
     Edit.TableCreate "tesFuncoes"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final do procedimento que define a estrutura de dados tabela de >Funcoes
REM =========================================================================

REM =========================================================================
REM Procedimento que desenha o formulário de edição de "Funcoes"
REM -------------------------------------------------------------------------
Public Sub GetFuncoes

  TableTesFuncoes

  Edit.FormBegin "tesFuncoes", "Funções", 1, conOptionFuncoes, efValLocClient

  Edit.FormList "fncDescricao"
  Edit.FormUnit "fncDescricao"

    Edit.AddField "fncCodigo",    "Código", , , , efNext
    Edit.AddField "fncDescricao", "&Descrição", , , , , "A descrição da função deve ser preenchida"
      Edit.FieldHint "fncDescricao", "Descrição da função"
  
      Edit.FieldInternalLink "fncDescricao", conOptionFuncoes, "fncCodigo"

  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do procedimento que desenha o formulário de edição de "Funcoes"
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela de "Valores"
REM -------------------------------------------------------------------------
Public Sub TableTesValores

  Edit.DataBegin "tesValores"

  Edit.DataAddField "valCodigo", efDataTypeInt,       4, efRequired
  Edit.DataAddField "valIndice", efDataTypeInt,       4, efNotRequired
  Edit.DataAddField "valValor",  efDataTypeMoney,     8, efNotRequired
  Edit.DataAddField "valData",   efDataTypeDateTime, 10, efNotRequired

  Edit.DataAddPrimaryKey "valCodigo"
  Edit.DataIndexClustered "Codigo", "valCodigo"

  If ( Not Edit.IsTable( "tesValores" ) ) Then
     Edit.TableCreate "tesValores"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final do procedimento que define a estrutura de dados tabela de >Valores
REM =========================================================================

REM =========================================================================
REM Procedimento que desenha o formulário de edição de "Valores"
REM -------------------------------------------------------------------------
Public Sub GetValores

  TableTesValores
  Edit.FormBegin "tesValores", "Valores", 1, conOptionValores, efValLocClient

  Edit.FormList "valIndice,valValor,valData"
  Edit.FormUnit "valIndice,valValor,valData"

    Edit.AddField "valCodigo", "&Código", , , , efNext
    Edit.AddField "valIndice", "&Índice", efFldTypeLookup, , , , "Selecione um índice"
      Edit.FieldLookup "valIndice", "tesIndices", "indCodigo", "indDescricao"
      Edit.FieldInternalLink "valIndice", conOptionIndices, "indCodigo", True
      
    Edit.AddField "valValor",  "&Valor", , , , , "Digite um valor para o índice"
      Edit.FieldInternalLink "valValor", conOptionValores, "valCodigo"

    Edit.AddField "valData",   "&Data", , efValOptDate, , Strings.ZTILongDate(Now)
 
  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do procedimento que desenha o formulário de edição de "Valores"
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela de "Produtos"
REM -------------------------------------------------------------------------
Public Sub TableTesProdutos

  Edit.DataBegin "tesProdutos"

  Edit.DataAddField "proCodigo",     efDataTypeInt,       4, efRequired
  Edit.DataAddField "proNome",       efDataTypeVarchar,  30, efNotRequired
  Edit.DataAddField "proDescricao",  efDataTypeVarchar, 255, efNotRequired
  Edit.DataAddField "proValor",      efDataTypeMoney,     8, efNotRequired
  Edit.DataAddField "proIndice",     efDataTypeInt,       4, efNotRequired
  Edit.DataAddField "proFabricante", efDataTypeVarchar,  40, efNotRequired
  Edit.DataAddField "proDisponivel", efDataTypeBit,       1, efRequired

  Edit.DataAddPrimaryKey "proCodigo"
  Edit.DataIndexClustered "Codigo", "proCodigo"
  Edit.DataAddIndex "Nome", "proNome"

  If ( Not Edit.IsTable( "tesProdutos" ) ) Then
     Edit.TableCreate "tesProdutos"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final do procedimento que define a estrutura de dados tabela de >Produtos
REM =========================================================================

REM =========================================================================
REM Procedimento que desenha o formulário de edição de "Produtos"
REM -------------------------------------------------------------------------
Public Sub GetProdutos

  TableTesProdutos
  
  Edit.FormBegin "tesProdutos", "Produtos", 1, conOptionProdutos, efValLocClient

    Edit.FormList "proNome,proValor,proIndice,proFabricante,proDisponivel"
    Edit.FormUnit "proNome;proDescricao;proValor,proIndice,proDisponivel;proFabricante"
    Edit.FormFind "proNome,proDescricao,proValor,proIndice,proFabricante,proDisponivel"
  
    Edit.AddReport "proIndice,#Total(gc,1=Total)", "Relatório por Índice", 1
    Edit.AddReport "proFabricante,#Total(gc,1=Total)", "Relatório por Fabricante", 2
    Edit.AddReport "proDisponivel,#Total(gc,1=Total)", "Relatório por Disponibilidade", 3
  
    Edit.AddField "proCodigo",     "C&ódigo", , , , efNext
    Edit.AddField "proNome",       "&Nome", , , , , "O nome do produto deve ser preenchido"
      Edit.FieldHint "proNome",       "Nome do produto"
      Edit.FieldInternalLink "proNome", conOptionProdutos, "proCodigo"
      
    Edit.AddField "proDescricao",  "&Descrição", efFldTypeTextArea
      Edit.FieldHint "proDescricao",  "Descrição mais detalhada do produto"
      Edit.FieldShowSize "proDescricao", 80
      Edit.FieldListChars "proDescricao", 30
  
    Edit.AddField "proValor",      "&Valor"
      Edit.FieldHint "proValor",      "Valor do produto baseado no índice"
  
    Edit.AddField "proIndice",     "&Índice", efFldTypeLookup, , , , "Selecione um Índice"
      Edit.FieldLookup "proIndice",   "tesIndices", "indCodigo", "indDescricao"
      Edit.FieldInternalLink "proIndice", conOptionIndices, "indCodigo", True
      
    Edit.AddField "proFabricante", "&Fabricante"
      Edit.FieldHint "proFabricante", "Fabricante do produto"
  
    Edit.AddField "proDisponivel", "&Disponível", efFldTypeCheck, , , True
      Edit.FieldHint "proDisponivel", "O produto está disponível para entrega neste momento"
      Edit.FieldAddValue "proDisponivel", True, "Sim"
      Edit.FieldAddValue "proDisponivel", False, "Não"

  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do procedimento que desenha o formulário de edição de "Produtos"
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela de "Funcionarios"
REM -------------------------------------------------------------------------
Public Sub TableTesFuncionarios

  Edit.DataBegin "tesFuncionarios"

  Edit.DataAddField "funCodigo",         efDataTypeInt,       4, efRequired
  Edit.DataAddField "funNome",           efDataTypeVarchar,  50, efNotRequired
  Edit.DataAddField "funEndereco",       efDataTypeVarchar,  50, efNotRequired
  Edit.DataAddField "funBairro",         efDataTypeVarchar,  30, efNotRequired
  Edit.DataAddField "funCidade",         efDataTypeVarchar,  30, efNotRequired
  Edit.DataAddField "funUF",             efDataTypeVarchar,   2, efNotRequired
  Edit.DataAddField "funCEP",            efDataTypeVarchar,  10, efNotRequired
  Edit.DataAddField "funCPF",            efDataTypeVarchar,  11, efNotRequired
  Edit.DataAddField "funRG",             efDataTypeVarchar,  10, efNotRequired
  Edit.DataAddField "funOrgExpedidor",   efDataTypeVarchar,   8, efNotRequired
  Edit.DataAddField "funTelefone",       efDataTypeVarchar,  15, efNotRequired
  Edit.DataAddField "funFuncao",         efDataTypeInt,       4, efNotRequired
  Edit.DataAddField "funSexo",           efDataTypeVarChar,   1, efNotRequired
  Edit.DataAddField "funEstadoCivil",    efDataTypeVarChar,   1, efNotRequired
  Edit.DataAddField "funConjuge",        efDataTypeVarchar,  50, efNotRequired
  Edit.DataAddField "funNumDependentes", efDataTypeTinyInt,   2, efNotRequired
  Edit.DataAddField "funDataNascimento", efDataTypeDateTime, 10, efNotRequired
  Edit.DataAddField "funEmail",          efDataTypeVarchar,  80, efNotRequired
  Edit.DataAddField "funObservacoes",    efDataTypeVarchar, 255, efNotRequired

  Edit.DataAddPrimaryKey "funCodigo"
  Edit.DataIndexClustered "Codigo", "funCodigo"
  Edit.DataAddIndex "Nome", "funNome"
  Edit.DataAddIndex "RG", "funRG"
  Edit.DataAddIndex "CPF", "funCPF"
  Edit.DataAddIndex "Email", "funEmail"

  If ( Not Edit.IsTable( "tesFuncionarios" ) ) Then
     Edit.TableCreate "tesFuncionarios"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final do procedimento que define a estrutura de dados tabela de >Funcionarios
REM =========================================================================

REM =========================================================================
REM Procedimento que desenha o formulário de edição de "Funcionarios"
REM -------------------------------------------------------------------------
Public Sub GetFuncionarios

  TableTesFuncionarios
  
  Edit.FormBegin "tesFuncionarios", "Funcionários", 1, conOptionFuncionarios, efValLocClient

  Edit.FormList "funNome,funFuncao,funSexo,funTelefone,funEmail"
  Edit.FormUnit "funNome,funFuncao;funEndereco,funBairro;funCEP,funCidade,funUF;funTelefone,funRG,funOrgExpedidor;funCPF,funDataNascimento,funSexo;funEstadoCivil,funNumDependentes,funConjuge;funEmail;funObservacoes"
  Edit.FormFind "funNome,funBairro,funCidade,funUF,funCPF,funRG,funOrgExpedidor,funFuncao,funSexo,funEstadoCivil,funNumDependentes,funDataNascimento"

    Edit.AddReport "funCidade,#Total(gc,1=Total)", "Relatório por Cidade", 1
    Edit.AddReport "funUF,#Total(gc,1=Total)", "Relatório por UF", 2
    Edit.AddReport "funOrgExpedidor,#Total(gc,1=Total)", "Relatório por Orgão Expedidor da RG", 3
    Edit.AddReport "funEstadoCivil,#Total(gc,1=Total)", "Relatório por Estado Civil", 4
    Edit.AddReport "funSexo,#Total(gc,1=Total)", "Relatório por Sexo", 5

    Edit.AddField "funCodigo",        "&Código", , , , efNext
    Edit.AddField "funNome",          "&Nome", , , , , "O nome do funcionario deve ser preenchido"
      Edit.FieldInternalLink "funNome", conOptionFuncionarios, "funCodigo"
  
    Edit.AddField "funEndereco",      "&Endereço"
    Edit.AddField "funBairro",        "&Bairro"
    Edit.AddField "funCidade",        "C&idade"
      Edit.FieldShowSize "funCidade", 25
  
    Edit.AddField "funUF",            "&UF", efFldTypeUF, , , "DF"
    Edit.AddField "funCEP",           "CE&P", , efValOptCEP
    Edit.AddField "funCPF",           "C&PF", , efValOptCPF
      Edit.FieldMask "funCPF", "000.000.000-00", "0"
  
    Edit.AddField "funRG",            "&RG"
      Edit.FieldMask "funRG", "0 000 000"
  
    Edit.AddField "funOrgExpedidor",  "Org. Exp."
    Edit.AddField "funTelefone",      "&Telefone"
      Edit.FieldMask "funTelefone", "(000) 000-0000", " "
  
    Edit.AddField "funFuncao",        "&Função", efFldTypeLookup
      Edit.FieldLookup "funFuncao", "tesFuncoes", "fncCodigo", "fncDescricao"
      Edit.FieldInternalLink "funFuncao", conOptionFuncoes, "fncCodigo", True
  
    Edit.AddField "funSexo",          "&Sexo", efFldTypeRadio, , , "M"
      Edit.FieldAddValue "funSexo", "M", "Masculino"
      Edit.FieldAddValue "funSexo", "F", "Feminino"
      Edit.FieldRadioColumns "funSexo", 2
  
    Edit.AddField "funEstadoCivil",    "Esta&do Civil", efFldTypeSelect, , , "S"
      Edit.FieldAddValue "funEstadoCivil", "S", "Solteiro(a)"
      Edit.FieldAddValue "funEstadoCivil", "C", "Casado(a)"
      Edit.FieldAddValue "funEstadoCivil", "V", "Viúvo(a)"
      Edit.FieldAddValue "funEstadoCivil", "D", "Desquitado(a)"
  
    Edit.AddField "funConjuge",        "Con&juge"
      Edit.FieldDisableValue "funConjuge", "funEstadoCivil", "S", efCondDisable
      Edit.FieldShowSize "funConjuge", 30
  
    Edit.AddField "funNumDependentes", "Nº de de&pendentes"
    Edit.AddField "funDataNascimento", "Data de Nasc&imento", ,efValOptDate , , Strings.ZTILongDate(Now)
    Edit.AddField "funEmail",          "E-&mail", efFldTypeEmail, efValOptEmail
    Edit.AddField "funObservacoes",    "O&bservações", efFldTypeTextArea
      Edit.FieldShowSize "funObservacoes", 80
   
      Edit.AddOrder "funSexo,funNome"
      Edit.AddOrder "funNumDependentes,funEstadoCivil"
    
      Edit.AddHeader "Dados Comerciais", 2, 3
      Edit.AddHeader "Dados Pessoais", 5, 4
  
  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do procedimento que desenha o formulário de edição de "Funcionarios"
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela de "Clientes"
REM -------------------------------------------------------------------------
Public Sub TableTesClientes

  Edit.DataBegin "tesClientes"

  Edit.DataAddField "cliCodigo",         efDataTypeInt,       4, efRequired
  Edit.DataAddField "cliNome",           efDataTypeVarchar,  50, efNotRequired
  Edit.DataAddField "cliEndereco",       efDataTypeVarchar,  50, efNotRequired
  Edit.DataAddField "cliBairro",         efDataTypeVarchar,  30, efNotRequired
  Edit.DataAddField "cliCidade",         efDataTypeVarchar,  30, efNotRequired
  Edit.DataAddField "cliUF",             efDataTypeVarchar,   2, efNotRequired
  Edit.DataAddField "cliTelefone",       efDataTypeVarchar,  15, efNotRequired
  Edit.DataAddField "cliFax",            efDataTypeVarchar,  15, efNotRequired
  Edit.DataAddField "cliCEP",            efDataTypeVarchar,  10, efNotRequired
  Edit.DataAddField "cliFuncaoSocial",   efDataTypeVarchar,   1, efNotRequired
  Edit.DataAddField "cliRazaoSocial",    efDataTypeVarchar,  50, efNotRequired
  Edit.DataAddField "cliCGC",            efDataTypeVarchar,  20, efNotRequired
  Edit.DataAddField "cliInscricaoEst",   efDataTypeVarchar,  13, efNotRequired
  Edit.DataAddField "cliDataInscricao",  efDataTypeDateTime, 10, efNotRequired
  Edit.DataAddField "cliCPF",            efDataTypeVarchar,  11, efNotRequired
  Edit.DataAddField "cliRG",             efDataTypeVarchar,  20, efNotRequired
  Edit.DataAddField "cliOrgExpedidor",   efDataTypeVarchar,  10, efNotRequired
  Edit.DataAddField "cliSexo",           efDataTypeVarchar,   1, efNotRequired
  Edit.DataAddField "cliEstadoCivil",    efDataTypeVarchar,   1, efNotRequired
  Edit.DataAddField "cliConjuge",        efDataTypeVarchar,  50, efNotRequired
  Edit.DataAddField "cliNumDependentes", efDataTypeTinyInt,   2, efNotRequired
  Edit.DataAddField "cliDataNascimento", efDataTypeDateTime, 10, efNotRequired
  Edit.DataAddField "cliEmail",          efDataTypeVarchar,  80, efNotRequired
  Edit.DataAddField "cliObservacoes",    efDataTypeVarchar, 255, efNotRequired

  Edit.DataAddPrimaryKey "cliCodigo"
  Edit.DataIndexClustered "Codigo", "cliCodigo"
  Edit.DataAddIndex "Nome", "cliNome"
  Edit.DataAddIndex "CGC", "cliCGC"
  Edit.DataAddIndex "Email", "cliEmail"

  If ( Not Edit.IsTable( "tesClientes" ) ) Then
     Edit.TableCreate "tesClientes"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final do procedimento que define a estrutura de dados tabela de >Clientes
REM =========================================================================

REM =========================================================================
REM Procedimento que desenha o formulário de edição de "Clientes"
REM -------------------------------------------------------------------------
Public Sub GetClientes

  TableTesClientes
  Edit.FormBegin "tesClientes", "Clientes", 1, conOptionClientes, efValLocClient

    Edit.FormList "cliNome,CliFuncaoSocial,cliCPF,CliTelefone,cliEmail,CliCGC,CliRazaoSocial"
    Edit.FormUnit "cliNome,cliTelefone;cliEndereco,cliBairro;cliCidade,cliUF;CliEmail;CliObservacoes|CliFuncaoSocial;cliCPF,cliRG,cliOrgExpedidor;cliSexo,CliDataNascimento,cliEstadoCivil;cliConjuge,cliNumDependentes;CliRazaoSocial;CliCGC,CliInscricaoEst,CliDataInscricao"
    Edit.FormFind "cliNome,cliBairro,cliCidade,cliUF,cliCPF,cliRG,cliOrgExpedidor,cliSexo,cliEstadoCivil,cliNumDependentes,cliDataNascimento,cliRazaoSocial"
  
    Edit.FormTabs "Informações gerais|Informações comerciais"
  
    Edit.AddReport "cliCidade,#Total(gc,1=Total)", "Relatório por Cidade", 1
    Edit.AddReport "cliUF,#Total(gc,1=Total)", "Relatório por UF", 2
    Edit.AddReport "cliFuncaoSocial,#Total(gc,1=Total)", "Relatório por Função Social", 2
  
    Edit.AddField "cliCodigo",         "&Código", , , , efNext
    Edit.AddField "cliNome",           "&Nome", , , , , "O nome do cliente deve ser preenchido"
      Edit.FieldInternalLink "cliNome", conOptionClientes, "cliCodigo"

    Edit.AddField "cliEndereco",       "&Endereço"
    Edit.AddField "cliBairro",         "&Bairro"
    Edit.AddField "cliCidade",         "C&idade"
    Edit.AddField "cliUF",             "&UF", efFldTypeUF, , , "DF"
    Edit.AddField "cliTelefone",       "&Telefone"
      Edit.FieldMask "cliTelefone", "(000) 000-0000", " "
  
    Edit.AddField "cliFax",            "F&ax"
    Edit.AddField "cliEmail",          "E-&mail", efFldTypeEmail, efValOptEmail
    Edit.AddField "cliObservacoes",    "O&bservações", efFldTypeTextArea
      Edit.FieldShowSize "cliObservacoes", 80
  
    Edit.AddField "cliFuncaoSocial",   "F&unção Social", efFldTypeSelect, , , "F"
      Edit.FieldAddValue "cliFuncaoSocial", "F", "Pessoa Fisica"
      Edit.FieldAddValue "cliFuncaoSocial", "J", "Pessoa Juridica"
  
    Edit.AddField "cliCPF",            "C&PF", , efValOptCPF
      Edit.FieldMask "cliCPF", "000.000.000-00", "0"
      Edit.FieldDisableValue "cliCPF",            "cliFuncaoSocial", "J", efCondDisable
  
    Edit.AddField "cliRG",             "&RG"
      Edit.FieldMask "cliRG", "0 000 000"
      Edit.FieldDisableValue "cliRG",             "cliFuncaoSocial", "J", efCondDisable
  
    Edit.AddField "cliOrgExpedidor",   "&Órgão Expedidor"
      Edit.FieldDisableValue "cliOrgExpedidor",   "cliFuncaoSocial", "J", efCondDisable
  
    Edit.AddField "cliSexo",           "&Sexo", efFldTypeRadio, , , "M"
      Edit.FieldAddValue "cliSexo", "M", "Masculino"
      Edit.FieldAddValue "cliSexo", "F", "Feminino"
      Edit.FieldRadioColumns "cliSexo", "2"
      Edit.FieldDisableValue "cliSexo",           "cliFuncaoSocial", "J", efCondDisable
  
    Edit.AddField "cliEstadoCivil",    "Esta&do Civil", efFldTypeSelect, , , "S"
      Edit.FieldAddValue "cliEstadoCivil", "S", "Solteiro(a)"
      Edit.FieldAddValue "cliEstadoCivil", "C", "Casado(a)"
      Edit.FieldAddValue "cliEstadoCivil", "V", "Viúvo(a)"
      Edit.FieldAddValue "cliEstadoCivil", "D", "Desquitado(a)"
      Edit.FieldDisableValue "cliEstadoCivil",    "cliFuncaoSocial", "J", efCondDisable
  
    Edit.AddField "cliConjuge",        "Con&juge"
      Edit.FieldDisableValue "cliConjuge", "cliEstadoCivil", "S", efCondDisable
      Edit.FieldDisableValue "cliConjuge",        "cliFuncaoSocial", "J", efCondDisable
  
    Edit.AddField "cliNumDependentes", "Nº de de&pendentes"
      Edit.FieldDisableValue "cliNumDependentes", "cliFuncaoSocial", "J", efCondDisable
  
    Edit.AddField "cliDataNascimento", "Data de Nasc&imento", , efValOptDate, , Strings.ZTILongDate(Now)
      Edit.FieldDisableValue "cliDataNascimento", "cliFuncaoSocial", "J", efCondDisable
  
    Edit.AddField "cliRazaoSocial",    "Razão Socia&l"
      Edit.FieldDisableValue "cliRazaoSocial",   "cliFuncaoSocial", "F", efCondDisable
    Edit.AddField "cliCGC",            "C&GC", , efValOptCGC
      Edit.FieldMask "cliCGC", "00.000.000/0000-00", "0"
      Edit.FieldDisableValue "cliCGC",           "cliFuncaoSocial", "F", efCondDisable
  
    Edit.AddField "cliInscricaoEst",   "Inscrição estadual"
      Edit.FieldDisableValue "cliInscricaoEst",  "cliFuncaoSocial", "F", efCondDisable
  
    Edit.AddField "cliDataInscricao",  "Data de Inscrição", , efValOptDate, , Strings.ZTILongDate(Now)
      Edit.FieldDisableValue "cliDataInscricao", "cliFuncaoSocial", "F", efCondDisable
    
      Edit.AddOrder "cliSexo,cliNome"
      Edit.AddOrder "cliNumDependentes,cliEstadoCivil"
    
      Edit.AddHeader "Dados Pessoais", 3, 3
      Edit.AddHeader "Dados Comerciais", 6, 2
    
  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do procedimento que desenha o formulário de edição de "Clientes
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela de "Fornecedores"
REM -------------------------------------------------------------------------
Public Sub TableTesFornecedores

  Edit.DataBegin "tesFornecedores"

  Edit.DataAddField "forCodigo",        efDataTypeInt,       4, efRequired
  Edit.DataAddField "forNome",          efDataTypeVarchar,  50, efNotRequired
  Edit.DataAddField "forEndereco",      efDataTypeVarchar,  50, efNotRequired
  Edit.DataAddField "forBairro",        efDataTypeVarchar,  30, efNotRequired
  Edit.DataAddField "forCidade",        efDataTypeVarchar,  30, efNotRequired
  Edit.DataAddField "forUF",            efDataTypeVarchar,   2, efNotRequired
  Edit.DataAddField "forTelefone",      efDataTypeVarchar,  15, efNotRequired
  Edit.DataAddField "forCEP",           efDataTypeVarchar,  10, efNotRequired
  Edit.DataAddField "forRazaoSocial",   efDataTypeVarchar,  50, efNotRequired
  Edit.DataAddField "forCGC",           efDataTypeVarchar,  20, efNotRequired
  Edit.DataAddField "forInscricaoEst",  efDataTypeVarchar,  13, efNotRequired
  Edit.DataAddField "forDataInscricao", efDataTypeDateTime, 10, efNotRequired
  Edit.DataAddField "forFax",           efDataTypeVarchar,  15, efNotRequired
  Edit.DataAddField "forEmail",         efDataTypeVarchar,  80, efNotRequired
  Edit.DataAddField "forObservacoes",   efDataTypeVarchar, 255, efNotRequired

  Edit.DataAddPrimaryKey "forCodigo"
  Edit.DataIndexClustered "Codigo", "forCodigo"
  Edit.DataAddIndex "Nome", "forNome"
  Edit.DataAddIndex "RazaoSocial", "forRazaoSocial"
  Edit.DataAddIndex "CGC", "forCGC"
  Edit.DataAddIndex "Email", "forEmail"

  If ( Not Edit.IsTable( "tesFornecedores" ) ) Then
     Edit.TableCreate "tesFornecedores"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final do procedimento que define a estrutura de dados tabela de >Fornecedores
REM =========================================================================

REM =========================================================================
REM Procedimento que desenha o formulário de edição de "Fornecedores"
REM -------------------------------------------------------------------------
Public Sub GetFornecedores

  TableTesFornecedores
  Edit.FormBegin "tesFornecedores", "Fornecedores", 1, conOptionFornecedores, efValLocClient

    Edit.FormList "forNome,forTelefone,forFax,forEmail"
    Edit.FormUnit "forNome,forInscricaoEst;forTelefone,forFax,forDataInscricao;forRazaoSocial,forCGC;forEndereco,forCEP;forBairro,forCidade,forUF;forEmail;forObservacoes"
    Edit.FormFind "forNome,forEndereco,forBairro,forCidade,forUF,forTelefone,forCEP,forRazaoSocial,forCGC,forInscricaoEst,forDataInscricao,forFax,forEmail,forObservacoes"

    Edit.AddReport "forCidade,#Total(gc,1=Total)", "Relatório por Cidade", 1
    Edit.AddReport "forUF,#Total(gc,1=Total)", "Relatório por UF", 2
  
    Edit.AddField "forcodigo",        "&Código", , , , efNext
    Edit.AddField "forNome",          "&Nome", , , , , "O nome do Fornecedor deve ser preenchido"
      Edit.FieldInternalLink "forNome", conOptionFornecedores, "forCodigo"

    Edit.AddField "forEndereco",      "&Endereço"
    Edit.AddField "forBairro",        "&Bairro"
    Edit.AddField "forCidade",        "C&idade"
    Edit.AddField "forUF",            "&UF",efFldTypeUF, , , "DF"
    Edit.AddField "forCep",           "Ce&p", , efValOptCEP
      Edit.FieldMask "forCep", "00.000-000", " "
  
    Edit.AddField "forTelefone",      "&Telefone"
      Edit.FieldMask "forTelefone", "(000) 000-0000", " "
  
    Edit.AddField "forRazaoSocial",   "&Razão Social"
    Edit.AddField "forCGC",           "C&GC", , efValOptCGC
      Edit.FieldMask "forCGC", "00.000.000/0000-00", "0"
  
    Edit.AddField "forInscricaoEst",  "In&scrição Estadual"
    Edit.AddField "forDataInscricao", "&Data de Inscrição", , , , Strings.ZTILongDate(Now)
    Edit.AddField "forFax",           "&Fax"
      Edit.FieldMask "forFax", "(000) 000-0000", " "
  
    Edit.AddField "forEmail",         "E-&mail", efFldTypeEmail, efValOptEmail
    Edit.AddField "forObservacoes",   "O&bservações", efFldTypeTextArea
      Edit.FieldShowSize "forObservacoes", 80
  
    Edit.AddOrder "forCodigo,forNome"
  
  REM  EditAddHeader "Dados Comerciais", 2, 3
  REM  EditAddHeader "Dados Pessoais", 5, 2
  
  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do procedimento que desenha o formulário de edição de "Fornecedores"
REM =========================================================================

REM =========================================================================
REM  Cria URL Para o Edit Form
REM -------------------------------------------------------------------------
Public Sub EditCreateURL(ByVal strURLShow, ByVal conWhatStr)

  URL.BeginURL Initializer.ScriptName
  URL.Add efQueryStrAction, efQSActionEditor
  URL.Add efQueryStrWhat, conWhatStr
  URL.Show strURLShow, efQueryStrEditableStr
  URL.EndURL

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditCreateURL
REM =========================================================================


REM -------------------------------------------------------------------------
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub MainBody

  FormatTable

  Edit.ConnectionString = Session("ConnectionString")

  Edit.QueryString

  Select Case Edit.parWhat
    Case conOptionIndices
      GetIndices

    Case conOptionValores
      GetValores

    Case conOptionProdutos
      GetProdutos

    Case conOptionFuncionarios
      GetFuncionarios

    Case conOptionFuncoes
      GetFuncoes

    Case conOptionClientes
      GetClientes

    Case conOptionFornecedores
      GetFornecedores

    Case conOptionCreateAll
      GetIndices
      Edit.TableCreate "tesIndices"

      GetValores
      Edit.TableCreate "tesValores"

      GetProdutos
      Edit.TableCreate "tesProdutos"

      GetFuncionarios
      Edit.TableCreate "tesFuncionarios"

      GetFuncoes
      Edit.TableCreate "tesFuncoes"

      GetClientes
      Edit.TableCreate "tesClientes"

      GetFornecedores
      Edit.TableCreate "tesFornecedores"

    Case Else
      ShowFrames

  End Select

  Edit.RedirectActions

  Default.BodyText  = "Brown"
  Default.BodyLink  = "Brown"
  Default.BodyVLink = "#0B8D94"
  Default.BodyALink = "#0B8D94"

  Default.HTMLBegin
  Default.HeadAll "Sistema de Entregas"
  Default.BodyBegin

  Show.HTMLCR "<H3 ALIGN=CENTER>Sistema de Entregas</H3>"
  Show.HTMLCR "<HR>"
  Show.Center

  If Edit.IsMyAction Then
    ShowFirstPage

  End If

  Default.PageFooterDefault
  Default.BodyEnd
  Default.HTMLEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================

REM -------------------------------------------------------------------------
REM Fim do Entregas.asp
REM =========================================================================
%>
