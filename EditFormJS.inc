<%
// -------------------------------------------------------------------------
//  /HiperTools/EditForm.inc
// -------------------------------------------------------------------------
//  Descricao  : Constantes do HiperTools30.EditForm
//  Criação    : 19/02/2000 11:21AM
//  Local      : Brasilia/DF
//  Elaborado  : Ruben Zevallos <ruben@zevallos.com.br>
//               Fernando Aquino <fernando@zevallos.com.br>
//               Kelson Medeiros <kelson@zevallos.com.br>
//               Ueslei Silva <ueslei@zevallos.com.br>
//               Eduardo Silva <ueslei@zevallos.com.br>
//               Ridai Govinda <ridai@zevallos.com.br>
//  Copyright  : 2000 by Zevallos(r) Tecnologia em Informacao
// -------------------------------------------------------------------------

// Database Software Type (ConnectionType)
// ---------------------------------------
var efConnDBDriverSQL = 1;
var efConnDBDriverDBase = 2;
var efConnDBDriverAccess = 3;
var efConnDBDriverExcel = 4;
var efConnDBDriverFoxPro = 5;
var efConnDBDriverText = 6;
var efConnDBDriverParadox = 7;
var efConnDBProviderJet35 = 8;
var efConnDBProviderJet40 = 9;
var efConnDBProviderOracle = 10;
var efConnDBProviderSQL = 11;
var efConnDBProviderODBC = 12;
var efConnDBProviderDTS = 13;
var efConnDBProviderSQLDTS = 14;
var efConnDBProviderSimple = 15;
var efConnDBProviderRemote = 16;

// Modos dos formulários de edição
// --------------------------
// Somente Leitura
var efFormReadOnly = 0;
// Possibilitar alteração/inclusão
var efFormWritable = 1;

// QueryString Parameters
// --------------------------
// Ações feitas pelo EditForm
var efQueryStrAction        = "EA";
// Tabela que está sendo editada
var efQueryStrWhat          = "EW";
// Indica a editabilidade da tabela
var efQueryStrEditable      = "EE";
// Simplifica a colocação da constante acima
var efQueryStrEditableStr   = "&EE=1";
// Posição absoluta de um registro a ser encontrado
var efQueryStrFind          = "EF";
// Filtro de registros(WHERE) aplicado à tabela
var efQueryStrFilter        = "ER";
// Relatório a ser apresentado
var efQueryStrList          = "EL";
// Valores default de campos na inclusão
var efQueryStrDefaults      = "ED";
// Tab da edição
var efQueryStrTab           = "ET";
// Campo de Ordenação (pelo índice)
var efQueryStrOrderField    = "EOF";
// Ordenação decrescente do campo de ordenação
var efQueryStrOrderDescStr = "&EOD=1";

// QueryString Action Values
// --------------------------
// Formulário
var efQSActionEditor = "h01";
// Listagem
var efQSActionList = "h02";
// Relatório
var efQSActionSummary = "h03";
// Localização simples
var efQSActionCommonFind = "h04";
// Localização avançada
var efQSActionAdvancedFind = "h05";
// Inclusão
var efQSActionAdd = "h06";
// Cópia
var efQSActionCopy = "h07";
// Edição
var efQSActionEdit = "h08";
// Exclusão
var efQSActionDelete = "h09";
// Gravação
var efQSActionSave = "h10";
// Gravação e inclusão
var efQSActionSaveAdd = "h11";
// Gravação e Cópia
var efQSActionSaveCopy = "h12";
// Apresentação do Formulário
var efQSActionScreen = "h14";

// Data Types;
// --------------------------;
// Ponto flutuante
var efDataTypeFloat          = "float";
// Real
var efDataTypeReal           = "real";
// String de tamanho variante
var efDataTypeVarChar        = "varchar";
// String de tamanho fixo
var efDataTypeChar           = "char";
// Texto longo
var efDataTypeText           = "text";
// Inteiro
var efDataTypeInt            = "int";
// Data e hora
var efDataTypeDateTime       = "datetime";
// Money
var efDataTypeMoney          = "money";
// Inteiro de 1 byte
var efDataTypeTinyInt        = "tinyint";
// Interio de 2 bytes
var efDataTypeSmallInt       = "smallint";
// Booleana
var efDataTypeBit            = "bit";

// Field Types
// --------------------------
// Texto
var efFldTypeText = 0;
// Estado
var efFldTypeUF = 1;
// Busca os valores apresentados em outra tabela
var efFldTypeLookup = 2;
// Marcação
var efFldTypeCheck = 3;
// Texto longo
var efFldTypeTextArea = 4;
// Botões de seleção
var efFldTypeRadio = 5;
// Caixa de seleção
var efFldTypeSelect = 6;
// Senha
var efFldTypePassword = 7;
// HTTP
var efFldTypeHTTP = 8;
// E-Mail
var efFldTypeEMail = 9;
// Data com campos separados
var efFldTypeSeparateDate = 10;
// Imagem
var efFldTypeImage = 11;
// Arquivo
var efFldTypeFile = 12;
// Data de atualização
var efFldTypeAtualization = 13;
// Cor
var efFldTypeColor = 14;
// País
var efFldTypeCountry = 15;
// FTP
var efFldTypeFTP = 16;
// Gopher
var efFldTypeGopher = 17;

// Validation Location
// --------------------------
// Cliente (Browser)
var efValLocClient = true;
// Servidor (HiperTools)
var efValLocServer = false;

// Validation Options
// --------------------------
// Nenhuma
var efValOptNone = 0;
// CGC
var efValOptCGC = 1;
// CPF
var efValOptCPF = 2;
// Data
var efValOptDate = 3;
// Data com campos separados
var efValOptSepDate = 4;
// Data maior que hoje
var efValOptDateMToday = 5;
// Data com campos separados maior que hoje
var efValOptSepDateMToday = 6;
// Hora
var efValOptTime = 7;
// E-Mail
var efValOptEmail = 8;
// Comparação de datas(Ex.:A data do primeiro campo não pode ser maior que a do segundo)
var efValOptCompareDates = 9;
// CEP
var efValOptCEP = 10;
// CGC ou CPF
var efValOptCGCCPF = 11;
// CPF ou CGC
var efValOptCPFCGC = 12;

// Field Requirement
// --------------------------
// Requerido
var efRequired = false;
// Não requerido
var efNotRequired = true;

// Relational Integrity;
// --------------------------
// Exclusão em cascata;
var efRelIntegrCascade = 0;
// Exclusão de registro pai apenas se não tiver filhos
var efRelIntegrRestricted = 1;
// Exclusão com a retirada do relacionamento do pai no filho
var efRelIntegrNullifies = 2;

// Relation Type
// --------------------------
// 1 para 1
var efRelType1to1 = 0;
// 1 para muitos
var efRelType1toN = 1;

// Field Disable Condition
// --------------------------
// Desabilitante
var efCondDisable = true;
// Abilitante
var efCondEnable = false;

// Operator types
// --------------------------
//OR operator
var efBooleanOperatorOR = 1;
//AND operator
var efBooleanOperatorAND = 2;

//     Character Case
// --------------------------
var efCharCaseNormal = 0;
var efCharCaseUpper = 1;
var efCharCaseLower = 2;

// Autonumeração/contador/identity
// --------------------------
var efNext        = "proximo ";

// Alphabetic index types
// --------------------------
var efAlphaIndexNumeric = 1;
var efAlphaIndexAlphaNumeric = 2;
var efAlphaIndexAlphabetic = 3;

%>