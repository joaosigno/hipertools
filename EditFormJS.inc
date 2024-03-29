<%
// -------------------------------------------------------------------------
//  /HiperTools/EditForm.inc
// -------------------------------------------------------------------------
//  Descricao  : Constantes do HiperTools30.EditForm
//  Cria��o    : 19/02/2000 11:21AM
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

// Modos dos formul�rios de edi��o
// --------------------------
// Somente Leitura
var efFormReadOnly = 0;
// Possibilitar altera��o/inclus�o
var efFormWritable = 1;

// QueryString Parameters
// --------------------------
// A��es feitas pelo EditForm
var efQueryStrAction        = "EA";
// Tabela que est� sendo editada
var efQueryStrWhat          = "EW";
// Indica a editabilidade da tabela
var efQueryStrEditable      = "EE";
// Simplifica a coloca��o da constante acima
var efQueryStrEditableStr   = "&EE=1";
// Posi��o absoluta de um registro a ser encontrado
var efQueryStrFind          = "EF";
// Filtro de registros(WHERE) aplicado � tabela
var efQueryStrFilter        = "ER";
// Relat�rio a ser apresentado
var efQueryStrList          = "EL";
// Valores default de campos na inclus�o
var efQueryStrDefaults      = "ED";
// Tab da edi��o
var efQueryStrTab           = "ET";
// Campo de Ordena��o (pelo �ndice)
var efQueryStrOrderField    = "EOF";
// Ordena��o decrescente do campo de ordena��o
var efQueryStrOrderDescStr = "&EOD=1";

// QueryString Action Values
// --------------------------
// Formul�rio
var efQSActionEditor = "h01";
// Listagem
var efQSActionList = "h02";
// Relat�rio
var efQSActionSummary = "h03";
// Localiza��o simples
var efQSActionCommonFind = "h04";
// Localiza��o avan�ada
var efQSActionAdvancedFind = "h05";
// Inclus�o
var efQSActionAdd = "h06";
// C�pia
var efQSActionCopy = "h07";
// Edi��o
var efQSActionEdit = "h08";
// Exclus�o
var efQSActionDelete = "h09";
// Grava��o
var efQSActionSave = "h10";
// Grava��o e inclus�o
var efQSActionSaveAdd = "h11";
// Grava��o e C�pia
var efQSActionSaveCopy = "h12";
// Apresenta��o do Formul�rio
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
// Marca��o
var efFldTypeCheck = 3;
// Texto longo
var efFldTypeTextArea = 4;
// Bot�es de sele��o
var efFldTypeRadio = 5;
// Caixa de sele��o
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
// Data de atualiza��o
var efFldTypeAtualization = 13;
// Cor
var efFldTypeColor = 14;
// Pa�s
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
// Compara��o de datas(Ex.:A data do primeiro campo n�o pode ser maior que a do segundo)
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
// N�o requerido
var efNotRequired = true;

// Relational Integrity;
// --------------------------
// Exclus�o em cascata;
var efRelIntegrCascade = 0;
// Exclus�o de registro pai apenas se n�o tiver filhos
var efRelIntegrRestricted = 1;
// Exclus�o com a retirada do relacionamento do pai no filho
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

// Autonumera��o/contador/identity
// --------------------------
var efNext        = "proximo ";

// Alphabetic index types
// --------------------------
var efAlphaIndexNumeric = 1;
var efAlphaIndexAlphaNumeric = 2;
var efAlphaIndexAlphabetic = 3;

%>