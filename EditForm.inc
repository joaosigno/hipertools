<%
REM -------------------------------------------------------------------------
REM  /HiperTools/EditForm.inc
REM -------------------------------------------------------------------------
REM  Descricao  : Constantes do HiperTools30.EditForm
REM  Cria��o    : 19/02/2000 11:21AM
REM  Local      : Brasilia/DF
REM  Elaborado  : Ruben Zevallos <ruben@zevallos.com.br>
REM               Fernando Aquino <fernando@zevallos.com.br>
REM               Kelson Medeiros <kelson@zevallos.com.br>
REM               Ueslei Silva <ueslei@zevallos.com.br>
REM               Eduardo Silva <ueslei@zevallos.com.br>
REM               Ridai Govinda <ridai@zevallos.com.br>
REM  Copyright  : 2000 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

REM Database Software Type (ConnectionType)
REM ---------------------------------------
Const efConnDBDriverSQL = 1
Const efConnDBDriverDBase = 2
Const efConnDBDriverAccess = 3
Const efConnDBDriverExcel = 4
Const efConnDBDriverFoxPro = 5
Const efConnDBDriverText = 6
Const efConnDBDriverParadox = 7
Const efConnDBProviderJet35 = 8
Const efConnDBProviderJet40 = 9
Const efConnDBProviderOracle = 10
Const efConnDBProviderSQL = 11
Const efConnDBProviderODBC = 12
Const efConnDBProviderDTS = 13
Const efConnDBProviderSQLDTS = 14
Const efConnDBProviderSimple = 15
Const efConnDBProviderRemote = 16
Const efConnDBDriverOracle = 17

REM Modos dos formul�rios de edi��o
REM --------------------------
REM Somente Leitura
Const efFormReadOnly = 0
REM Possibilitar altera��o/inclus�o
Const efFormWritable = 1


REM QueryString Parameters
REM --------------------------
REM A��es feitas pelo EditForm
Const efQueryStrAction        = "EA"
REM Tabela que est� sendo editada
Const efQueryStrWhat          = "EW"
REM Indica a editabilidade da tabela
Const efQueryStrEditable      = "EE"
REM Simplifica a coloca��o da constante acima
Const efQueryStrEditableStr   = "&EE=1"
REM Posi��o absoluta de um registro a ser encontrado
Const efQueryStrFind          = "EF"
REM Filtro de registros(WHERE) aplicado � tabela
Const efQueryStrFilter        = "ER"
REM Relat�rio a ser apresentado
Const efQueryStrList          = "EL"
REM Valores default de campos na inclus�o
Const efQueryStrDefaults      = "ED"
REM Tab da edi��o
Const efQueryStrTab           = "ET"
REM Campo de Ordena��o (pelo �ndice)
Const efQueryStrOrderField    = "EOF"
REM Ordena��o decrescente do campo de ordena��o
Const efQueryStrOrderDescStr = "&EOD=1"

REM QueryString Action Values
REM --------------------------
REM Formul�rio
Const efQSActionEditor = "h01"
REM Listagem
Const efQSActionList = "h02"
REM Relat�rio
Const efQSActionSummary = "h03"
REM Localiza��o simples
Const efQSActionCommonFind = "h04"
REM Localiza��o avan�ada
Const efQSActionAdvancedFind = "h05"
REM Inclus�o
Const efQSActionAdd = "h06"
REM C�pia
Const efQSActionCopy = "h07"
REM Edi��o
Const efQSActionEdit = "h08"
REM Exclus�o
Const efQSActionDelete = "h09"
REM Grava��o
Const efQSActionSave = "h10"
REM Grava��o e inclus�o
Const efQSActionSaveAdd = "h11"
REM Grava��o e C�pia
Const efQSActionSaveCopy = "h12"
REM Apresenta��o do Formul�rio
Const efQSActionScreen = "h14"

REM Data Types
REM --------------------------
REM Ponto flutuante
Const efDataTypeFloat          = "float"
REM Real
Const efDataTypeReal           = "real"
REM String de tamanho variante
Const efDataTypeVarChar        = "varchar"
REM String de tamanho fixo
Const efDataTypeChar           = "char"
REM Texto longo
Const efDataTypeText           = "text"
REM Inteiro
Const efDataTypeInt            = "int"
REM Data e hora
Const efDataTypeDateTime       = "datetime"
REM Money
Const efDataTypeMoney          = "money"
REM Inteiro de 1 byte
Const efDataTypeTinyInt        = "tinyint"
REM Interio de 2 bytes
Const efDataTypeSmallInt       = "smallint"
REM Booleana
Const efDataTypeBit            = "bit"

REM Field Types
REM --------------------------
REM Texto
Const efFldTypeText = 0
REM Estado
Const efFldTypeUF = 1
REM Busca os valores apresentados em outra tabela
Const efFldTypeLookup = 2
REM Marca��o
Const efFldTypeCheck = 3
REM Texto longo
Const efFldTypeTextArea = 4
REM Bot�es de sele��o
Const efFldTypeRadio = 5
REM Caixa de sele��o
Const efFldTypeSelect = 6
REM Senha
Const efFldTypePassword = 7
REM HTTP
Const efFldTypeHTTP = 8
REM E-Mail
Const efFldTypeEMail = 9
REM Data com campos separados
Const efFldTypeSeparateDate = 10
REM Imagem
Const efFldTypeImage = 11
REM Arquivo
Const efFldTypeFile = 12
REM Data de atualiza��o
Const efFldTypeAtualization = 13
REM Cor
Const efFldTypeColor = 14
REM Pa�s
Const efFldTypeCountry = 15
REM FTP
Const efFldTypeFTP = 16
REM Gopher
Const efFldTypeGopher = 17

REM Validation Location
REM --------------------------
REM Cliente (Browser)
Const efValLocClient = True
REM Servidor (HiperTools)
Const efValLocServer = False

REM Validation Options
REM --------------------------
REM Nenhuma
Const efValOptNone = 0
REM CGC
Const efValOptCGC = 1
REM CPF
Const efValOptCPF = 2
REM Data
Const efValOptDate = 3
REM Data com campos separados
Const efValOptSepDate = 4
REM Data maior que hoje
Const efValOptDateMToday = 5
REM Data com campos separados maior que hoje
Const efValOptSepDateMToday = 6
REM Hora
Const efValOptTime = 7
REM E-Mail
Const efValOptEmail = 8
REM Compara��o de datas(Ex.:A data do primeiro campo n�o pode ser maior que a do segundo)
Const efValOptCompareDates = 9
REM CEP
Const efValOptCEP = 10
REM CGC ou CPF
Const efValOptCGCCPF = 11
REM CPF ou CGC
Const efValOptCPFCGC = 12

REM Field Requirement
REM --------------------------
REM Requerido
Const efRequired = False
REM N�o requerido
Const efNotRequired = True

REM Relational Integrity
REM --------------------------
REM Exclus�o em cascata
Const efRelIntegrCascade = 0
REM Exclus�o de registro pai apenas se n�o tiver filhos
Const efRelIntegrRestricted = 1
REM Exclus�o com a retirada do relacionamento do pai no filho
Const efRelIntegrNullifies = 2

REM Relation Type
REM --------------------------
REM 1 para 1
Const efRelType1to1 = 0
REM 1 para muitos
Const efRelType1toN = 1

REM Field Disable Condition
REM --------------------------
REM Desabilitante
Const efCondDisable = True
REM Abilitante
Const efCondEnable = False

REM Operator types
REM --------------------------
Const efBooleanOperatorOR = 1
Const efBooleanOperatorAND = 2

REM     Character Case
REM --------------------------
Const efCharCaseNormal = 0
Const efCharCaseUpper = 1
Const efCharCaseLower = 2

REM Autonumera��o/contador/identity
REM --------------------------
Const efNext        = "proximo "

REM Alphabetic index types
REM --------------------------
Const efAlphaIndexNumeric = 1
Const efAlphaIndexAlphaNumeric = 2
Const efAlphaIndexAlphabetic = 3
%>