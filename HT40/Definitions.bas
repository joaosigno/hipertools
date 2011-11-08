Attribute VB_Name = "Definitions"
'-------------------------------------------------------------------------
' Definitions.bas
'-------------------------------------------------------------------------
' Descricao   : Definições de constantes e procedimento públicos
' Criaçao     : 16/01/2000 18:55
' Local       : Brasilia/DF
' Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
'               Ridai Govinda Pombo <ridai@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 97-2000 by Zevallos(r) Tecnologia em Informacao
'-------------------------------------------------------------------------

Option Explicit

'Criado (Ridai Govinda Pombo)
Private Const htEvalExpiresAt = #12/31/2003#
Private Const htEvalExpiresString = "O período de teste desta versão do " & _
                                    "HiperTools 3.0 chegou ao fim no dia 31/12/2003, " & _
                                    "para obter uma nova versão, por favor, contate a Zevallos Tecnologia em Informação " & _
                                    "no e-mail info@hipertools.com.br ou pelo telefone (61) 328-3575"
'---------------------------

Rem ======================================================
Rem Inicio das Constantes do Connection
Rem ------------------------------------------------------
  
Public Const conConnSQL = 1
Public Const conConnDBase = 2
Public Const conConnAccess = 3
Public Const conConnExcel = 4
Public Const conConnFoxPro = 5
Public Const conConnText = 6
Public Const conConnParadox = 7

Rem ------------------------------------------------------
Rem Final das Constantes do Connection
Rem ======================================================

Rem ======================================================
Rem Inicio das Constantes do Registry
Rem ------------------------------------------------------

Rem Entradas da Registry
Rem -----------------------------
Public Const rgHKEYClassesRoot = &H80000000
Public Const rgHKEYCurrentUser = &H80000001
Public Const rgHKEYLocalMachine = &H80000002
Public Const rgHKEYUsers = &H80000003
Public Const rgHKEYPerformanceData = &H80000004
Public Const rgHKEYCurrentConfig = &H80000005
Public Const rgHKEYDynData = &H80000006

Rem Tipos de Valores do Registry
Rem -----------------------------
Public Const rgNone = 0                        ' Sem tipo de dado
Public Const rgSz = 1                          ' String tipo Unicode terminada com nulo
Public Const rgExpandSz = 2                    ' String tipo Unicode terminada com nulo com
                                        ' referências à variáveis de ambiente: %windir% p/ exemplo
Public Const rgBinary = 3                      ' Dado binário (livre)
Public Const rgDword = 4                       ' Número 32-bit (Long Integer)
Public Const rgDwordLittleEndian = 4           ' Número 32-bit  (o mesmo que rgDword)
Public Const rgDwordBigEndian = 5              ' Número 32-bit
Public Const rgLink = 6                        ' Link simbólico (unicode)
Public Const rgMultiSz = 7                     ' Strings tipo Unicode múltiplas
Public Const rgResourceList = 8                ' Resource list in the resource map
Public Const rgFullResourceDescriptor = 9      ' Resource list in the hardware description
Public Const rgResourceRequirementsList = 10   ' Resource requirements list
Rem ------------------------------------------------------
Rem Final das Constantes do Registry
Rem ======================================================


Rem ======================================================
Rem Inicio das Constantes do EditForm
Rem ------------------------------------------------------
  
Rem Criado (Ridai Govinda)
Rem     Lock operations (ReadFieldLock parameter)
  Public Const efLockGetState = 1
  Public Const efLockClearLock = 2
  Public Const efLockDoLock = 3
  Public Const efLockForceClearLock = 4

Rem     Modos dos formulários de edição
  Public Const efFormReadOnly = 0
  Public Const efFormWritable = 1

Rem     Database Software Type
  Public Const efCnnStrDriverSQL = "{sql server}"
  Public Const efCnnStrDriverAccess = "{microsoft access driver (*.mdb)}"
  Public Const efCnnStrDriverdBase = "{microsoft dbase driver (*.dbf)}"
  Public Const efCnnStrDriverFoxPro = "{microsoft foxpro driver (*.dbf)}"
  Public Const efCnnStrDriverParadox = "{microsoft paradox driver (*.db )}"
  Public Const efCnnStrDriverText = "{microsoft text driver (*.txt;*.csv)}"
  Public Const efCnnStrDriverExcel = "{microsoft excel driver (*.xls)}"
  Public Const efCnnStrProviderJet35 = "microsoft.jet.oledb.3.51"
  Public Const efCnnStrProviderJet40 = "microsoft.jet.oledb.4.0"
  Public Const efCnnStrProviderOracle = "msdaora.1"
  Public Const efCnnStrProviderSQL = "sqloledb.1"
  Public Const efCnnStrProviderODBC = "msdasql.1"
  Public Const efCnnStrProviderDTS = "dtspackagedso.1"
  Public Const efCnnStrProviderSQLDTS = "dtsflatfile.1"
  Public Const efCnnStrProviderSimple = "msdaosp.1"
  Public Const efCnnStrProviderRemote = "ms remote.1"
  Public Const efCnnStrDriverOracle = "{microsoft odbc for oracle}"
  
  Public Const efConnDBDriverSQL = 1
  Public Const efConnDBDriverDBase = 2
  Public Const efConnDBDriverAccess = 3
  Public Const efConnDBDriverExcel = 4
  Public Const efConnDBDriverFoxPro = 5
  Public Const efConnDBDriverText = 6
  Public Const efConnDBDriverParadox = 7
  Public Const efConnDBProviderJet35 = 8
  Public Const efConnDBProviderJet40 = 9
  Public Const efConnDBProviderOracle = 10
  Public Const efConnDBProviderSQL = 11
  Public Const efConnDBProviderODBC = 12
  Public Const efConnDBProviderDTS = 13
  Public Const efConnDBProviderSQLDTS = 14
  Public Const efConnDBProviderSimple = 15
  Public Const efConnDBProviderRemote = 16
  Public Const efConnDBDriverOracle = 17
  
Rem     Operador p/ desabilitação de butões do Unit
  Public Const efBttnOperatorEqual = 1
  Public Const efBttnOperatorNotEqual = 2
  Public Const efBttnOperatorGreaterThan = 3
  Public Const efBttnOperatorLessThan = 4
  
'----------------------------
Rem     QueryString Parameters
  Public Const efQueryStrAction = "EA"
  Public Const efQueryStrWhat = "EW"
  Public Const efQueryStrMove = "EM"
  Public Const efQueryStrEditable = "EE"
  Public Const efQueryStrEditableStr = "&EE=1"
  Public Const efQueryStrOrderField = "EOF"
  Public Const efQueryStrOrderDesc = "EOD"
  Public Const efQueryStrOrderDescStr = "&EOD=1"
  Public Const efQueryStrFind = "EF"
  Public Const efQueryStrFilter = "ER"
  Public Const efQueryStrList = "EL"
  Public Const efQueryStrDefaults = "ED"
  Public Const efQueryStrTab = "ET"
  Public Const efQueryStrGridStr = "&EGO=1"
  Public Const efQueryStrGrid = "EGO"
  'Criado (Ridai Govinda)
  Public Const efQueryStrFileOpt = "EFO"
  Public Const efQueryStrFileName = "EFN"
  Public Const efQueryStrFileFolder = "EFF"
  Public Const efQueryStrFileField = "EFL"
  Public Const efQueryStrLookUpField = "ELF"

Rem     QueryString File Option Values
  Public Const efFileOptChoose = 1
  Public Const efFileOptSave = 2

Rem     QueryString Action Values
  Public Const efQSActionEditor = "h01"
  Public Const efQSActionList = "h02"
  Public Const efQSActionSummary = "h03"
  Public Const efQSActionCommonFind = "h04"
  Public Const efQSActionAdvancedFind = "h05"
  Public Const efQSActionAdd = "h06"
  Public Const efQSActionCopy = "h07"
  Public Const efQSActionEdit = "h08"
  Public Const efQSActionDelete = "h09"
  Public Const efQSActionSave = "h10"
  Public Const efQSActionSaveAdd = "h11"
  Public Const efQSActionSaveCopy = "h12"

  Public Const efQSActionGrid = "h13"
  Public Const efQSActionScreen = "h14"
  Public Const efQSActionExeEdition = "h15"
  Public Const efQSActionExeFind = "h16"
  'Criado (Ridai Govinda)
  Public Const efQSActionGetFile = "h17"
  Public Const efQSActionLookupList = "h18"
  '---------------------

Rem     Data Types
  Public Const efDataTypeFloat = "float"
  Public Const efDataTypeReal = "real"
  Public Const efDataTypeVarChar = "varchar"
  Public Const efDataTypeChar = "char"
  Public Const efDataTypeText = "text"
  Public Const efDataTypeInt = "int"
  Public Const efDataTypeDateTime = "datetime"
  Public Const efDataTypeMoney = "money"
  Public Const efDataTypeTinyInt = "tinyint"
  Public Const efDataTypeSmallInt = "smallint"
  Public Const efDataTypeBit = "bit"

Rem     Field Types
  Public Const efFldTypeText = 0
  Public Const efFldTypeUF = 1
  Public Const efFldTypeLookup = 2
  Public Const efFldTypeCheck = 3
  Public Const efFldTypeTextArea = 4
  Public Const efFldTypeRadio = 5
  Public Const efFldTypeSelect = 6
  Public Const efFldTypePassword = 7
  Public Const efFldTypeHTTP = 8
  Public Const efFldTypeEMail = 9
  Public Const efFldTypeSeparateDate = 10
  Public Const efFldTypeImage = 11
  Public Const efFldTypeFile = 12
  Public Const efFldTypeAtualization = 13
  Public Const efFldTypeColor = 14
  'Criado (Ridai)
  Public Const efFldTypeCountry = 15
  Public Const efFldTypeFTP = 16
  Public Const efFldTypeGopher = 17
  Public Const efFldTypeDateOfUpdate = 18
  Public Const efFldTypeDateOfDelete = 19
  Public Const efFldTypeDateOfInsert = 20
  '---------------
  
  'Criação Ruben Zevallos Jr.
  Public Const efFldTypeDateOfDeactivate = 21
  Public Const efFldTypeDateOfUnDelete = 22
  Public Const efFldTypeDateOfReactivate = 23
  Public Const efFldTypeUserIP = 24
  Public Const efFldTypeHTML = 25
  Public Const efFldTypeNumeric = 26
  Public Const efFldTypeTextEditor = 27
  Public Const efFldTypeTextAreaCounter = 28
  Public Const efFldTypeAlphabetic = 29
  Public Const efFldTypeCreditCard = 30
  Public Const efFldTypeCCAmex = 31
  Public Const efFldTypeCCVisa = 32
  Public Const efFldTypeCCMaster = 33
  Public Const efFldTypeMudule11 = 34
  Public Const efFldTypeMudule10 = 35
  '---------------

Rem     Validation Location
  Public Const efValLocClient = True
  Public Const efValLocServer = False

Rem     Validation Options
  Public Const efValOptNone = 0
  Public Const efValOptCGC = 1
  Public Const efValOptCPF = 2
  Public Const efValOptDate = 3
  Public Const efValOptSepDate = 4
  Public Const efValOptDateMToday = 5
  Public Const efValOptSepDateMToday = 6
  Public Const efValOptTime = 7
  Public Const efValOptEmail = 8
  Public Const efValOptCompareDates = 9
  Public Const efValOptCEP = 10
  'Criado (Ridai Govinda)
  Public Const efValOptCGCCPF = 11
  Public Const efValOptCPFCGC = 12
  '---------------------

  Public Const efValOptUserIP = 13
  Public Const efValOptNumeric = 14
  Public Const efValOptAlphabetic = 15
  Public Const efValOptCreditCard = 16
  Public Const efValOptCCAmex = 17
  Public Const efValOptCCVisa = 18
  Public Const efValOptCCMaster = 19
  Public Const efValOptMudule11 = 20
  Public Const efValOptMudule10 = 21
  Public Const efValOptRange = 22
  Public Const efValOptRangeNumeric = 23
  Public Const efValOptRangeAlphabetic = 24

Rem     Field Requirement
  Public Const efRequired = False
  Public Const efNotRequired = True

Rem     Relational Integrity
  Public Const efRelIntegrCascade = 0
  Public Const efRelIntegrRestricted = 1
  Public Const efRelIntegrNullifies = 2

Rem     Relation Type
  Public Const efRelType1to1 = 0
  Public Const efRelType1toN = 1

Rem     Field Disable Condition
  Public Const efCondDisable = True
  Public Const efCondEnable = False

Rem Operator types
  Public Const efBooleanOperatorOR = 1
  Public Const efBooleanOperatorAND = 2
  
Rem     Character Case
  Public Const efCharCaseNormal = 0
  Public Const efCharCaseUpper = 1
  Public Const efCharCaseLower = 2
  
'Criado (Ridai Govinda)
Rem AlphabeticIndex Types
  Public Const efAlphaIndexNumeric = 1
  Public Const efAlphaIndexAlphaNumeric = 2
  Public Const efAlphaIndexAlphabetic = 3
'---------------------

Rem ------------------------------------------------------
Rem Final das Constantes do EditForm
Rem ======================================================

Rem ======================================================
Rem Inicio das constantes do Form
Rem ------------------------------------------------------
Public Const frModeEdit = 0
Public Const frModeShow = 1
Public Const frModeValidate = 3
Rem ------------------------------------------------------
Rem Final das Constantes do Form
Rem ======================================================

Rem ======================================================
Rem Inicio das constantes do FormField
Rem ------------------------------------------------------
Rem Que tipos de actions o FormField.ValidateAction vai
Rem apresentar
Public Const ffVldActionButton = 0 'Botão
Public Const ffVldActionLink = 1   'Link
Public Const ffVldActionImage = 2  'Imagem
Rem ------------------------------------------------------
Rem Final das Constantes do FormField
Rem ======================================================

Rem ======================================================
Rem Inicio das constantes do String
Rem ------------------------------------------------------

Public Const stDateTypeAAMMDD = 0
Public Const stDateTypeAAAAMMDD = 1
Public Const stDateTypeMMDDAA = 2
Public Const stDateTypeMMDDAAAA = 3
Public Const stDateTypeDDMMAA = 4
Public Const stDateTypeDDMMAAAA = 5

Public Const stMoneyType9999D99 = 0
Public Const stMoneyType9D999D99 = 1
Public Const stMoneyType9999 = 2

Rem ------------------------------------------------------
Rem Final das Constantes do String
Rem ======================================================

Rem ======================================================
Rem Inicio das constantes do Show
Rem ------------------------------------------------------
Public Const shListTypeDisc = 1
Public Const shListTypeCircle = 2
Public Const shListTypeSquare = 3
Public Const shListTypeNumber = 4
Public Const shListTypeUChar = 5
Public Const shListTypeLChar = 6
Public Const shListTypeURoman = 7
Public Const shListTypeLRoman = 8
Rem ------------------------------------------------------
Rem Final das Constantes do Show
Rem ======================================================

Rem ======================================================
Rem Inicio das constantes do TableStyle
Rem ------------------------------------------------------

' Criar constantes: (Criadas Ridai)
' (Estilos de cores:
' 0 - Cabecalho cor 1, resto cor 2
' 1 - Alterna colunas
' 2 - Cabecalho fixo, alterna colunas
' 3 - Alterna linhas
' 4 - Cabecalho fixo, alterna linhas
' 5 - Todo da cor 1(menos o Título)
' 6 - Todo da cor 1(inclusive Título))

Public Const tbStFormatHeader = 0
Public Const tbStFormatHeaderAlternateColumns = 1
Public Const tbStFormatAlternateColumns = 2
Public Const tbStFormatHearderAlternateLines = 3
Public Const tbStFormatAlternateLines = 4
Public Const tbStFormatNothing = 5
Public Const tbStFormatTitle = 6

' (Estilo de bordas: (Ridai)
' 0 - Grade completa
' 1 - Apenas linhas
' 2 - Grade externa + linhas)
' else - sem grade (fundo da página), neste caso coloquei 4
' Não funciona (implementar):
' 3 - Apenas colunas
' 4 - Grade externa + colunas))
    
Public Const tbBdFormatCompleteGrid = 0
Public Const tbBdFormatOnlyLines = 1
Public Const tbBdFormatOutterGridLines = 2
Public Const tbBdFormatInvisible = 3

Rem ------------------------------------------------------
Rem Final das Constantes do TableStyle
Rem ======================================================

Rem ======================================================
Rem Início das Constantes do Validate
Rem ------------------------------------------------------

Rem Precedência de validação de CPF ou CGC
Rem ---------------------------------------
Public Const vlCp_CPFFirst = 0
Public Const vlCp_CGCFirst = 1

Rem Valores de Retorno do Validate.CompareDate
Rem ---------------------------------------
Public Const vlCmpDateFirstInvalid = 1
Public Const vlCmpDateSecondInvalid = 2
Public Const vlCmpDateBothInvalid = 3
Public Const vlCmpDateOutOfRange = 4

Rem ------------------------------------------------------
Rem Final das Constantes do Validate
Rem ======================================================

Rem ======================================================
Rem Final das Constantes de Proteção via HASP
Rem ------------------------------------------------------
Rem Seeds for Passwords
Public Const ZYC_SEED = 123
Public Const Query_SEED = 458

Rem Object Seeds
Public Const xtbBarcode = 23
Public Const xtbBrowse = 23
Public Const xtbCharacter = 23
Public Const xtbConnection = 23
Public Const xtbConfig = 23
Public Const xtbDatabase = 23
Public Const xtbDefault = 23
Public Const xtbEditForm = 23
Public Const xtbFile = 23
Public Const xtbFont = 23
Public Const xtbForm = 23
Public Const xtbInitializer = 23
Public Const xtbMenu = 23
Public Const xtbPack = 23
Public Const xtbPath = 23
Public Const xtbSecurity = 23
Public Const xtbUpload = 23
Public Const xtbFormField = 23
Public Const xtbRegistry = 23
Public Const xtbList = 23
'Criado (Ridai)
Public Const xtbSimpleForm = 23
'----
Public Const xtbShow = 23
Public Const xtbString = 23
Public Const xtbTable = 23
Public Const xtbTableStyle = 23
Public Const xtbTip = 23
Public Const xtbTreeView = 23
Public Const xtbURL = 23
Public Const xtbValidate = 23
Rem ------------------------------------------------------
Rem Final das Constantes de Proteção via HASP
Rem ======================================================

Rem ======================================================
Rem A list of HASP Services.
Rem ------------------------------------------------------
Public Const IS_HASP = 1
Public Const GET_HASP_CODE = 2
Public Const READ_WORD = 3
Public Const WRITE_WORD = 4
Public Const GET_HASP_STATUS = 5
Public Const GET_ID_NUM = 6
Public Const READ_MEMO_BLOCK = 50
Public Const WRITE_MEMO_BLOCK = 51
Public Const NET_SET_SERVER_BY_NAME = 96

' A list of TimeHASP Services.
Public Const TIMEHASP_SET_TIME = 70
Public Const TIMEHASP_GET_TIME = 71
Public Const TIMEHASP_SET_DATE = 72
Public Const TIMEHASP_GET_DATE = 73
Public Const TIMEHASP_WRITE_MEMORY = 74
Public Const TIMEHASP_READ_MEMORY = 75
Public Const TIMEHASP_WRITE_BLOCK = 76
Public Const TIMEHASP_READ_BLOCK = 77
Public Const TIMEHASP_GET_ID_NUM = 78


' A list of NetHASP Services.
Public Const NET_LAST_STATUS = 40
Public Const NET_GET_HASP_CODE = 41
Public Const NET_LOGIN = 42
Public Const NET_LOGOUT = 43
Public Const NET_READ_WORD = 44
Public Const NET_WRITE_WORD = 45
Public Const NET_GET_ID_NUMBER = 46
Public Const NET_SET_IDLE_TIME = 48
Public Const NET_READ_MEMO_BLOCK = 52
Public Const NET_WRITE_MEMO_BLOCK = 53


Public Const OK = 0
Public Const NET_READ_ERROR = 131
Public Const NET_WRITE_ERROR = 132


' A list of LptNum codes for the different types of keys
Public Const LPT_IBM_ALL_HASP25 = 0
Public Const LPT_IBM_ALL_HASP36 = 50
Public Const LPT_NEC_ALL_HASP36 = 60

Public Const MEMO_BUFFER_SIZE = 20

'TimeHASP maximum block size
Public Const TIME_BUFFER_SIZE = 16

' Show parameters
Public Const MODAL = 1

Public Const max_int = 32767

Public Const NO_HASP = "HASP plug not found !"

'
' The HASP memory buffer.
'
Public Type HBuff
    txt As String * 500
End Type

Global MemoHaspBuffer As HBuff
'Alterado (Ridai) Teste:
Global gdatServerStarted As Date

Rem
Rem Use this declaration to call the hasp() routine directly.
Rem

Public Declare Sub HaspLib Lib "haspvb32.dll" Alias "hasp" (ByVal Service&, ByVal seed&, ByVal LPT&, ByVal pass1&, ByVal pass2&, ByRef retcode1&, ByRef retcode2&, ByRef retcode3&, ByRef retcode4&)
Public Declare Sub WriteHaspBlock Lib "haspvb32.dll" (ByVal Service&, Buff As Any, ByVal Length&)
Public Declare Sub ReadHaspBlock Lib "haspvb32.dll" (ByVal Service&, Buff As Any, ByVal Length&)

Rem Para a função que retorna o Temp File
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


'Alterado (Ridai) Teste:
Sub Main()
   gdatServerStarted = Now
End Sub

Public Function GetDebugID() As Long
   Static lngDebugID As Long
   lngDebugID = lngDebugID + 1
   GetDebugID = lngDebugID
End Function
'---

Public Sub TimeBombX(ByVal ZSeed, ByVal XZP1, ByVal XZP2, ByVal XZP3, ByVal XZP4)
  If CDbl(Now) > htEvalExpiresAt Then
    Err.Raise 300, "HiperTools 3.0", htEvalExpiresString
  End If
End Sub

Rem =========================================================================
Rem TimeBombX - Proteção com o HASP (sendo executada em todos as Classes
Rem -------------------------------------------------------------------------
Public Sub TimeBombOld(ByVal ZSeed, ByVal XZP1, ByVal XZP2, ByVal XZP3, ByVal XZP4)
                     
  Const lconEvalExpiresAt = 36618
  Const lconHaspTestPeriod = (2 / 24)
  
  Dim SeedCode
  Dim LptNum
  Dim Password1
  Dim Password2
  Dim ZYC
  Dim Query
  
  Dim P1 'As Long
  Dim P2 'As Long
  Dim P3 'As Long
  Dim P4 'As Long
  
  Dim RC1
  
  Dim RC2
  Dim RC3
  Dim RC4
  
  Dim ZP1
  Dim ZP2
  Dim ZP3
  Dim ZP4
  
  
  ZP1 = XZP1
  ZP2 = XZP2
  ZP3 = XZP3
  ZP4 = XZP4
  
Rem  Dim Hasp As HiperTools30.Hasp
  
Rem  Set Hasp = CreateObject("HiperTools30.Hasp")
  Dim dateNext As Date, strAux As String
  
  Dim blnTeste As Boolean
  
  'Hasp.ZYC = 356 * 56 + 8543 + ZYC_SEED ' 28479
  ZYC = 356 * 56 + 8543 + ZYC_SEED ' 28479
 
  strAux = GetSetting("HiperTools30", "Timer", "Tick", "")
  
  If strAux > "" Then
    strAux = HTUncript(strAux)
  Else
    strAux = CStr(Now)
  End If
  
  'SaveSetting "HiperTools30", "Timer", "Tick2", strAux
  
  blnTeste = True
  
  If IsDate(strAux) Then
    dateNext = CDate(strAux)
    
    If Now >= dateNext Then
      dateNext = dateNext + lconHaspTestPeriod
      strAux = HTEncript(CStr(dateNext))
    
Rem      On Error GoTo Error
Rem      SaveSetting "HiperTools30", "Timer", "Tick", strAux
Rem      On Error GoTo 0
      
      blnTeste = False
      
      Call HaspLib(IS_HASP, SeedCode, LptNum, Password1, Password2, XZP1, XZP2, XZP3, XZP4)
      
      If (P1 <> 0) Then
        'Hasp.LptNum = LPT_IBM_ALL_HASP25
        'Hasp.Query = 24 * 1000 + 978 - Query_SEED ' 24978
        'Hasp.HaspCode ZSeed
        
        LptNum = LPT_IBM_ALL_HASP25
        Query = 24 * 1000 + 978 - Query_SEED ' 24978
        
        Call HaspLib(GET_HASP_CODE, ZSeed, LptNum, ZYC - ZYC_SEED, Query + Query_SEED, XZP1, XZP2, XZP3, XZP4)
        
        RC1 = XZP1
        RC2 = XZP2
        RC3 = XZP3
        RC4 = XZP4
        
        'If Not Hasp.RC1 = ZP1 And Not Hasp.RC2 = ZP2 And Not Hasp.RC3 = ZP3 And Not Hasp.RC4 = ZP4 Then
        If Not RC1 = ZP1 And Not RC2 = ZP2 And Not RC3 = ZP3 And Not RC4 = ZP4 Then
          blnTeste = True
        End If
      Else
        blnTeste = True
      End If
    End If
  Else
    Err.Raise 300, "HiperTools 3.0", "Houve violação das informações necessárias para validar sua licença. Para obter uma nova versão, por favor, contate a Zevallos Tecnologia em Informação no e-mail info@hipertools.com.br ou no telefone (61) 321-4711"
  End If
  
  If blnTeste Then
    If CDbl(Now) > lconEvalExpiresAt Then
      Err.Raise 300, "HiperTools 3.0", "O período de teste desta versão do HiperTools 3.0 chegou ao fim no dia 02/04/2000, para obter uma nova versão, por favor, contate a Zevallos Tecnologia em Informação no e-mail info@hipertools.com.br ou no telefone (61) 321-4711"
    End If
  End If
  
  Exit Sub
  'Set Hasp = Nothing
  
Error:
    'Err.Raise 300, "Teste", ShowError("TimeBombX")
    Err.Raise 300, "HiperTools30.Hasp", "O HASP não esta instalado!"

    
End Sub
