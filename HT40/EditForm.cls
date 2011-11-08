VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "EditForm"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------------------------------------------
' EditForm.cls
'-------------------------------------------------------------------------
' Descricao   : Objeto do sistema de Edicao
' Criacao     : 11:30h 22/02/1999
' Local       : Brasilia/DF
' Elaborado   : Eduardo Alves Cortes <edualves@zevallos.com.br>
'               Ruben Zevallos Jr. <zevallos@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 1999 by Zevallos(r) Tecnologia em Informacao
'-------------------------------------------------------------------------
' Atualização:
' Responsável : Ridai Govinda Pombo <ridai@zevallos.com.br>
' Versão      : 1.0.1b
' Copyright   : 2000 by Zevallos(r) Tecnologia em Informacao
'-------------------------------------------------------------------------

  Const conMoveFirst = "0"
  Const conMoveLast = "-1"

  Const constrProximo = "proximo "

  Const constrAdd = "Incluir"
  Const constrEdit = "Alterar"
  Const constrDelete = "Excluir"
  Const constrSave = "Salvar"
  Const constrFind = "Localizar"
  Const constrList = "Listagem"
  Const constrRefresh = "Atualizar"
  Const constrCopy = "Copiar"
  Const constrSaveAdd = "Salvar e Incluir"
  Const constrSaveCopy = "Salvar e Copiar"
  Const constrFirst = " |<  "
  Const constrPrevious = "  <  "
  Const constrNext = "  >  "
  Const constrLast = "  >| "
  Const constrTab = "Tab"

  Const constrOK = "   OK   "
  Const constrCancel = "Cancelar"
  Const constrYes = "   Sim   "
  Const constrNo = "   Não   "

  ' Linguagens
  Const efLanPortuguese = 0
  Const efLanEnglish = 1

  Public AlphabeticIndex As Boolean

  'Criado (Ridai Govinda)
  Private Enum Errors
    efErrorEmptyConnString = -200
    efErrorEmptyUserID
    efErrorEmptyPassword
    efErrorNoField
    efErrorInvalidIdentity
    efErrorGridNoTable
    efErrorNoLookUp
    efErrorTableNoStruct
    efErrorSQLQuery
    efErrorDataDefinition
    efErrorDuplicateKey
    efErrorFileHandling
    efErrorNullRequiredField
    efErrorCalcFieldNotFound
    efErrorBadLookup
    efErrorOtherPart
  End Enum
  '----------------------

  Private sparEditAction As String
  Private sparEditWhat As String
  Private sparEditMove As String
  Private sparEditEditable As String
  Private sparEditFind As String
  Private sparEditFilter As String
  Private sparEditList As String
  Private sparEditDefaults As String
  Private sparEditTab As String
  Private sparEditGrid As String

  Private sstrConnectionString As String
  Private sstrUserID As String
  Private sstrPassword As String

  Private sintPosition As Integer
  Private sblnFind As Boolean

  Private sstrText As String
  Private sstrText2 As String
  Private sstrLastValidate As String
  Private sintSize As Integer
  Private sintNumCol As Integer
  Private sstrCheckAsRadio As String
  Private sstrTab As String

  Private sintCurrentTableNum As Integer
  Private sintCurrentRelation As Integer
  Private sstrSubText As String
  Private sstrGroupFieldLastValue As String
  Private sstrGroupField As String
  Private sstrFields As String
  Private sstrFields2 As String
  Private sblnHaveSum As Boolean
  Private sblnHaveGraph As Boolean
  Private sintLastNumber As Integer
  Private sblnIsMyAction As Boolean
  Private sblnShowMessage As Boolean
  'Criado (Ridai)
  'Private sblnDoFPLPLink As Boolean
  Private sblnQSParsed As Boolean
  Private sblnShowSQL As Boolean
  Private sstrLinkTarget As String
  Private sintConnectionType As Integer
  Private sintLockTimeOut As Integer
  Private sbytLockRemainType As Byte


  'a ser Criado (Ridai) -- Objeto HiperTools30.EditErrorHandler
  'Private sobjEditErrorHandler As HiperTools30.EditErrorHandler
  '-- é melhor criar ErrorHandler
  '   com controle de CallStack
  '---

  Private sobjRS As ADODB.RecordSet
  Private sobjRS2 As ADODB.RecordSet
  Rem Private sobjCMD As ADODB.Command
  Rem Private sobjCMD2 As ADODB.Command
  Private sobjConn As ADODB.Connection
  Private sobjConn2 As ADODB.Connection

  'Alterado/Criado (Ridai)
  Public AllowEnumFields As Boolean
  Public AllowIdentityUpdate As Boolean
  'Mais propriedades (Documentar)
  ' + Força o CommonFind
  Public AllowCommonFind As Boolean
  ' + Permite que se apresente o IP do cliente na mensage de lock de registro.
  Public AllowOnLockShowIP As Boolean
Attribute AllowOnLockShowIP.VB_VarDescription = "Permite que se apresente o IP do cliente na mensage de lock de registro."
  Public LockShowIpText As String
Attribute LockShowIpText.VB_VarDescription = "Texto a ser apresentado quando a propriedade AllowOnLockShowIP for verdadeira."
  ' + Permite que se auto-atualize a página quando um registro estiver travado, fazendo
  '   com que fique-se tentando acessar o registro (ele irá conseguir quando der TimeOut)
  Public AllowOnLockRefresh As Boolean
  Public LockRefreshTime As Long
  Public LockRefreshText As String
  Public LockMessage As String
  '-----------------------------
  Public ShowFirstLastPageLink As Boolean
  Public ShowListUpperBar As Boolean
  Public ShowUnitUpperBar As Boolean
  Public ShowRecordCount As Boolean
  Public ShowButtons As Boolean
  Public ShowReportsLink As Boolean
  Public ShowSetOfOrdersLink As Boolean
  Public ListTitle As String
  Public UnitTitle As String
  Public ReportTitle As String
  Public GridTitle As String
  Public ListReportsCaption As String
  Public ListSetOfOrdersCaption As String

  Private sobjCurrConn As ADODB.Connection
  '---------------------
  Private sobjErrors As ADODB.Errors

  'Criado pelo (Ruben Zevallos)
  'Propriedades para o Formulário <FORM>
  '<FORM method=post action="submit.asp" AUTOCOMPLETE="OFF" Style = "style" Class = "Class">

  Public AutoComplete As Boolean
  Public StyleSheet As String
  Public Class As String
  Private sintLanguage As Integer

  Private sstrErrorMessage As String

  ' Removido (Ridai Govinda)
  ' // Não é utilizado em lugar algum!
  ' Private Security As HiperTools30.Security
  '------------------------
  Private Path As HiperTools30.Path
  Private Form As HiperTools30.Form
  Private sobjShow As HiperTools30.Show
  Private Table As HiperTools30.Table
  Private Browse As HiperTools30.Browse
  Private Database As HiperTools30.Database
  Private sobjString As HiperTools30.String
  Private Validate As HiperTools30.Validate
  Private Initializer As HiperTools30.Initializer
  
  Private FindFields As HiperTools30.FindFields
  
  Public Style As HiperTools30.TableStyle
  Public Buttons As HiperTools30.Buttons


  Private ScriptContext As ScriptingContext

  Private sstrComments As String, _
    sstrTitle As String, _
    sstrVersion As String, _
    sstrCompanyName As String, _
    sstrProductName As String, _
    sstrLegalCopyright As String

  Private Request As Request, _
    Response As Response, _
    Server As Server, _
    Session As Session, _
    Application As Application

  Option Explicit

Public Sub OnStartPage(ScriptContext As ScriptingContext)

  TimeBombX xtbEditForm, 15656, 12934, 3080, 14624

  Set ScriptContext = ScriptContext

  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application

  Set sobjRS = Server.CreateObject("ADODB.RecordSet")
  Set sobjRS2 = Server.CreateObject("ADODB.RecordSet")
  Rem Set sobjCMD = Server.CreateObject("ADODB.Command")
  Rem Set sobjCMD2 = Server.CreateObject("ADODB.Command")
  Set sobjConn = Server.CreateObject("ADODB.Connection")
  Set sobjConn2 = Server.CreateObject("ADODB.Connection")
  'Alterado/Criado (Ridai Govinda)
  Set sobjCurrConn = Server.CreateObject("ADODB.Connection")
  '---

  'Set Security = Server.CreateObject("HiperTools30.Security")
  Set Form = Server.CreateObject("HiperTools30.Form")
  Set sobjShow = Server.CreateObject("HiperTools30.Show")
  Set Browse = Server.CreateObject("HiperTools30.Browse")
  Set Database = Server.CreateObject("HiperTools30.Database")
  Set sobjString = Server.CreateObject("HiperTools30.String")
  Set Validate = Server.CreateObject("HiperTools30.Validate")
  Set Initializer = Server.CreateObject("HiperTools30.Initializer")
  Set Style = Server.CreateObject("HiperTools30.TableStyle")
  Set Buttons = Server.CreateObject("HiperTools30.Buttons")
  Set Table = Server.CreateObject("HiperTools30.Table")
  Set Path = Server.CreateObject("HiperTools30.Path")
  Set FindFields = Server.CreateObject("HiperTools30.FindFields")

  sblnIsMyAction = True
  sblnShowMessage = False
  sblnQSParsed = False

  sobjShow.ShowComments = False
  'Alterado/Criado (Ridai)
  ShowFirstLastPageLink = True
  sblnShowSQL = False
  sstrLinkTarget = ""
  sintLockTimeOut = Session.timeout  'minutos
  sbytLockRemainType = 0

  ListTitle = ""
  UnitTitle = ""
  ReportTitle = ""
  GridTitle = ""
  ShowRecordCount = True
  ShowButtons = True
  ShowUnitUpperBar = True
  ShowListUpperBar = True
  ShowReportsLink = True
  ShowSetOfOrdersLink = True
  AllowIdentityUpdate = False
  AllowCommonFind = False
  AllowOnLockRefresh = True
  AllowOnLockShowIP = True
  LockRefreshText = "Nova tentativa de alteração será executada em $d $s automaticamente."
  LockMessage = "Este registro está sendo alterado por outro usuário neste momento."
  LockShowIpText = "IP do host da máquina que bloqueou este registro: $s."
  AllowEnumFields = False
  ListReportsCaption = "Outras listagens"
  ListSetOfOrdersCaption = "Outras Ordenações"
  '---

  If Session("HTLanguage") > 0 Then
    Language = Session("HTLanguage")

  Else
    Language = 0

  End If

  AutoComplete = True

' (Faz a formatação default do Style, que é setado na Sub ShowActions para formatar as telas geradas em funções, como:
' ShowList, ShowScreen, FindFistPage e ShowSummary)
  Style.TitleFont.Size = 3
  Style.TitleFont.Bold = True
  Style.TitleColor = "white"
  Style.Color1 = "#CCCCFF"
  Style.Color2 = "#ECECFF"
  Style.Color3 = "#FFFFFF"
  Style.ColorFormat = 0
  Style.BorderFormat = 0
  Style.BorderColor = ""

Rem  sobjShow.CrLf = False

Rem  Set objHasp = Server.CreateObject("HaspLib.Hasp")

Rem  objHasp.LptNum = LPT_IBM_ALL_HASP25

Rem  Response.Write objHasp.IsHasp()

End Sub

Public Sub OnEndPage()
  Set ScriptContext = Nothing

  Set Request = Nothing
  Set Response = Nothing
  Set Server = Nothing
  Set Session = Nothing
  Set Application = Nothing
  Set Form = Nothing
  Set sobjShow = Nothing
  Set Table = Nothing
  Set Browse = Nothing
  Set sobjString = Nothing
  Set Validate = Nothing
  Set Initializer = Nothing
  Set Style = Nothing
  Set Path = Nothing
  'Set Security = Nothing

  Database.CloseObject sobjRS2
  Set sobjRS2 = Nothing

  Database.CloseObject sobjRS
  Set sobjRS = Nothing

  Rem Set sobjCMD = Nothing
  Rem Set sobjCMD2 = Nothing

  CloseConnection

  Set Database = Nothing

End Sub

Public Property Get ConnectionString() As String

  ' (Busca a conexão com o banco de dados, se esta não tiver sido setada então é dada uma mensagem de erro)
  'Alterado (Ridai Govinda)
  If Not sstrConnectionString > "" Then
    If Session("ConnectionString") > "" Then
       sstrConnectionString = Session("ConnectionString")
    Else
      'Alterado (Ridai)
      ShowError efErrorEmptyConnString, "EditForm", "ConnectionString", LanguageLocalizer(1)
    End If
  End If

  EstimateConnType

  ConnectionString = sstrConnectionString
  '------------------------

End Property

Public Property Let ConnectionString(strConnection As String)
' (Altera a conexão, no EditForm, com o banco de dados[função acima busca-a])
  sstrConnectionString = strConnection

  'Criado/Alterado (Ridai Govinda)
  EstimateConnType
  '------------------------------
End Property

Public Property Get UserID() As String
' (Busca o usuário do banco de dados, se esta não tiver sido setado então é dada uma mensagem de erro)
  If sstrUserID > "" Then
    UserID = sstrUserID
  Else
    If Session("RuntimeUserName") > "" Then
       UserID = Session("RuntimeUserName")
    Else
      UserID = sstrUserID
    End If
  End If
End Property

Public Property Let UserID(strUserID As String)
' (Altera o usuário, no EditForm, do banco de dados[função acima busca-o])
  sstrUserID = strUserID
End Property

Public Property Get Password() As String
' (Busca a senha do usuário do banco de dados, se esta não tiver sido setada então é dada uma mensagem de erro)
  If sstrPassword > "" Then
    Password = sstrPassword
  Else
    If Session("RunTimePassword") > "" Then
       Password = Session("RunTimePassword")
    Else
      Password = sstrPassword
    End If
  End If
End Property

Public Property Let Password(strPassword As String)
' (Altera a senha do usuário, no EditForm, do banco de dados[função acima busca-a])
  sstrPassword = strPassword
End Property

Public Property Get parAction() As String
' (Ação que está sendo tomada no momento pelo EditForm, este dado é buscado através da Request.QueryString(efQueryStrAction)
' na Sub QueryString do EditForm)
  parAction = sparEditAction
End Property

Public Property Get parWhat() As String
' (Tabela que está sendo apresentada no momento, este dado é buscado através da Request.QueryString(efQueryStrWhat) na Sub
' QueryString do EditForm)
  parWhat = sparEditWhat
End Property

Public Property Get parMove() As String
  parMove = sparEditMove
End Property

Public Property Get parEditable() As String
  parEditable = sparEditEditable
End Property

Public Property Get parFind() As String
  parFind = sparEditFind
End Property

Public Property Get parFilter() As String
  parFilter = sparEditFilter
End Property

Public Property Get parList()
  parList = sparEditList
End Property

Public Property Get parDefaults() As String
  parDefaults = sparEditDefaults
End Property

Public Property Get parTab() As String
  parTab = sparEditTab
End Property

Public Property Get parGrid() As String
  parGrid = sparEditGrid
End Property

Public Function IsMyAction() As Boolean
On Error GoTo ERROR_HANDLER
' (Esta função testa se o ASP que está sendo executado neste momento pertence ao EditForm. Ela o faz testando se alguma tabela
' foi definida através da Sub FormBegin. Se for uma função do EditForm então executa a Sub ShowActions, explicada abaixo)
  IsMyAction = sblnIsMyAction
  If Not sblnIsMyAction Then
    ShowActions
  End If

ERROR_HANDLER_END:
  On Error Resume Next

  Err.Clear

  If sstrErrorMessage > "" Then
    ShowError efErrorOtherPart, "EditForm", "ShowActions", sstrErrorMessage
    
  End If
Exit Function

ERROR_HANDLER:
  sstrErrorMessage = "Err.Number: " & Err.number & "<BR>Err.Description: " & Err.Description

  GoTo ERROR_HANDLER_END
End Function

Public Sub ShowActions()
On Error GoTo ERROR_HANDLER

' (Esta Sub chama as funções que montam telas, decidindo qual montar através do parametro sparEditAction vindo da QueryString)
  Set Table.Style = Style

  Select Case sparEditAction
    Case efQSActionScreen
      ShowScreen

    Case efQSActionEditor
' (Session("RecordAction") indica que tipo de edição está sendo feita, encontrada várias vezes no código, quando é esvaziada, como
' na função abaixo significa que esta não é uma função de edição)
      Session("RecordAction") = ""
      ReadFieldValues False
      ShowScreen
      QSSave

    Case efQSActionList
      Session("RecordAction") = ""
      ShowList
      QSSave

    Case efQSActionSummary
      Session("RecordAction") = ""
      ShowSummary
      QSSave

    Case efQSActionCommonFind, efQSActionAdvancedFind
      FindFirstPage
      QSSave

    Case efQSActionGrid
' (Chama a Sub que monta o grid, não utiliza as funções de start de html do default, para economizar código HTML no frame)
      sobjShow.HTMLCR "<HTML><BASEFONT FACE=""Arial, Helvetica, Sans-Serif""><BODY BGCOLOR=WHITE LINK=BLACK VLINK=BLACK ALINK=BLACK TOPMARGIN=0 LEFTMARGIN=0>"
      ShowGrid
      sobjShow.HTMLCR "</BODY></HTML>"
' (Dentro da Sub ShowGrid a tabela atual é alterada para a tabela representada no frame. No código abaixo a tabela atual volta a ser
' a tabela do form)
      CurrentTable = Session("LastCurrentTable")
    'Criado (Ridai Govinda)
    Case efQSActionGetFile
      ShowUpFileAction

  End Select

' (Guarda a última ação feita pelo EditForm, com a resalva de que a tela atual seja diferente da anterior, por exemplo no caso do uso
' do refresh. Para que se volte, num cancelamento, para a tela anterior)
  If sparEditAction <> Session("EditAction") Then
    Session("LastEditAction") = Session("EditAction")
    Session("EditAction") = sparEditAction
    QSSave
  End If

ERROR_HANDLER_END:
  On Error Resume Next
  
  Err.Clear

  If sstrErrorMessage > "" Then
    ShowError efErrorOtherPart, "EditForm", "ShowActions", sstrErrorMessage
  
  End If
Exit Sub

ERROR_HANDLER:
  sstrErrorMessage = "Err.Number: " & Err.number & "<BR>Err.Description: " & Err.Description

  GoTo ERROR_HANDLER_END

End Sub

Public Sub RedirectActions()
  On Error GoTo ERROR_HANDLER

' (Esta Sub chama as funções que não montam telas, por exemplo as Subs que fazem gravação no BD e que processam os Form
' das telas, decidindo que função chamar através do parametro sparEditAction vindo da QueryString, o mesmo do ShowActions)
  Select Case sparEditAction
    Case efQSActionExeEdition
      ExeEdition

    Case efQSActionAdd, efQSActionCopy
' (Session("RecordAction") indica que tipo de edição está sendo feita, encontrada várias vezes no código, quando possui valor "1",
' como na função abaixo significa que esta é uma função de inclusão)
      Session("RecordAction") = "1"
      EditClear sparEditAction
      Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionScreen & _
      WhatWhereHow

    Case efQSActionEdit
      Session("RecordAction") = "2"
      If ReadFieldValues(True) Then
' (Este aplication indica que este usuário começou a editar o registro indicado pelo MakeTarget)
        'Alterado (Ridai Govinda)
        'Application.Lock
        'Application(CurrentTable & MakeTarget) = Now()
        'Application.Unlock
        RecordLocked efLockDoLock

' (Este session indica que registro está sendo editado pelo usuário)
        Session("InEdition") = CurrentTable & MakeTarget
        Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionScreen & _
        WhatWhereHow & GetTab

      Else
        Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionEditor & _
        WhatWhereHow & GetTab

      End If

    Case efQSActionExeFind
      ExeFindFirstPage

  End Select

ERROR_HANDLER_END:
  On Error Resume Next

  Err.Clear

  If sstrErrorMessage > "" Then
    ShowError efErrorOtherPart, "EditForm", "RedirectActions", sstrErrorMessage
    
  End If
  
Exit Sub

ERROR_HANDLER:
  sstrErrorMessage = "Err.Number: " & Err.number & "<BR>Err.Description: " & Err.Description

  GoTo ERROR_HANDLER_END

End Sub

Public Sub ExeEdition()
  On Error GoTo ERROR_HANDLER

' (Na tela de edição todos os botões são retornados através do hidden "btmEditor".
' Esta Sub lê o botão clicado e chama a função que realiza a tarefa relacionada ao botão.

' Se a função de um botão precisar ser alterada, é a partir daqui que se deve começar a fazê-lo)
  Select Case Request.Form("btmEditor")
' (Botão Incluir)
    Case Buttons.AddCaption
      Session("RecordAction") = "1"
      EditClear efQSActionAdd
      Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionScreen & WhereHow

' (Botão Copiar)
    Case Buttons.CopyCaption
      Session("RecordAction") = "1"
      EditClear efQSActionCopy
      Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionScreen & WhereHow

' (Botão Alterar)
    Case Buttons.EditCaption
' (Session("RecordAction") indica que tipo de edição está sendo feita, encontrada várias vezes no código, quando possui valor "2",
' como na função abaixo significa que esta é uma função de alteração)
      Session("RecordAction") = "2"
      If ReadFieldValues(True) Then
'        Application.Lock
'        Application(CurrentTable & MakeTarget) = Now()
'        Application.Unlock
        'Criado (Ridai Govinda)
        EditClear2
        RecordLocked efLockDoLock
        '--------------------
        Session("InEdition") = CurrentTable & MakeTarget
        Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionScreen & _
        WhatWhereHow & GetTab

      Else
        Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionEditor & _
        WhatWhereHow & GetTab

      End If

' (Botão Excluir)
    Case Buttons.DeleteCaption
      Session("RecordAction") = ""
      EditDelete
' (Nas condições abaixo se testa a posição do registro excluído, se for o primeiro então vai para o primeiro novamente
' [efQueryStrMove & "=" & conMoveFirst], se for o último então vai para o último novamente [efQueryStrMove & "=" & conMoveLast],
' senão vai para a posição anterior ao do registro excluído[efQueryStrMove & "=" & Session("AbsolutePosition") - 2])
      If Session("AbsolutePosition") = 1 Then
        Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionEditor & _
        WhereHow & "&" & efQueryStrMove & "=" & conMoveFirst

      ElseIf Session("AbsolutePosition") = Session("RecordCount") Then
        Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionEditor & _
        WhereHow & "&" & efQueryStrMove & "=" & conMoveLast

      Else
        Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionEditor & _
        WhatWhereHow & "&" & efQueryStrMove & "=" & Session("AbsolutePosition") - 2

      End If

' (Botão Salvar)
    Case Buttons.SaveCaption
      GetForm Session("LastTab")
      MakeServerValidation
      EditSave

' (Este aplication indica que este usuário começou a editar o registro indicado pelo MakeTarget, neste caso como o registro está sendo
' salvo, saindo do modo de alteração, então o parametro está sendo esvaziado)
'      Application.Lock
'      Application(CurrentTable & MakeTarget) = ""
'      Application.Unlock
      'Criado (Ridai Govinda)
      RecordLocked efLockForceClearLock
      '---------------------

' (Este session indica que registro está sendo editado pelo usuário, neste caso como o registro está sendo salvo, saindo do modo de
' alteração,  então o parametro está sendo esvaziado)
      Session("InEdition") = ""

      Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionEditor & WhatWhereHow & GetTab

' (Botão Salvar e Incluir)
    Case Buttons.SaveAddCaption
      GetForm Session("LastTab")
      MakeServerValidation
      EditSave

'      Application.Lock
'      Application(CurrentTable & MakeTarget) = ""
'      Application.Unlock

      'Criado (Ridai Govinda)
      RecordLocked efLockForceClearLock
      '---------------------

      Session("InEdition") = ""
      Session("RecordAction") = "1"
      EditClear efQSActionAdd
      Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionScreen & WhatWhereHow


' (Botão Salvar e Copiar)
    Case Buttons.SaveCopyCaption
      GetForm Session("LastTab")
      MakeServerValidation
      EditSave

'      Application.Lock
'      Application(CurrentTable & MakeTarget) = ""
'      Application.Unlock

      'Criado (Ridai Govinda)
      RecordLocked efLockForceClearLock
      '---------------------

      Session("InEdition") = ""
      Session("RecordAction") = "1"
      EditClear efQSActionCopy
      Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionScreen & WhatWhereHow

' (Botão Listagem)
    Case Buttons.ListCaption
      Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionList & _
      WhatList & GetFilter

' (Botão Cancelar)
    Case Buttons.CancelCaption
'      Application.Lock
'      Application(CurrentTable & MakeTarget) = ""
'      Application.Unlock

      'Criado (Ridai Govinda)
      RecordLocked efLockForceClearLock
      '---------------------

      Session("InEdition") = ""
      Response.Redirect QSLast
      Session("RecordAction") = ""

' (Botão Primeiro Registro)
    Case Buttons.FirstCaption
      Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionEditor & _
      WhereHow & "&" & efQueryStrMove & "=" & conMoveFirst & GetFilter

' (Botão Registro Anterior)
    Case Buttons.PreviousCaption
      Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionEditor & _
      WhereHow & "&" & efQueryStrMove & "=" & Session("AbsolutePosition") - 2 & GetFilter

' (Botão Próximo Registro)
    Case Buttons.NextCaption
      Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionEditor & _
      WhereHow & "&" & efQueryStrMove & "=" & Session("AbsolutePosition") & GetFilter

' (Botão Último Registro)
    Case Buttons.LastCaption
      Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionEditor & _
      WhereHow & "&" & efQueryStrMove & "=" & conMoveLast & GetFilter

' (Botão Localizar)
    Case Buttons.FindCaption
      If Session(CurrentTable & "FieldFind") > "" Then
        Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionAdvancedFind & _
        WhereHow

      Else
        Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionCommonFind & _
        WhereHow

      End If

  End Select

' (Faz a ação do clique de um tab)
  If Request.Form("hidEditor") > "" Then
    GetForm Session("LastTab")
    If Session("IsEdition") > "" Or Session("RecordAction") > "" Then
      Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionScreen & _
      WhatWhereHow & "&" & efQueryStrTab & "=" & Request.Form("hidEditor")

    Else
      Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionEditor & _
      WhatWhereHow & "&" & efQueryStrTab & "=" & Request.Form("hidEditor")

    End If
  End If

ERROR_HANDLER_END:
  On Error Resume Next

  Err.Clear

  If sstrErrorMessage > "" Then
    ShowError efErrorOtherPart, "EditForm", "ExeEdition", sstrErrorMessage
    
  End If
  
Exit Sub

ERROR_HANDLER:
  sstrErrorMessage = "Err.Number: " & Err.number & "<BR>Err.Description: " & Err.Description

  GoTo ERROR_HANDLER_END

End Sub

'Criado (Ridai Govinda)
Public Sub OpenConnection()
  If Not sblnQSParsed Then ParseQueryString

  sobjConn.ConnectionTimeout = Session("ConnectionTimeout")
  sobjConn.CommandTimeout = Session("CommandTimeout")
  sobjConn.Open ConnectionString, UserID, Password

  sobjConn2.ConnectionTimeout = Session("ConnectionTimeout")
  sobjConn2.CommandTimeout = Session("CommandTimeout")
  sobjConn2.Open ConnectionString, UserID, Password
End Sub

Rem Público para se poder forçar o fechamento da conexão, documentar que é perigoso utilizar
Rem este método em lugares errados
Public Sub CloseConnection()
  Database.CloseObject sobjConn
  Set sobjConn = Nothing

  'Alterado (Ridai)
  Database.CloseObject sobjConn2
  Set sobjConn2 = Nothing

  Database.CloseObject sobjCurrConn
  Set sobjCurrConn = Nothing
  '---
End Sub
'---------------------
Public Sub ParseQueryString()
' (Busca todos os QueryStrings utilizados pelo EditForm e os guarda em parametros. Esta Sub tb seta os objetos do ADO)
  sparEditAction = LCase(Request.QueryString(efQueryStrAction))
  sparEditWhat = LCase(Request.QueryString(efQueryStrWhat))
  sparEditMove = LCase(Request.QueryString(efQueryStrMove))
  sparEditEditable = LCase(Request.QueryString(efQueryStrEditable))
  sparEditFind = LCase(Request.QueryString(efQueryStrFind))
  sparEditFilter = LCase(Request.QueryString(efQueryStrFilter))
  sparEditList = LCase(Request.QueryString(efQueryStrList))
  sparEditDefaults = Request.QueryString(efQueryStrDefaults)
  sparEditTab = Request.QueryString(efQueryStrTab)
  sparEditGrid = Request.QueryString(efQueryStrGrid)

  Session("OrderField") = LCase(Request.QueryString(efQueryStrOrderField))
  Session("OrderDesc") = LCase(Request.QueryString(efQueryStrOrderDesc))

  If Not sparEditEditable > "" Then
    sparEditEditable = 0

  End If

  If Not sparEditTab > "" Then
    sparEditTab = 0

  End If

  sblnQSParsed = True

End Sub

Public Sub QueryString()
  ParseQueryString
  OpenConnection
End Sub

Public Sub DataTable(ByVal strTable As String)
  Session("TableNumber") = Session("TableNumber") + 1

  Session("Table" & Session("TableNumber")) = strTable

  CurrentTable = strTable
  sintCurrentTableNum = Session("TableNumber")
  Session(strTable & "FieldNumber") = 0
  Session(strTable & "RelationNumber") = 0
  Session(strTable & "IndexNumber") = 0
  Session(strTable & "IndexClusteredName") = ""

End Sub

Rem Definições de dados do EditForm
Public Sub DataBegin(ByVal strTable As String)
' (Inicio da definição do dicionário de dados de uma tabela, passada como parametro)
  DataTable strTable
End Sub

Public Sub DataAddField(ByVal strField As String, ByVal strType As String, ByVal intSize As Long, ByVal blnNull As Boolean, Optional ByVal intPrec As Long = -1)
' (Definição de um campo, parametros:nome do campo, tipo, tamanho, requerimento)
  Dim i As Integer
  Session(CurrentTable & "FieldNumber") = Session(CurrentTable & "FieldNumber") + 1

  i = Session(CurrentTable & "FieldNumber")

  'Alterado/Criado (Ridai Govinda)
  If strField > "" And strType > "" And intSize > 0 Then
    Session(CurrentTable & i & "Field") = strField
    Session(CurrentTable & i & "Size") = intSize
    Session(CurrentTable & i & "Type") = strType
    Session(CurrentTable & i & "Null") = blnNull
    Session(CurrentTable & i & "IsKey") = False
    Session(CurrentTable & i & "Prec") = intPrec

  Else
    ShowError efErrorDataDefinition, "EditForm", "DataAddField", LanguageLocalizer(2)

  End If
  '---------------------------

End Sub

Public Sub DataAddPrimaryKey(ByVal strField As String)
' (Definição de campo chave, podem ser definidas N chaves, no caso de chaves mulpiplas)
  If IsFieldExist(strField, "DataAddPrimaryKey") Then
    Session(CurrentField & "IsKey") = True

  End If

End Sub

Public Sub DataAddRelation(ByVal strTable As String, ByVal intDeleteOption As Integer, ByVal intUpdateOption As Integer, ByVal intRelationType As Integer)
' (Definição de relacionamento da tabela, parametros:Tabela a ser relacionada, opção para exclusão(ex.: Cascata), opção para gravação(ex.: Cascata), tipo de relacionamento(ex.: 1 para N))
Dim i As Integer
  Session(CurrentTable & "RelationNumber") = Session(CurrentTable & "RelationNumber") + 1

  i = Session(CurrentTable & "RelationNumber")

  Session(CurrentTable & i & "RelationTable") = strTable
  Session(CurrentTable & i & "DeleteOption") = intDeleteOption
  Session(CurrentTable & i & "UpdateOption") = intUpdateOption
  Session(CurrentTable & i & "RelationType") = intRelationType
  Session(CurrentTable & i & "FieldtoFieldNumber") = 0

End Sub

Public Sub DataAddRelationFields(ByVal strField As String, ByVal strField2 As String)
' (Definição de campos de relacionamento entre tabelas. Serve, apenas, para o último relacionamento definido. Parametros: campo
' de relacionamento na tabela principal, campo de relacionamento na tabela relacionada)
Dim i, j As Integer

  i = Session(CurrentTable & "RelationNumber")

  j = Session(CurrentTable & i & "FieldtoFieldNumber")

  Session(CurrentTable & i & "FieldtoFieldNumber") = j + 1

  j = j + 1

  Session(CurrentTable & i & "," & j & "RelationField1") = strField
  Session(CurrentTable & i & "," & j & "RelationField2") = strField2

End Sub

Public Sub DataAddIndex(ByVal strName As String, ByVal strContent As String, Optional ByVal blnUnique As Boolean = False)
' (Definição de índices a serem criados junto com a tabela, parametros: nome do índice, campos)
Dim i As Integer
  Session(CurrentTable & "IndexNumber") = Session(CurrentTable & "IndexNumber") + 1

  i = Session(CurrentTable & "IndexNumber")

  Session(CurrentTable & i & "IndexName") = strName
  Session(CurrentTable & i & "IndexContent") = strContent
' Acrescentado por Ueslei/Kelson
  Session(CurrentTable & i & "IndexUnique") = blnUnique

End Sub

Public Sub DataIndexClustered(ByVal strName As String, ByVal strContent As String)
' (Definição de índices aninhados a serem criados junto com a tabela, parametros: nome do índice, campos)
  Session(CurrentTable & "IndexClusteredName") = strName
  Session(CurrentTable & "IndexClusteredContent") = strContent

End Sub

Public Sub FormBegin(ByVal strTable As String, ByVal strCaption As String, ByVal intEditOption As Integer, ByVal strWhatList As String, _
                         ByVal blnClientValidation As Boolean)
    BeginForm strTable, strCaption, intEditOption, strWhatList, blnClientValidation
End Sub

Rem Definições de formulário do EditForm
Public Sub BeginForm(ByVal strTable As String, ByVal strCaption As String, ByVal intEditOption As Integer, ByVal strWhatList As String, _
                         ByVal blnClientValidation As Boolean)
' (Inicio da definição de formulário, parametros: nome da tabela, rótulo, opção de edição, identificador da tabela a ser
' usado nos QueryStrings, forma de validação(Browser ou ASP))
  sblnIsMyAction = False
  CurrentTable = strTable
  Session(strTable & "Caption") = strCaption
  Session(strTable & "EditOption") = intEditOption
  Session(strTable & "WhatList") = strWhatList
  Session(strTable & "ClientValidation") = blnClientValidation

  Session(strTable & "ListNumber") = 0
  Session(strTable & "OrderNumber") = 0
  Session(strTable & "HeaderNumber") = 0
  Session(strTable & "FilterNumber") = 0
  Session(strTable & "MarkNumber") = 0
  Session(strTable & "InsertLink") = True
  Session(strTable & "UnitTabs") = ""

End Sub

Public Sub FormEnd()
    EndForm
End Sub

Public Sub EndForm()
' (Finalização da definição de formulário.
' Monta os JOINs que serão usados para buscar campos estrangeiros em suas tabelas, e os guarda no
' Session(CurrentTable & "JoinFormUnit"))
Dim strJoinFormUnit As String
Dim i As Integer

  strJoinFormUnit = ""

  For i = 1 To Session(CurrentTable & "FieldNumber")
    Session("CurrentField") = i
    If IsExternalField Then
      strJoinFormUnit = MakeJoin(CurrentTable, Session(CurrentField & "OwnerTable"), strJoinFormUnit)

    End If
  Next

  Session(CurrentTable & "JoinFormUnit") = strJoinFormUnit
Rem  sobjShow.MessageError "*" & Session(CurrentTable & "JoinFormUnit") & "*"

End Sub

Public Sub AddField(ByVal strField As String, ByVal strCaption As String, Optional ByVal intEditOption As Integer = efFldTypeText, Optional ByVal intValidationOption As Integer = 0, _
                            Optional ByVal intShowOption As Integer = 0, Optional ByVal DefaultValue As String = "", Optional ByVal strRequireMessage As String = "", _
                            Optional ByVal blnPersistentDefValue As Boolean = False)
' (Define um campo de formulário, este campo pode ser externo à tabela[Ex.:"TabelaExemplo.CampoExemplo"],  parametros:nome
' do campo, rótulo, opção de edição, opção de validação, opção de apresentação, valor default na inclusão, mensagem para quando
' este campo estiver vazio)
Dim i As Integer
' (Testa se o campo está definido no dicionário de dados da tabela, ou se é um campo externo[InStr(strField, ".") <> 0])
  If IsFieldExist(strField, "AddField") Or (InStr(strField, ".") <> 0) Then
    If (InStr(strField, ".") <> 0) Then
      Session(CurrentTable & "FieldNumber") = Session(CurrentTable & "FieldNumber") + 1
      Session("CurrentField") = Session(CurrentTable & "FieldNumber")
      i = Session("CurrentField")

    End If
    Session(CurrentField & "Field") = strField
    Session(CurrentField & "Caption") = strCaption
    Session(CurrentField & "EditOption") = intEditOption
    Session(CurrentField & "ValidationOption") = intValidationOption
    Session(CurrentField & "ShowOption") = intShowOption
    Session(CurrentField & "DefaultValue") = DefaultValue
    Session(CurrentField & "RequireMessage") = strRequireMessage
    Session(CurrentField & "ValueNumber") = 0
    Session(CurrentField & "DisableNumber") = 0
    Session(CurrentField & "ListChars") = 100
    Session(CurrentField & "ExternalAddress") = False
    'Criado (Ridai Govinda)
    Session(CurrentField & "PersistentDefValue") = blnPersistentDefValue
    '---------------------

    'Criado (Ruben Zevallos Jr.)
    Session(CurrentField & "AutoComplete") = True
    Session(CurrentField & "ReadOnly") = False
    Session(CurrentField & "Disabled ") = False

    Session(CurrentField & "FontBold") = False
    Session(CurrentField & "FontItalic") = False
    Session(CurrentField & "FontUnderLine") = False

    Session(CurrentField & "CaptionFontBold") = True
    Session(CurrentField & "CaptionFontItalic") = False
    Session(CurrentField & "CaptionFontUnderline") = False

    Session(CurrentField & "FontSize") = ""
    Session(CurrentField & "CaptionFontSize") = ""
    '---------------------

    'Alterado (Ridai Govinda)
    If Not IdentityCheck() Then
      ShowError efErrorInvalidIdentity, "EditForm", "AddField", sobjString.FormatText(LanguageLocalizer(3), strField, CurrentTable)

    End If
    '----------------------

    ' (Os comandos abaixo limpam as propriedades do campo, no caso deste já estar definido no session)
    ClearSession CurrentField & "AlphabeticIndex"
    ClearSession CurrentField & "Hint"
    ClearSession CurrentField & "IsUnique"
    ClearSession CurrentField & "ExternalLink"
    ClearSession CurrentField & "InternalLink"
    ClearSession CurrentField & "ShowSize"
    ClearSession CurrentField & "BeginYear"
    ClearSession CurrentField & "EndYear"
    ClearSession CurrentField & "LookupTable"
    ClearSession CurrentField & "LookupField"
    ClearSession CurrentField & "LookupFieldText"
    ClearSession CurrentField & "ListSize"

    'Criado pelo (Ruben Zevallos Jr.)
    ClearSession CurrentField & "TabIndex"

    ClearSession CurrentField & "StyleSheet"
    ClearSession CurrentField & "Class"
    ClearSession CurrentField & "Align"

    ClearSession CurrentField & "FontFace"
    ClearSession CurrentField & "FontColor"
    ClearSession CurrentField & "FontBGColor"

    ClearSession CurrentField & "CaptionStyleSheet"
    ClearSession CurrentField & "CaptionClass"
    ClearSession CurrentField & "CaptionAlign"

    ClearSession CurrentField & "CaptionFontFace"
    ClearSession CurrentField & "CaptionFontColor"
    ClearSession CurrentField & "CaptionFontBGColor"

    '-------------------------

' (Aplicação dos valores default às propriedades do campo)
    Session(CurrentField & "CharCase") = efCharCaseNormal
    Session(CurrentField & "RadioColumns") = 1
    If intEditOption = efFldTypeTextArea Or intEditOption = efFldTypeHTML Or intEditOption = efFldTypeTextAreaCounter Then
      Session(CurrentField & "TextAreaHeight") = 4

    End If

' (Definições de um campo externo à tabela)
    If (InStr(strField, ".") <> 0) Then
      Dim strCurrentField, strCurrentTable As String

      strCurrentField = CurrentField
      strCurrentTable = CurrentTable

' (Tabela mãe do campo)
      CurrentTable = Left(strField, InStr(strField, ".") - 1)

      If IsFieldExist(Mid(strField, InStr(strField, ".") + 1), "AddField") Then
        Session(strCurrentTable & i & "ExternalAddress") = True
        Session(strCurrentTable & i & "OwnerTable") = CurrentTable
        Session(strCurrentTable & i & "Field") = Mid(strField, InStr(strField, ".") + 1)
        Session(strCurrentTable & i & "Size") = Session(CurrentField & "Size")
        Session(strCurrentTable & i & "Null") = Session(CurrentField & "Null")
        Session(strCurrentTable & i & "Type") = Session(CurrentField & "Type")

      End If
      Session("CurrentField") = i
      CurrentTable = strCurrentTable

    End If
  End If

End Sub

Public Sub AddOrder(ByVal strOrder As String)
' (Adiciona uma forma de ordenação à tabela)
Dim i As Integer
  Session(CurrentTable & "OrderNumber") = Session(CurrentTable & "OrderNumber") + 1
  i = Session(CurrentTable & "OrderNumber")

  Session(CurrentTable & i & "Order") = strOrder

End Sub

Public Sub AddGrid(ByVal strTable, ByVal strFields, ByVal strName, ByVal intOption)
' (Define um grid que pode ser chamado pelo nome no FormUnit, parametros: tabela do grid, campos
' que apareceram no grid, nome, identificador do grid a ser usado nos QueryStrings)
Dim i As Integer
  If Session(strTable & "FieldNumber") > "" Then
    Session(CurrentTable & "GridNumber") = Session(CurrentTable & "GridNumber") + 1
    i = Session(CurrentTable & "GridNumber")

    Session(CurrentTable & i & "GridTable") = strTable
    Session(CurrentTable & i & "GridFields") = strFields
    Session(CurrentTable & i & "GridName") = strName
    Session(CurrentTable & i & "GridOption") = intOption

  Else
    ShowError efErrorGridNoTable, "EditForm", "AddGrid", sobjString.FormatText(LanguageLocalizer(4), strTable)

  End If

End Sub

Public Sub AddFilter(ByVal strField As String, ByVal strValue As String, _
                     Optional ByVal strOperator As String = "")
' (Adiciona um filtro à tabela, parametros: campo e seu valor a ser usado no filtro)
  Dim i As Integer
  Session(CurrentTable & "FilterNumber") = Session(CurrentTable & "FilterNumber") + 1
  i = Session(CurrentTable & "FilterNumber")

  Session(CurrentTable & i & "FilterField") = strField
  Session(CurrentTable & i & "FilterValue") = strValue

  If Session(CurrentTable & "FilterBooleanOperator") > "" Then
    Session(CurrentTable & i & "FilterBooleanOperator") = Session(CurrentTable & "FilterBooleanOperator")
  End If

  If strOperator > "" Then
    Session(CurrentTable & i & "FilterOperator") = strOperator
  End If

End Sub

' Criado:
Public Property Let FilterBooleanOperator(ByVal byteOperator As Byte)
'Troca os PutAnds por PutOrs na hora de montar o filter
    Session(CurrentTable & "FilterBooleanOperator") = byteOperator
End Property

Public Property Let OrderBy(ByVal strOrder As String)
' (Adiciona uma forma de ordenação prioritária à tabela)

  Session(CurrentTable & "OrderBy") = strOrder

End Property

Public Property Get OrderBy() As String
  OrderBy = Session(CurrentTable & "OrderBy")
End Property
'---

Public Sub AddMark(ByVal strWhere As String, Optional ByVal strBGColor As String = "", Optional ByVal strTextColor As String = "", Optional ByVal strField As String = "")
' (Adiciona uma marca de linha, parametrtos: WHERE que indica que linhas marcar, cor das linhas,
' cor dos textos, campo que será marcado ao invés de marcar a linha)
Dim i As Integer
  Session(CurrentTable & "MarkNumber") = Session(CurrentTable & "MarkNumber") + 1
  i = Session(CurrentTable & "MarkNumber")

  Session(CurrentTable & i & "MarkWhere") = strWhere
  Session(CurrentTable & i & "MarkBGColor") = strBGColor
  Session(CurrentTable & i & "MarkFontColor") = strTextColor

' (Procura o campo a ser marcado se encontrar guarda-o como campo marcado, senão guarda vazio)
  If strField > "" Then
    If FindField(CurrentTable, strField) Then
      Session(CurrentTable & i & "MarkField") = Session("CurrentField")

    Else
      Session(CurrentTable & i & "MarkField") = ""

    End If
  Else
    Session(CurrentTable & i & "MarkField") = ""

  End If

End Sub

Public Sub FieldAlphabeticIndex(ByVal strField As String, ByVal strNumber As String)
' (Define a forma de apresentação deste campo no índice alfabético das listagens)

  If IsFieldExist(strField, "FieldAlphabeticIndex") Then
    Session(CurrentField & "AlphabeticIndex") = strNumber

  End If

End Sub

Public Sub FieldIsUnique(ByVal strField As String)
' (Define que o campo recebe um só valor para toda a tabela)
  If IsFieldExist(strField, "FieldIsUnique") Then
    Session(CurrentField & "IsUnique") = "1"

  End If

End Sub

Public Sub FieldTextAreaHeight(ByVal strField As String, ByVal intHeight As Integer)
' (Define a altura de um campo do tipo TextArea)
  If IsFieldExist(strField, "FieldTextAreaHeight") Then
    Session(CurrentField & "TextAreaHeight") = intHeight

  End If

End Sub

Public Sub FieldFile(ByVal strField As String, ByVal strFileRoot As String, ByVal strFileNamePrefix As String, _
                     Optional ByVal strExtsList As String = "*")
' (Define um campo arquivo, parametros: nome do campo, diretório onde deveram ser colocados os
' arquivos deste campo, prefixo do nome dos arquivos)
  If IsFieldExist(strField, "FieldFile") Then
    Session(CurrentField & "FileRoot") = strFileRoot
    Session(CurrentField & "FileNamePrefix") = strFileNamePrefix
    'Criado (Ridai Govinda)
    Session(CurrentField & "ExtensionsList") = strExtsList
    '---------------------

  End If

End Sub

Public Sub FieldImage(ByVal strField As String, ByVal strImageRoot As String, ByVal strWidth As String, _
                      ByVal strImageNamePrefix As String, Optional ByVal strExtsList As String = "jpg,jpeg,png,gif,bmp")
' (Define um campo imagem, parametros: nome do campo, diretório onde deveram ser colocadas as
' imagens deste campo, largura para apresentação da imagem, prefixo do nome dos arquivos das
' imagens)
  If IsFieldExist(strField, "FieldImage") Then
    Session(CurrentField & "ImageRoot") = strImageRoot
    Session(CurrentField & "ImageWidth") = strWidth
    Session(CurrentField & "ImageNamePrefix") = strImageNamePrefix
    'Criado (Ridai Govinda)
    Session(CurrentField & "ExtensionsList") = strExtsList
    '---------------------

  End If

End Sub

Public Sub FieldRadioColumns(ByVal strField As String, ByVal intColumns As Integer)
' (Define o nº de colunas de um campo Radio)
  If IsFieldExist(strField, "FieldRadioColumns") Then
    Session(CurrentField & "RadioColumns") = intColumns

  End If

End Sub

Public Sub EmptyListMessage(Optional ByVal strMessage As String = "", Optional ByVal blnInsertLink As Boolean = True)
' (Define a mensagem a ser dada no caso da listagem estar vazia, parametros:mensagem, deve ser
' colocado um link de inserção?)
  Session(CurrentTable & "EmptyListMessage") = strMessage
  Session(CurrentTable & "InsertLink") = blnInsertLink

End Sub

Public Sub FieldCharCase(ByVal strField As String, ByVal intCase As Integer)
' (Define a Captulaçao dos caracteres do texto de um campo, parametros: nome do campo, captulação
' ex.:Todos os caracteres maiúsculos)
  'Criado (Ridai Govinda) -- Para teste do String.FormatText
  Const lconMessage = "O campo ""$s"", da tabela ""$s"", não existe"

  If IsFieldExist(strField, "FieldCharCase") Then
    Session(CurrentField & "CharCase") = intCase

  End If

End Sub

'Criado (Ridai Govinda)
Public Sub FieldDisableButton(ByVal strField As String, ByVal vntOnValue As Variant, _
                              ByVal strBttnCaption As String, Optional ByVal intOperator As Integer = efBttnOperatorEqual)
  'Define uma condição
  If IsFieldExist(strField, "FieldDisableButton") Then
    Session(CurrentField & "DisableButton") = strBttnCaption
    Session(CurrentField & "DisableBttnOnValue") = vntOnValue

    Session(CurrentField & "DisableBttnOperator") = intOperator

  End If
End Sub
'--------------------------

Public Sub FieldDisableValue(ByVal strField As String, ByVal strField2 As String, ByVal Value As String, ByVal blnDisable)
' (Define as formas de desabilitação de um campo, parametros: campo desabilitado, campo desabilitante
' , valor de desbilitação, opção de abilitação ou desabilitação para o valor)
Dim i, j, k As Integer
Dim blnFind As Boolean
  If IsFieldExist(strField, "FieldDisableValue") Then
    If IsFieldExist(strField2, "FieldDisableValue") Then
      i = 1
      blnFind = False
' (Procura por uma outra desabilitação contendo os dois campos)
      Do While i <= Session(CurrentField & "DisableNumber") And Not blnFind
        If Session(CurrentField & "," & i & "DisableField") = strField Then
          blnFind = True

        Else
          i = i + 1

        End If

      Loop

' (Se tiver encontrado uma outra desabilitação então adiciona o valor na desabilitação, senão adiciona a
' desabilitação)
      If blnFind Then
        j = Session(CurrentField & "," & i & "DisableValueNumber") + 1

      Else
        j = 1
        i = Session(CurrentField & "DisableNumber") + 1
        Session(CurrentField & "DisableNumber") = i
        Session(CurrentField & "," & i & "DisableField") = strField

      End If
      Session(CurrentField & "," & i & "DisableValueNumber") = j
      Session(CurrentField & "," & i & "DisableCondition") = blnDisable
' (Define os valores para desabilitação. Se o campo for radio então o valor é a posição do valor passado por parametro, no radio
' , senão é o próprio valor passado por parametro)
      If Session(CurrentField & "EditOption") = efFldTypeRadio Then
        For k = 1 To Session(CurrentField & "ValueNumber")
          If CStr(Session(CurrentField & "," & k & "Value")) = CStr(Value) Then
            Session(CurrentField & "," & i & "," & j & "DisableValue") = k - 1

          End If
        Next
      Else
        Session(CurrentField & "," & i & "," & j & "DisableValue") = Value

      End If
    End If
  End If

End Sub

Public Sub FieldListSize(ByVal strField As String, ByVal intSize As Integer)
' (Define a largura de um campo na listagem, parametros:nome do campo, largura)
  If IsFieldExist(strField, "FieldListSize") Then
    Session(CurrentField & "ListSize") = intSize

  End If

End Sub

Public Sub FieldListChars(ByVal strField As String, ByVal intSize As Integer)
' (Define a largura de um campo na listagem em caracteres, parametros:nome do campo, nº de caracteres)
  If IsFieldExist(strField, "FieldListChars") Then
    Session(CurrentField & "ListChars") = intSize

  End If

End Sub

Public Sub FieldMask(ByVal strField As String, ByVal strMask As String, Optional ByVal strCompleter As String = "")
' (Define a máscara a ser colocada em um campo pela função Mask do objeto String, parametros:nome
' do campo, máscara, completador da máscara)
  If IsFieldExist(strField, "FieldMask") Then
    Session(CurrentField & "Mask") = strMask
    Session(CurrentField & "MaskCompleter") = strCompleter

  End If

End Sub

Public Sub FormFind(ByVal strFieldFind As String)
' (Define os campos a serem utilizados na localização)
  Session(CurrentTable & "FieldFind") = strFieldFind

End Sub

Public Sub FormList(ByVal strFieldList As String)
' (Define os campos a serem utilizados na listagem)
  Session(CurrentTable & "FieldList") = strFieldList
  TotalOperation strFieldList, "ListOperationNumber", False, ""

End Sub

Public Sub FormUnit(ByVal strFieldUnit As String)
' (Define os campos a serem utilizados na edição. A formação da string a ser utilizada segue a seguinte
' regra:
' conteúdo = [nome do campo] ou [texto HTML a ser apresentado]
' separador = ",", ";", "-", "=", "|"
' conteúdo + separador + conteúdo
' Cada separador possui um significado na montagem da tela:
' "," = Quebra de coluna
' ";" = Quebra de linha
' "-" = Linha separadora
' "=" = Título separador [Ex.:blablabla=Título separador,blablabla]
' "|" = Separador de Tabs[Citados mais especificamente na sub adiante])
  Session(CurrentTable & "FieldUnit") = strFieldUnit

End Sub

Public Sub FormTabs(ByVal strUnitTabs As String)
' (Define os títulos dos tabs, que são unidades de formulário utilizadas para diminuir a quantidade de
' campos apresentados por vez ao usuário. A formação da string a ser utilizada segue a seguinte
' regra: título + "|" + título)
  Session(CurrentTable & "UnitTabs") = strUnitTabs

End Sub

Private Sub TabTask(ByVal strTab As String, ByVal strTask As String)
' (Define ações a serem executadas no inicio da montagem de um tab, esta opção não está funcionando
' desde o HiperTools foi componentizado, pois o VB não possui uma função estilo Execute para executar
' um código VBScript no código externo ao objeto)
  Session(CurrentTable & strTab & "Task") = strTask

End Sub

Public Sub FieldShowSize(ByVal strField As String, ByVal intSize As Integer)
' (Define o tamanho de apresentação de um campo, parametros: nome do campo, tamanho)
  If IsFieldExist(strField, "FieldShowSize") Then
    Session(CurrentField & "ShowSize") = intSize

  End If

End Sub

Public Sub FieldSepDate(ByVal strField As String, ByVal intBeginYear As Integer, ByVal intEndYear As Integer)
' (Define os anos de início e de fim na apresentação de um campo de data separada[dia,mês,ano cada
' um em um campo separado], parametros: nome do campo, primeiro ano, último ano)
  If IsFieldExist(strField, "FieldSepDate") Then
     Session(CurrentField & "BeginYear") = intBeginYear
     Session(CurrentField & "EndYear") = intEndYear

  End If

End Sub
Public Sub FieldExternalLink(ByVal strField As String, ByVal strLink As String, _
                             Optional ByVal strQSParam As String = "", Optional ByVal strShowField As String = "")
' (Define um link externo ao EditfForm para um campo)
  Dim strFieldFound As String

  If IsFieldExist(strField, "FieldExternalLink") Then
    Session(CurrentField & "ExternalLink") = strLink

    strFieldFound = CurrentField

    If strQSParam > "" Then
      Session(CurrentField & "ExtLinkQSParam") = strQSParam

      If strShowField > "" Then
        If IsFieldExist(strShowField, "FieldExternalLink") Then
          Session(strFieldFound & "ExtLinkField") = strShowField

        End If
      Else
        Session(CurrentField & "ExtLinkField") = strField

      End If

    End If
  End If

End Sub

Public Sub FieldInternalLink(ByVal strField As String, ByVal strLink As String, ByVal strILTarget As String, Optional ByVal strILEditable As Boolean = True)
' (Define um link interno do EditfForm para um campo, parametros:nome do campo, tabela a ser lincada,
' campo de ligação, tabela a ser lincada entra em modo de edição?)
  If IsFieldExist(strField, "FieldInternalLink") Then
    Session(CurrentField & "InternalLink") = strLink
    Session(CurrentField & "ILTargetNumber") = 1
    Session(CurrentField & "ILEditable") = strILEditable
    Session(CurrentField & ",1ILTarget") = strILTarget

  End If

End Sub

Public Sub FieldAddILTarget(ByVal strField As String, ByVal strILTarget As String)
' (Adiciona um campo de ligação para o InternalLink, parametros: nome do campo que possui o link,
' nome do campo da tabela lincada)
  If IsFieldExist(strField, "FieldAddILTarget") Then
    Session(CurrentField & "ILTargetNumber") = Session(CurrentField & _
      "ILTargetNumber") + 1
    Session(CurrentField & "," & Session(CurrentField & "ILTargetNumber") & "ILTarget") = strILTarget

  End If

End Sub

Public Sub FieldAddValue(ByVal strField As String, ByVal Value, ByVal strCaption As String)
' (Adiciona valores ao campo do tipo select, radio ou checkbox. Parametros: nome do campo, valor
' adicionado, texto de rótulo relacionado ao valor)
  If IsFieldExist(strField, "FieldAddValue") Then
    Session(CurrentField & "ValueNumber") = Session(CurrentField & "ValueNumber") + 1

    Session(CurrentField & "," & Session(CurrentField & "ValueNumber") & "Value") = Value
    Session(CurrentField & "," & Session(CurrentField & "ValueNumber") & "Caption") = strCaption

  End If

End Sub

Public Sub FieldHint(ByVal strField As String, ByVal strHint As String)
' (Define uma explicação básica para preenchimento do campo, parametros:nome do campo, texto
' explicativo)
  If IsFieldExist(strField, "FieldHint") Then
    Session(CurrentField & "Hint") = strHint

  End If

End Sub

Public Sub FieldLookup(ByVal strField As String, ByVal strLookUpTable As String, ByVal strLookupField As _
                                String, ByVal strLookupFieldText As String, _
                                Optional ByVal blnPopup As Boolean = False, Optional ByVal strTitle As String = "")
' (Campo lookup são os que seu conteúdo é chave estrangeira numa relação 1 para N, este campo é
' apresentado na forma de select, e sua definição é o motivo desta função, parametros: nome do campo,
' tabela ligada, campo de ligação na tabela ligada, campo de apresentação que caracteriza o registro
' ligado[existe o intuito de transformar este parametro de um campo para coleção
' de campos representados por uma fórmula como a do FormUnit, que permitiria apresentar dentro do
' select mais de um campo separados da forma que o programador desejar], especifica se o Lookup será feito em uma
' nova janela para não haver problemas com janelas muito grandes.)

  If IsFieldExist(strField, "FieldLookup") Then
    Session(CurrentField & "LookupTable") = strLookUpTable
    Session(CurrentField & "LookupField") = strLookupField
    Session(CurrentField & "LookupFieldText") = strLookupFieldText
    Session(CurrentField & "LookupFieldPopup") = blnPopup

    If strTitle > "" Then _
    Session(CurrentField & "LookupFieldPopupTitle") = strTitle

    If Not (blnPopup) And InStr(strLookupFieldText, ",") > 0 Then
      ShowError efErrorBadLookup, "EditForm", "FieldLookup", LanguageLocalizer(5)
    End If

  End If

End Sub

Public Sub AddHeader(ByVal strHeader As String, ByVal intBegin As Integer, ByVal intWidth As Integer)
' (Adiciona um super-título a cabeçalhos, parametros: texto do título, início[coluna da listagem], largura)
Dim i As Integer
  Session(CurrentTable & "HeaderNumber") = Session(CurrentTable & "HeaderNumber") + 1
  i = Session(CurrentTable & "HeaderNumber")

  Session(CurrentTable & i & "HeaderText") = strHeader
  Session(CurrentTable & i & "HeaderBegin") = intBegin
  Session(CurrentTable & i & "HeaderWidth") = intWidth

End Sub

'Criado (Ridai Govinda)
Public Sub AddReport(ByVal strList As String, ByVal strCaption As String, ByVal strName As String)
  AddList strList, strCaption, strName
End Sub
'---------------------

Public Sub AddList(ByVal strList As String, ByVal strCaption As String, ByVal strName As String)
' (Adiciona um relatório, parametros: conteúdo do relatório[explicado a seguir], título do relatório,
' identificador do relatório na QueryString
' O conteúdo do relatório é designado a partir da regra:
' Campo = [nome de um campo existente na tabela] ou ["#" + nome do campo calculado{criado a partir
' de uma fórmula}]
' Parametros do campo:
' "s" = soma dos valores do campo
' "c" = contagem dos registros do agrupamento
' "g" = com gráfico
' "a" = campo agrupador
' "t" = totalizado na mudança do valor deste campo
' "f" = tipo de dados de um campo calculado
' "l" = rótulo
' "o" = operação
' "v" = valor do primeira linha em um contador montador pela operação
' Formato:
' [Campo([Parametros do campo])])
Dim i As Integer
  Session(CurrentTable & "ListNumber") = Session(CurrentTable & "ListNumber") + 1
  i = Session(CurrentTable & "ListNumber")

  Session(CurrentTable & i & "ListName") = strName
  Session(CurrentTable & i & "List") = strList
  Session(CurrentTable & i & "ListCaption") = strCaption

End Sub

Public Sub FieldLookupFilter(ByVal strField As String, ByVal strDependentField As String, Optional ByVal strFilterField As String = "", Optional ByVal strValue As String = "", Optional ByVal strType As String = "int")
' (Os campos tipo lookup podem receber um filtro na seleção de seus itens. Este filtro pode ser por um
' relacionamento de campos ou pelo valor do campo.
' Esta sub define o filtro, parametros: nome do campo, campo da tabela de lookup que será utilizado no
' filtro, campo desta tabela, valor, tipo do campo de ligação)
  If IsFieldExist(strField, "FieldLookupFilter") Then
    Session(CurrentField & "LookupFilterField") = strFilterField
    Session(CurrentField & "DependentField") = strDependentField
    Session(CurrentField & "LookupFilterValue") = strValue
    Session(CurrentField & "FilterFieldType") = strType

  End If

End Sub

Public Sub FieldDoubleValidate(ByVal strField As String, ByVal strField2 As String, ByVal strMessage As String)
' (Define uma validação que envolva dois campos, como uma comparação de datas ou confirmação de
' senha, parametros:primeiro campo da validação, segundo campo, mensagem de erro se os valores
' forem inválidos)
  If IsFieldExist(strField, "FieldDoubleValidate") Then
    Session(CurrentField & "DoubleValidateField") = strField2
    Session(CurrentField & "DoubleValidateMessage") = strMessage

  End If

End Sub

'Criado (Ridai Govinda)
Private Function FixExtRelationJoin(ByVal strBaseJoin As String, ByVal strExtJoin As String)
  Dim strResult As String, strFix As String
  Dim intJoins As Integer, lngPos As Long, i As Integer

  intJoins = sobjString.CountSubString(strExtJoin, "LEFT")

  If intJoins > 1 Then
    lngPos = 1
    For i = 1 To 2
      lngPos = InStr(lngPos + 1, strExtJoin, "LEFT")
    Next
    strFix = Replace(strExtJoin, "LEFT", ") LEFT", lngPos)
    strExtJoin = Left(strExtJoin, lngPos - 1) & strFix
  End If

  strResult = strBaseJoin
  If strExtJoin > "" Then
    strResult = strResult & strExtJoin

  End If

  If intJoins > 1 Then strResult = String(intJoins - 1, "(") & strResult
  FixExtRelationJoin = strResult

End Function

Private Function FixJoinForOracle(ByVal strSQL As String) As String
  'Função que retorna o join corrigido para o Oracle (Join "Intrínscico")...
  Dim arrAux, arrAux2, arrSQL
  Dim i As Long
  Dim blnGroup As Boolean
  Dim strAnd As String

  blnGroup = False
  FixJoinForOracle = strSQL

  If InStr(UCase(strSQL), " JOIN ") Then
    strSQL = Replace(strSQL, vbCrLf, "")
    If InStr(UCase(strSQL), " GROUP ") Then blnGroup = True
    strSQL = Replace(strSQL, "SELECT ", "", 1, -1, 1)
    strSQL = Replace(strSQL, " FROM ", "^^", 1, -1, 1)

    If blnGroup Then
      strSQL = Replace(strSQL, "GROUP ", "^^ ^^", 1, -1, 1)
    Else
      strSQL = Replace(strSQL, " WHERE ", "^^", 1, -1, 1)
    End If

    strSQL = Replace(strSQL, " ORDER ", "^^", 1, -1, 1)

    arrSQL = Split(strSQL, "^^")

    arrSQL(1) = Replace(arrSQL(1), "(", "", 1, -1, 1)
    arrSQL(1) = Replace(arrSQL(1), ")", "", 1, -1, 1)
    If Not blnGroup Then arrSQL(2) = "(" & arrSQL(2) & ")"

    arrAux = Split(arrSQL(1), " LEFT JOIN ", -1, 1)

    For i = 0 To UBound(arrAux)

      arrAux2 = Split(arrAux(i), " ON ", -1, 1)
      arrAux(i) = Trim(arrAux2(0))

      strAnd = " AND "
      If blnGroup And Not Trim(arrSQL(2)) > "" Then strAnd = " "
      If UBound(arrAux2) > 0 Then arrSQL(2) = arrSQL(2) & strAnd & arrAux2(1) & " (+)"

      arrAux2 = ""

    Next

    arrSQL(1) = Join(arrAux, ", ") & " "
    FixJoinForOracle = "SELECT " & arrSQL(0) & " FROM " & arrSQL(1) & " WHERE " & arrSQL(2)

    If blnGroup Then
      FixJoinForOracle = FixJoinForOracle & " GROUP " & arrSQL(3) & " ORDER " & arrSQL(4)
    Else
      FixJoinForOracle = FixJoinForOracle & " ORDER " & arrSQL(3)
    End If

  End If

End Function
'---------

Private Function MakeJoin(ByVal strTable As String, ByVal strTable2 As String, ByVal strJoin As String)
' (Esta função foi feita com o intuito de facilitar a montagem de joins nas funções em geral, tem uma natureza semi-recursiva, pois seu
' retorno é utilizado como um de seus parametros.
' Monta um JOIN entre duas tabelas somado ao já feito, parametros: tabelas a serem ligadas, join já existente no select)
Dim strJoinAux, strTableAux, strFirstTable, strSecondTable As String
Dim i As Integer

  strJoinAux = strJoin
  If InStr(strJoinAux, " " & strTable2 & " ") = 0 Then
    If FindRelation(strTable, strTable2) Then
      strTableAux = strTable & CurrentRelation
      strFirstTable = strTable
      strSecondTable = strTable2

    ElseIf FindRelation(strTable2, strTable) Then
      strTableAux = strTable2 & CurrentRelation
      strFirstTable = strTable2
      strSecondTable = strTable

    End If
    If Session(strTableAux & "RelationType") = efRelType1to1 Then
      If strTableAux > "" Then
        If Session(strTableAux & "FieldtoFieldNumber") <> 0 Then
          strJoinAux = strJoinAux & " LEFT JOIN " & strTable2

        End If

        For i = 1 To Session(strTableAux & "FieldtoFieldNumber")
          If i = 1 Then
            strJoinAux = strJoinAux & " ON "

          Else
            strJoinAux = strJoinAux & " AND "

          End If
          strJoinAux = strJoinAux & strSecondTable & "." & Session(strTableAux & "," & i & "RelationField2") & " = " & strFirstTable & "." & Session(strTableAux & "," & i & "RelationField1")

        Next
      End If
    End If
  End If
  MakeJoin = strJoinAux

End Function

Public Property Get CurrentTable() As String
' (Tabela que está sendo utilizada)
  CurrentTable = Session("CurrentTable")

End Property
Public Property Let CurrentTable(strText As String)
  Session("CurrentTable") = strText

End Property

Public Property Get CurrentField() As String
' (Campo que está sendo utilizado)
  CurrentField = CurrentTable & Session("CurrentField")

End Property
Public Property Let CurrentField(strText As String)
  Session("CurrentField") = strText

End Property

Private Property Get CurrentGrid() As String
' (Grid que está sendo utilizado)
  CurrentGrid = Session("CurrentGrid")

End Property
Private Property Let CurrentGrid(strText As String)
  Session("CurrentGrid") = strText

End Property

Private Property Get CurrentList() As String
' (Relatório que está sendo apresentado)
  CurrentList = CurrentTable & Session("CurrentList")

End Property
Private Property Let CurrentList(strText As String)
  Session("CurrentList") = strText

End Property

Private Property Get CurrentRelation() As String
' (Relacionamento que está sendo apresentado)
  CurrentRelation = sintCurrentRelation

End Property
Private Property Let CurrentRelation(strText As String)
  sintCurrentRelation = strText

End Property

Public Function FindField(ByVal strTable As String, ByVal strField As String) As Boolean
' (Procura por um campo, se encontrar coloca o campo como CurrentField e retorna true, parametros:
' nome da tabela, nome do campo)
Dim i As Integer
  i = 1
  Do While (LCase(Trim(Session(strTable & i & "Field"))) <> LCase(Trim(strField))) And _
           (i < Session(strTable & "FieldNumber"))
    i = i + 1

  Loop

  If LCase(Trim(Session(strTable & i & "Field"))) = LCase(Trim(strField)) Then
    FindField = True
    CurrentField = i

  Else
    FindField = False

  End If

End Function

Private Function FindGrid(ByVal strGrid As String) As Boolean
' (Procura por um grid, se encontrar coloca o grid como CurrentGrid e retorna true, parametros:
' nome do grid)
Dim i As Integer
  i = 1
  Do While (LCase(Session(CurrentTable & i & "GridName")) <> LCase(strGrid)) And _
           (i < Session(CurrentTable & "GridNumber"))

    i = i + 1

  Loop

  If LCase(Session(CurrentTable & i & "GridName")) = LCase(strGrid) Then
    FindGrid = True
    Session("CurrentGrid") = i

  Else
    FindGrid = False

  End If

End Function

Private Function FindList(ByVal strList As String) As Boolean
' (Procura por um relatório, se encontrar coloca o relatório como CurrentList e retorna true, parametros:
' nome do relatório)
Dim i As Integer
  i = 1
  Do While (LCase(Session(CurrentTable & i & "ListName")) <> LCase(strList)) And _
           (i < Session(CurrentTable & "ListNumber"))

    i = i + 1

  Loop

  If LCase(Session(CurrentTable & i & "ListName")) = LCase(strList) Then
    FindList = True

    Session("CurrentList") = i

  Else
    FindList = False

  End If

End Function

Private Function FindRelation(ByVal strTable As String, ByVal strTable2 As String) As Boolean
' (Procura por um relation, se encontrar coloca o relation como CurrentList e retorna true, parametros:
' nome do relation)
Dim i As Integer
  i = 1
  Do While (LCase(Session(strTable & i & "RelationTable")) <> LCase(strTable2)) And _
           (i < Session(strTable & "RelationNumber"))
    i = i + 1

  Loop

  If LCase(Session(strTable & i & "RelationTable")) = LCase(strTable2) Then
    FindRelation = True
    CurrentRelation = i

  Else
    FindRelation = False

  End If


End Function

Private Function OwnerTable(ByVal strTable As String, ByVal intField As String) As String
' (Retorna a tabela a qual pertence um campo, parametros:nome da tabela, nome do campo)
  If Session(strTable & intField & "OwnerTable") > "" Then
    OwnerTable = Session(strTable & intField & "OwnerTable")

  Else
    OwnerTable = strTable

  End If

End Function

Public Function FieldValue(ByVal strTable As String, ByVal strField As String) As Variant
' (Retorna o valor de um campo, parametros:nome da tabela, nome do campo)
Dim strCurrentField As String
  strCurrentField = Session("CurrentField")
  If FindField(strTable, strField) Then
    FieldValue = Session(strTable & Session("CurrentField") & "Value")

  Else
    FieldValue = ""

  End If
  Session("CurrentField") = strCurrentField

End Function

Private Sub FindFirstPage()
' (Apresenta a tela de localização)
Dim i, j As Integer
 Session("LastOption") = sparEditAction

  sobjShow.HTMLCR "<STYLE>"
  sobjShow.HTMLCR "INPUT {font: 8pt Arial}"
  sobjShow.HTMLCR "</STYLE>"

' (Apresenta os campos tipo checkbox como radio)
  CheckAsRadio

  sobjShow.BR
  sobjShow.Center

' (Apresenta a localização montada até o momento através do botão "ou")
  MakeFindOr

' (Apresenta as mensagens de erro)
  If Message > "" Then
    If Session(CurrentTable & "ClientValidation") Then
      sobjShow.JSAlert Message

    Else
      Table.BeginTable "60%", "", 1
      Table.BeginRow 2
      Table.BeginColumn

      sobjShow.HTMLCR Message

      Table.EndColumn
      Table.EndRow
      Table.EndTable

    End If
    Message = ""

  End If

  Form.BeginForm Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionExeFind & WhatList & GetEditable, "frmFind", "POST", "", ""

  Table.BeginTable "85%", sobjString.FormatText(LanguageLocalizer(13), Session(CurrentTable & "Caption")), 4, True

  Session("NumberOfFinds") = 0
  Session("LastAction") = sparEditAction

  If sparEditAction = efQSActionCommonFind Then
' (Tela de localização comum)
    Table.ColumnAlign = "left"
    Table.BeginRow 1
    Table.Column LanguageLocalizer(14)
    Table.EndRow

    Table.BeginRow 1
    Table.ColumnVAlign = "Top"
    Table.BeginColumn

' (Monta o select com os campos para localização. Apenas os campos do tipo text, e que sejam da tabela, não sejam externos, são apresentados)
    Form.SelectBegin "txtCampo1", 1, False
    For i = 1 To Session(CurrentTable & "FieldNumber")
      Session("CurrentField") = i
      If (Session(CurrentField & "EditOption") = efFldTypeText) And (Not IsExternalField) And (EditCaption(False) > "") Then
        Form.SelectOption Session(CurrentField & "Field"), EditCaption(False), False

      End If

    Next
    Form.SelectEnd

    Form.text "txtValor", 40, "", 100
    sobjShow.Nbsp
    sobjShow.HTMLCR LanguageLocalizer(73)
    Form.CheckBox "strIgual", 1, "left", 0
    sobjShow.BR
    Form.Hidden "txtOperador1", "2"

    Session("NumberOfFinds") = 1

  Else
' (Tela de localização avançada)
    BeginReadOperation Session(CurrentTable & "FieldFind")

    Table.BeginRow 1
    Table.Column LanguageLocalizer(74)
    Table.Column LanguageLocalizer(75)
    Table.Column LanguageLocalizer(76)
    If Session(CurrentTable & "ListNumber") > 0 Then
' (Monta o select com os relatórios)
      Table.ColumnAlign = "right"
      Table.BeginColumn
      Response.Write LanguageLocalizer(15)
      Form.SelectBegin "selLists"
      For i = 1 To Session(CurrentTable & "ListNumber")
        Form.SelectOption Session(CurrentTable & i & "ListName"), Session(CurrentTable & i & "ListCaption")

      Next
      Form.SelectEnd
      Table.EndColumn
      Table.ColumnAlign = ""

    End If
    Table.EndRow

    Do While IsThereOperation
      Select Case NextOperation
         Case ","

         Case Else
           If FindField(CurrentTable, LastOperation) Then
' (Apresenta a linha de um campo, com seu caption, um hidden com o nome, um select de operadores
' possíveis e o input para ser escolhido o valor a ser localizado)
             Session("NumberOfFinds") = Session("NumberOfFinds") + 1
             Table.BeginRow 1

' (Caption)
             Table.BeginColumn
             ShowCaption False
' (Hidden)
             Form.Hidden "txtCampo" & Session("NumberOfFinds"), Session(CurrentField & "Field")
             Table.EndColumn
             Table.BeginColumn

             Dim strEditOption As String
               strEditOption = Session(CurrentField & "EditOption")

' (Operadores)
             If (strEditOption = efFldTypeText) Or _
                (strEditOption = efFldTypeEMail) Or _
                (strEditOption = efFldTypeHTTP) Or _
                (strEditOption = efFldTypeAtualization) Or _
                (strEditOption = efFldTypeSeparateDate) Or _
                (strEditOption = efFldTypeTextArea) Or _
                (strEditOption = efFldTypeTextAreaCounter) Or _
                (strEditOption = efFldTypeHTML) Then

               Form.SelectBegin "txtOperador" & Session("NumberOfFinds"), 1, False
               Form.SelectOption 1, TextualOperators(1), True

               If (Session(CurrentField & "Type") = efDataTypeVarChar) Or (Session(CurrentField & "Type") = efDataTypeChar) Then
                 For j = 2 To 4
                   Form.SelectOption j, TextualOperators(j), False

                 Next

               Else
                 For j = 5 To 8
                   Form.SelectOption j, TextualOperators(j), False

                 Next

               End If
               Form.SelectOption 9, TextualOperators(9), False
               Form.SelectEnd

             Else
' (Se o único operador for o igual, é criado não um select mas um hidden)
               sobjShow.HTMLCR LanguageLocalizer(16)
               Form.Hidden "txtOperador" & Session("NumberOfFinds"), "1"

             End If

             Table.EndColumn

             If Session(CurrentTable & "ListNumber") > 0 Then
               Table.ColumnColSpan = "2"

             End If

' (Input)
             Table.BeginColumn
             ShowField False, True, True, True
             Table.EndColumn

             Table.ColumnColSpan = ""

             Table.EndRow

          End If

      End Select

    Loop
    Table.ColumnColSpan = 4

  End If

  Table.BeginRow 1
  Table.BeginColumn
  sobjShow.Center

' (Barra de botões)
  Form.Submit "btmLocaliza", LanguageLocalizer(17)
  sobjShow.Nbsp
  sobjShow.Nbsp
  sobjShow.Nbsp
  Form.Submit "btmLocaliza", LanguageLocalizer(18)
  Form.Submit "btmLocaliza", LanguageLocalizer(19)
  If Session(CurrentTable & "ListNumber") > 0 Then
    Form.Submit "btmLocaliza", LanguageLocalizer(20)

  End If
  Form.Reset LanguageLocalizer(21)
  Form.Submit "btmCancela", Buttons.CancelCaption
  Table.EndColumn
  Table.EndRow

  Table.EndTable

  Form.EndForm

  Table.ColumnNoWrap = False

  sobjShow.CenterEnd

  FPLPLink

End Sub
Rem -------------------------------------------------------------------------
Rem Final da Sub FistPage

Rem -------------------------------------------------------------------------
Rem Executa
Rem -------------------------------------------------------------------------
Private Sub ExeFindFirstPage()
  On Error GoTo ERROR_HANDLER
  ' (Monta o WHERE, a partir dos dados entrados na tela de localização, que será passado como parametro
  ' pela QueryString.
  ' Uma estrutura de localização é montada utilizando o objeto FindFields[do HiperTools] e o objeto Collection
  ' , e a partir desta estrutura é feito o WHERE
  ' Alguns tipos de operador permitem a localização por multiplos valores, para um mesmo campo, através
  ' de valores separados por "ou" ou "," no caso dos operadores "igual a", "contendo o texto", "terminando
  ' com", "iniciando com" e separados por "e" ou "|" no caso do operador "entre")
  Dim strCampo, strValor, strValorAux, strOperador As String
  Dim i, intPos, intLess As Integer
  Dim OrGroup As New Collection
  Dim AllOrs As New Collection

  ' (Testa se algum botão foi clicado)
  If Request.Form("btmLocaliza") > "" Then
  ' (Testa se a estrutura de localização está montada)
    If IsObject(Session("FindObject")) Then
      Set AllOrs = Session("FindObject")

    End If
    For i = 1 To Session("NumberOfFinds")
      ' (Para cada campo utilizado na localização, é testada a existência de um valor a ser localizado e sua
      ' validade, e são armazenados campo, operador e valores)
      strCampo = Request.Form("txtCampo" & i)
      FindField CurrentTable, strCampo
      strOperador = Request.Form("txtOperador" & i)
      If Session("LastOption") = efQSActionCommonFind Then
        strValor = Request.Form("txtValor")
        If (Request.Form("strIgual") = "1") Or (Session(CurrentField & "Type") <> efDataTypeVarChar And Session(CurrentField & "Type") <> efDataTypeChar) Then
          strOperador = "1"

        Else
          strOperador = "2"

        End If
      Else
        strValor = Request.Form(Session(CurrentField & "Field"))

      End If

      If strValor > "" Then
        ' (Monta a parte da estrutura de localização referente aos valores)
        Do While strValor > ""
          If strOperador <> "9" Then
            intPos = InStr(LCase(strValor), ",")
            If intPos = 0 Then
              intPos = InStr(LCase(strValor), " ou ")
              intLess = 3

            Else
              intLess = 0

            End If

          Else
            intPos = InStr(LCase(strValor), "-")
            If intPos = 0 Then
              intPos = InStr(LCase(strValor), " e ")
              intLess = 2

            Else
              intLess = 0

            End If
          End If

          If intPos > 0 Then
             strValorAux = Left(strValor, intPos - 1)
             strValor = Right(strValor, Len(strValor) - intPos - intLess)

          Else
            strValorAux = strValor
            strValor = ""

          End If
          FindFields.Value.Add FormatForFind(strValorAux)

        Loop

        ' (Seta o nome do campo e o operador)
        FindFields.Field = Session(CurrentField & "Field")
        FindFields.Operator = strOperador

        ' (Monta a parte da estrutura de localização referente aos valores)
        OrGroup.Add FindFields
      End If
    Next
    ' (Se houver algo na estrutura de localização é testado qual o botão clicado, para montar uma nova tela
    ' de consulta ou realizar a consulta, senão volta para a última tela anterior à tela localização)
    If OrGroup.Count > 0 Then
      AllOrs.Add OrGroup
      Set Session("FindObject") = AllOrs

      Select Case Request.Form("btmLocaliza")
        Case "Ou"
          Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & Session("LastOption") & WhatList & GetEditable

        Case "Edição"
          Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionEditor & efQueryStrEditableStr & WhatList & MakeFindFilter

        Case "Listagem"
          Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionList & efQueryStrEditableStr & WhatList & MakeFindFilter

        Case "Relatório"
          Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionSummary & efQueryStrEditableStr & WhatList & "&" & efQueryStrList & "=" & Request.Form("selLists") & MakeFindFilter

      End Select
    Else
      Select Case Request.Form("btmLocaliza")
        Case "Edição"
          Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionEditor & efQueryStrEditableStr & WhatList

        Case "Listagem"
          Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionList & efQueryStrEditableStr & WhatList

        Case "Relatório"
          Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionSummary & efQueryStrEditableStr & WhatList & "&" & efQueryStrList & "=" & Request.Form("selLists")

        Case Else
          Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & Session("LastOption") & WhatList & GetEditable

      End Select
    End If
  Else
    Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & Session("LastOption") & WhatList & GetEditable

  End If

  If Request.Form("btmCancela") > "" Then
    'Alterado (Ridai Govinda)
    Response.Redirect QSLast
    '-----------------------
  End If

  'Alterado (Ridai Govinda)
  Session("NumberOfFinds") = 0
  Session("LastOption") = ""
  '-----------------------

ERROR_HANDLER_END:
  On Error Resume Next

  Err.Clear

  If sstrErrorMessage > "" Then
    ShowError efErrorOtherPart, "EditForm", "ExeFindFirstPage", sstrErrorMessage
    
  End If
  
Exit Sub

ERROR_HANDLER:
  sstrErrorMessage = "Err.Number: " & Err.number & "<BR>Err.Description: " & Err.Description

  GoTo ERROR_HANDLER_END

End Sub

Private Function MakeFindFilter()
  On Error GoTo ERROR_HANDLER
  ' (Auxilia o ExeFindFirstPage criando o WHERE da localização)
  Dim AllOrs As New Collection
  Dim i, j, k As Integer
  Dim strFind As String

  ' (Testa se a estrutura de localização está montada)
  If IsObject(Session("FindObject")) Then
    Set AllOrs = Session("FindObject")

  End If
  ' (Para cada "OU" da localização monta o trecho de WHERE refente a este)
  For i = 1 To AllOrs.Count
    If i > 1 Then
      strFind = strFind & " OR "

    End If
    strFind = strFind & " ("
    ' (Para cada campo da localização monta o trecho de WHERE refente a este)
    For j = 1 To AllOrs(i).Count
      If j > 1 Then
        strFind = strFind & " AND "

      End If
      FindField CurrentTable, AllOrs(i).Item(j).Field
      If AllOrs(i).Item(j).Value.Count = 1 Then
        strFind = strFind & AllOrs(i).Item(j).Field
        strFind = strFind & ChangeOperators(AllOrs(i).Item(j).Operator)
        strFind = strFind & FormatForSQL(FormatForWhere(AllOrs(i).Item(j).Value(1), AllOrs(i).Item(j).Operator), Session(CurrentField & "Type"))

      ElseIf AllOrs(i).Item(j).Value.Count > 1 Then
        Select Case AllOrs(i).Item(j).Operator
          ' (Operador "entre")
          Case "9"
            strFind = strFind & "(" & AllOrs(i).Item(j).Field
            strFind = strFind & ChangeOperators(AllOrs(i).Item(j).Operator)
            strFind = strFind & FormatForSQL(FormatForWhere(AllOrs(i).Item(j).Value(1), AllOrs(i).Item(j).Operator), Session(CurrentField & "Type"))
            strFind = strFind & " AND "
            strFind = strFind & FormatForSQL(FormatForWhere(AllOrs(i).Item(j).Value(2), AllOrs(i).Item(j).Operator), Session(CurrentField & "Type")) & ")"

          ' (Operador "igual")
          Case "1"
            strFind = strFind & AllOrs(i).Item(j).Field
            ' (Monta o "IN" com os valores)
            strFind = strFind & " IN ("
            For k = 1 To AllOrs(i).Item(j).Value.Count
              If k > 1 Then
                strFind = strFind & ", "

              End If
              strFind = strFind & FormatForSQL(FormatForWhere(AllOrs(i).Item(j).Value(k), AllOrs(i).Item(j).Operator), Session(CurrentField & "Type"))

            Next
            strFind = strFind & ")"

          ' (Outros operadores)
          Case Else
            strFind = strFind & "("
            ' (Para cada valor de um campo da localização monta o trecho de WHERE refente a este)
            For k = 1 To AllOrs(i).Item(j).Value.Count
              If k > 1 Then
                strFind = strFind & " OR "

              End If
              strFind = strFind & AllOrs(i).Item(j).Field
              strFind = strFind & ChangeOperators(AllOrs(i).Item(j).Operator)
              strFind = strFind & FormatForSQL(FormatForWhere(AllOrs(i).Item(j).Value(k), AllOrs(i).Item(j).Operator), Session(CurrentField & "Type"))

            Next
          strFind = strFind & ")"
        End Select
      End If
    Next
    strFind = strFind & ") "
  Next
  Set Session("FindObject") = Nothing
  MakeFindFilter = "&" & efQueryStrFilter & "=" & Server.URLEncode(strFind)

ERROR_HANDLER_END:
  On Error Resume Next

  Err.Clear

  If sstrErrorMessage > "" Then
    ShowError efErrorOtherPart, "EditForm", "MakeFindFilter", sstrErrorMessage
    
  End If
  
Exit Function

ERROR_HANDLER:
  sstrErrorMessage = "Err.Number: " & Err.number & "<BR>Err.Description: " & Err.Description

  GoTo ERROR_HANDLER_END

End Function

Private Sub MakeFindOr()
' (Auxilia o FindFirstPage criando a mensagem da localização escolhida pelo usuário, até o momento.
' Esta sub repete o processo da função anterior mudando apenas a string de resultado)
Dim AllOrs As New Collection
Dim i, j, k As Integer
Dim strFind As String

' (Testa se a estrutura de localização está montada)
  If IsObject(Session("FindObject")) Then
    Set AllOrs = Session("FindObject")

  End If
  If AllOrs.Count > 0 Then
' (Para cada "OU" da localização monta o trecho da mensagem refente a este)
    For i = 1 To AllOrs.Count
      If i > 1 Then
        strFind = strFind & LanguageLocalizer(22) & "<BR>"

      End If
' (Para cada campo da localização monta o trecho da mensagem refente a este)
      For j = 1 To AllOrs(i).Count
        If j > 1 Then
          strFind = strFind & LanguageLocalizer(23)

        End If
        FindField CurrentTable, AllOrs(i).Item(j).Field
' (Testa se para o campo existe um ou mais valores a serem localizados)
        If AllOrs(i).Item(j).Value.Count = 1 Then
          strFind = strFind & EditCaption(False)
          strFind = strFind & " " & TextualOperators(AllOrs(i).Item(j).Operator) & " "
          strFind = strFind & AllOrs(i).Item(j).Value(1)

' (Se para o campo existe mais de um valor a ser localizado)
        ElseIf AllOrs(i).Item(j).Value.Count > 1 Then
          Select Case AllOrs(i).Item(j).Operator
' (Operador "entre")
            Case "9"
              strFind = strFind & EditCaption(False)
              strFind = strFind & " " & TextualOperators(AllOrs(i).Item(j).Operator) & " "
              strFind = strFind & AllOrs(i).Item(j).Value(1)
              strFind = strFind & LanguageLocalizer(23)
              strFind = strFind & AllOrs(i).Item(j).Value(2)

' (Outros operadores)
            Case Else
              strFind = strFind & EditCaption(False)
              strFind = strFind & " " & TextualOperators(AllOrs(i).Item(j).Operator) & " "
' (Para cada valor de um campo da localização monta o trecho da mensagem refente a este)
              For k = 1 To AllOrs(i).Item(j).Value.Count
                If k > 1 Then
                  strFind = strFind & LanguageLocalizer(22)

                End If
                strFind = strFind & AllOrs(i).Item(j).Value(k)

              Next
          End Select
        End If
      Next
    Next
' (Apresenta a mensagem numa tabela)
    Table.BeginTable "85%"
    Table.BeginRow 2
    Table.BeginColumn
    sobjShow.Bold
    Response.Write strFind
    sobjShow.BoldEnd
    Table.EndColumn
    Table.EndRow
    Table.EndTable
    sobjShow.BR
  End If

End Sub

Private Function FormatForFind(ByVal strValor As String)
' (Faz a formatação dos valores dos campos para a localização, parametro: valor)
  Select Case Session(CurrentField & "EditOption")
    Case efFldTypeCheck
      If LCase(strValor) = LCase(sobjString.ToString(Session(CurrentField & ",1Value"))) Then
        FormatForFind = Session(CurrentField & ",1Value")

      Else
        FormatForFind = Session(CurrentField & ",2Value")

      End If

    Case Else
        FormatForFind = FormatField(strValor, Session(CurrentField & "Type"), Session(CurrentField & "Size"), Session(CurrentField & "CharCase"))

  End Select

End Function

Private Function FormatForWhere(ByVal strValor As String, ByVal strOperador As String) As String
' (Faz a formatação dos valores do tipo string para a localização dependendo do operador,
' parametros: valor, operador)
  Select Case strOperador
    Case "2"
      FormatForWhere = "%" & strValor & "%"

    Case "3"
      FormatForWhere = strValor & "%"

    Case "4"
      FormatForWhere = "%" & strValor

    Case Else
      FormatForWhere = strValor

  End Select

End Function

Private Function ChangeOperators(ByVal strOperador As String) As String
' (Retorna o operador correspondente ao nº passado nos inputs da localização, parametro: operador)
  Select Case strOperador
    Case "1"
      ChangeOperators = " = "

    Case "2", "3", "4"
      ChangeOperators = " like "

    Case "5"
      ChangeOperators = " > "

    Case "6"
      ChangeOperators = " < "

    Case "7"
      ChangeOperators = " >= "

    Case "8"
      ChangeOperators = " <= "

    Case "9"
      ChangeOperators = " Between "

  End Select

End Function

Private Function TextualOperators(ByVal strOperador As String) As String
' (Retorna o texto do operador correspondente ao nº passado nos inputs da localização, parametro:
' operador)
  Select Case strOperador
    Case "1"
      TextualOperators = LanguageLocalizer(24)

    Case "2"
      TextualOperators = LanguageLocalizer(25)

    Case "3"
      TextualOperators = LanguageLocalizer(26)

    Case "4"
      TextualOperators = LanguageLocalizer(27)

    Case "5"
      TextualOperators = LanguageLocalizer(28)

    Case "6"
      TextualOperators = LanguageLocalizer(29)

    Case "7"
      TextualOperators = LanguageLocalizer(30)

    Case "8"
      TextualOperators = LanguageLocalizer(31)

    Case "9"
      TextualOperators = LanguageLocalizer(32)

  End Select

End Function

Private Function EditCaption(ByVal blnForm As Boolean) As String
' (Os campos do formulário de edição podem ser acessados através do atalho <Alt> + <[Letra]>,
' bastando para isto que o programador coloque no rótulo do campo o caractere "&" antes da letra do
' atalho. Esta função retorna o rótulo, colocando sublinhado nos caracteres que seguem "&" se a tela for
' de edição)
Dim i As Integer
  i = InStr(Session(CurrentField & "Caption"), "&")

  If i > 0 Then
    If blnForm Then
      EditCaption = FormatShowFieldCaption(Left(Session(CurrentField & "Caption"), i - 1) & "<U>" & Mid(Session(CurrentField & "Caption"), i + 1, 1) & "</U>" & Mid(Session(CurrentField & "Caption"), i + 2))
      Form.AccessKey = sobjString.ChangeChar(Mid(Session(CurrentField & "Caption"), i + 1, 1))

    Else
      EditCaption = Left(Session(CurrentField & "Caption"), i - 1) & Mid(Session(CurrentField & "Caption"), i + 1)

    End If

  Else
    If blnForm Then
      EditCaption = FormatShowFieldCaption(Session(CurrentField & "Caption"))

    Else
      EditCaption = Session(CurrentField & "Caption")

    End If
  End If

End Function

Private Function FormatVariable(ByVal Value As Variant, ByVal strType As String) As String
' (Transforma um valor para o tipo desejado, parametros: valor, tipo)
  If Value > "" Then
    If LCase(TypeName(Value)) <> "string" Then
      Select Case strType
        Case efDataTypeInt, efDataTypeSmallInt, efDataTypeTinyInt
          FormatVariable = sobjString.ZTIInt(Value)

        Case efDataTypeMoney
          FormatVariable = sobjString.ZTIMoney(Value)

        Case efDataTypeFloat, efDataTypeReal
          FormatVariable = sobjString.ZTINumber(Value)

        Case efDataTypeDateTime
          If LCase(TypeName(Value)) = "date" Then
            FormatVariable = sobjString.ZTILongDate(Value)

          Else
            If Validate.IsValidDate(Value) Then
              FormatVariable = Validate.FormatDate(Value)

            Else
              FormatVariable = Value

            End If
          End If

        Case Else
          FormatVariable = Value

      End Select

    Else
      FormatVariable = Trim(Value)

    End If
  Else
    FormatVariable = ""

  End If

End Function

Private Function FormatForSQL(ByVal Value, ByVal strType As String) As String
' (Transforma um valor a ser utilizado num comando SQL, parametros: valor, tipo)
  Select Case strType
    Case efDataTypeVarChar, efDataTypeChar, efDataTypeText
      If Value > "" Then
        FormatForSQL = sobjString.SQLString(Value)

      Else
        FormatForSQL = "''"

      End If

    Case efDataTypeDateTime
        FormatForSQL = sobjString.SQLDateTime(Value)

    Case Else
      FormatForSQL = sobjString.ToString(Value)

    End Select

End Function

Private Function ImportDouble(ByVal strValue As String) As Double
' (Transforma uma string para double)
Dim Target As String
Dim i As Integer
Dim cOneChar As String
Dim intPos As Integer
Dim monAux As Double

  If Trim(strValue) > "" Then
    Target = ""
    strValue = UCase(Trim(strValue))

' (Retira os caracteres diferentes de "-", "," e "0"-"9" )
    For i = 1 To Len(strValue)
      cOneChar = Mid(strValue, i, 1)
      If cOneChar = "," Or cOneChar = "-" Or (cOneChar >= "0" And cOneChar <= "9") Then
         Target = Target & cOneChar

      End If
    Next

    strValue = Target

    intPos = InStr(strValue, ",")

    monAux = 0

' (Se existe valor decimal então a parte inteira corresponde ao que está antes da vírgula e a parte
' fracionária o que está depois)
    If intPos > 0 Then
      monAux = CLng(Left(strValue, intPos - 1))
' (A parte fracionária)
      monAux = monAux + (CLng(Mid(strValue, intPos + 1)) / Exp((Len(strValue) - intPos) * Log(10)))

    Else
      monAux = CLng(strValue)

    End If

    ImportDouble = monAux

 Else
    ImportDouble = 0

 End If

End Function

Private Function FormatField(ByVal Value As String, ByVal strType As String, ByVal intSize As Long, ByVal intCharCase As Integer) As Variant
' (Formata um variável para apresentação conforme seu tipo, parametros: valor, tipo, tamanho[se for string], case[maiúsculas ou minúsculas])
  If Value > "" Then
    Select Case LCase(strType)
      Case efDataTypeInt, efDataTypeSmallInt, efDataTypeTinyInt
        If LCase(TypeName(Value)) = "string" Then
          FormatField = CLng(sobjString.TotalNormalizeString(Value))

        Else
          FormatField = Value

        End If

      Case efDataTypeDateTime
        If TypeName(Value) = "date" Then
          FormatField = sobjString.ZTILongDate(Value)

        Else
          If Validate.IsValidDate(Value) Then
            Value = Validate.FormatDate(Value)

          End If
        End If
        FormatField = Value

      Case efDataTypeVarChar, efDataTypeChar, efDataTypeText
' (Aplica o case)
        Select Case intCharCase
          Case efCharCaseLower
            FormatField = LCase(Left(Trim(sobjString.ToString(Value)), intSize))

          Case efCharCaseUpper
            FormatField = UCase(Left(Trim(sobjString.ToString(Value)), intSize))

          Case Else
            FormatField = Left(Trim(sobjString.ToString(Value)), intSize)

        End Select

      Case efDataTypeMoney, efDataTypeFloat, efDataTypeReal
        Select Case LCase(TypeName(Value))
          Case "currency", "double", "single", "decimal"
            FormatField = Value

          Case Else
            FormatField = ImportDouble(Value)

        End Select

      Case Else
        FormatField = Value

      End Select

  Else
    FormatField = Null

  End If
End Function

Private Function FormatFieldtoSave(ByVal Value As Variant, ByVal strType As String, ByVal intSize As Long, _
    ByVal intCharCase As Integer) As Variant
' (Formata o valor de um campo para salvar no BD, parametros: valor, tamanho, case[maiúsculas ou minúsculas])
  If Value > "" And Not IsNull(Value) Then
    Select Case LCase(strType)
      Case efDataTypeBit
        If Value Then
          FormatFieldtoSave = 1

        Else
          FormatFieldtoSave = 0

        End If

      Case efDataTypeInt, efDataTypeSmallInt, efDataTypeTinyInt
        If LCase(TypeName(Value)) = "string" Then
          FormatFieldtoSave = CLng(sobjString.NormalizeNumber(Value, True))

        Else
          FormatFieldtoSave = Value

        End If

      Case efDataTypeDateTime
        If LCase(TypeName(Value)) = "date" Then
          FormatFieldtoSave = Value

        Else
          FormatFieldtoSave = sobjString.ImportDate(Value, stDateTypeDDMMAAAA)

        End If

      Case efDataTypeVarChar, efDataTypeChar, efDataTypeText
' (Aplica o case)
        Select Case intCharCase
          Case efCharCaseLower
            FormatFieldtoSave = LCase(Left(sobjString.ToString(Value), intSize))

          Case efCharCaseUpper
            FormatFieldtoSave = UCase(Left(sobjString.ToString(Value), intSize))

          Case Else
            FormatFieldtoSave = Left(sobjString.ToString(Value), intSize)

        End Select

      Case efDataTypeMoney, efDataTypeFloat, efDataTypeReal
        If LCase(TypeName(Value)) = "string" Then
          FormatFieldtoSave = sobjString.ImportDouble(Value)

        Else
          FormatFieldtoSave = Value

        End If

      Case Else
        FormatFieldtoSave = Value

      End Select

  Else
    FormatFieldtoSave = Null

  End If
End Function

Private Sub FindFirst(ByVal RS As ADODB.RecordSet, ByVal strField As String, ByVal strText As String)
' (Na listagem existe uma opção chamada AlphabeticIndex que permite ao usuário clicar numa letra
' encontrando o primeiro item da listagem iniciado com esta)
Dim Inicio, Fim, Last As Integer
Dim strRSText As String

' (Popula o RecordCount do RecordSet)
  RS.MoveLast
  RS.MoveFirst

  sblnFind = False
  sintPosition = 0
  Inicio = 0
  Fim = RS.RecordCount - 1
' (Faz uma pesquisa binária pela letra)
  Do While Inicio <= Fim And Not sblnFind
    Last = sintPosition
    sintPosition = (Inicio + Fim) \ 2
    RS.Move sintPosition - Last
    strRSText = LCase(sobjString.ChangeChar(Left(RS(strField), 1)))
    If strRSText > strText Then
      Fim = sintPosition - 1

    Else
      If strRSText < strText Then
        Inicio = sintPosition + 1

      Else
        sblnFind = True

      End If
    End If
  Loop
  If sblnFind Then
' (Procura o primeiro item com esta letra, numa pesquisa sequêncial inversa)
    Do While strRSText = strText And Not RS.BOF
      strRSText = LCase(sobjString.ChangeChar(Left(RS(strField), 1)))
      sintPosition = sintPosition - 1
      RS.MovePrevious

    Loop
  End If
  sintPosition = sintPosition + 2

End Sub

Private Sub MarkLines(ByRef pacMarks As HiperTools30.Pack, ByVal intLines As Integer)
' (Filtra os registros que obedecem ao WHERE das marcas de linha)
Dim i, j, k As Integer
Dim strAux, sql As String
Dim blnFirst As Boolean

' (Filtro de cada marca)
  For j = 1 To Session(CurrentTable & "MarkNumber")
    sql = "SELECT "
    blnFirst = True
    For i = 1 To Session(CurrentTable & "FieldNumber")
      If Session(CurrentTable & i & "IsKey") Then
        If Not blnFirst Then
          sql = sql & ", "

        Else
          blnFirst = False

        End If
        sql = sql & Session(CurrentTable & i & "Field")

      End If
    Next
    sql = sql & " FROM " & CurrentTable & " WHERE (("

    Database.OrState = False

    For k = 1 To intLines
      blnFirst = True
      For i = 1 To Session(CurrentTable & "FieldNumber")
        If Session(CurrentTable & i & "IsKey") Then
          If Not blnFirst Then
            sql = sql & " AND "

          Else
            blnFirst = False

          End If
          sql = sql & Session(CurrentTable & i & "Field") & "=" & FormatForSQL(Session("Listagem" & k & "," & i & "Value"), Session(CurrentTable & i & "Type"))

        End If

      Next
      If k < intLines Then
        sql = sql & ") OR ("

      End If
    Next
    sql = sql & ")) AND (" & Session(CurrentTable & j & "MarkWhere") & ")"

    Rem sobjCMD2.CommandText = sql
    Rem sobjCMD2.CommandType = adCmdText

    'Alterado (Ridai)
    Set CurrConnection = sobjConn2
    If ShowSQLQuery Then ShowCurrentSQL sql, "MarkLines", adOpenKeyset, adLockReadOnly

    On Error GoTo SQLError

    sobjRS2.Open sql, sobjConn2, adOpenKeyset, adLockReadOnly

    On Error GoTo 0
    '----

    If Not sobjRS2.EOF And Not sobjRS2.BOF Then
' (Seleciona os registros que possuem esta marca)
      Do While Not sobjRS2.EOF
        strAux = ""
        For i = 1 To Session(CurrentTable & "FieldNumber")
          If Session(CurrentTable & i & "IsKey") Then
            strAux = strAux & sobjRS2(Session(CurrentTable & i & "Field")) & ";"

          End If
        Next
        strAux = strAux & Session(CurrentTable & j & "MarkField")

        'Alterado/Corrigido (Ridai)
        If pacMarks.Find(strAux) Then pacMarks.Remove strAux
        '---
' (Coloca estes registros num Objeto Pack juntamente com o nº da marca)
        pacMarks.Add strAux, j
        sobjRS2.MoveNext
      Loop
    End If
    sobjRS2.Close
  Next

  'Criado (Ridai Govinda)
  'Rotina de tratamento de Erro (consultas SQL)
  Exit Sub
SQLError:
  If Err.number <> 0 Then
    If DebugMode Then
      ShowSQLErrors sql, "MarkLines", adOpenKeyset, adLockReadOnly
    Else
      ShowDefaultError
    End If
    Err.Clear

  End If
End Sub

Private Sub ShowList()
  Const lconArrFieldsMax = 100
' (A listagem é montada a partir de dois SELECTs, um sem JOINs que só busca os campos chave dos
' registros selecionados e paginados, utilizados para buscar na outro SELECT somente os registros
' desejados para a busca com JOINs.
' Esta sub apresenta a listagem)
  Dim i As Integer, j As Integer, k As Integer, l As Integer
  Dim arrFields(lconArrFieldsMax) As Integer, intCount As Integer
  Dim sql As String, strAux As String, strAux2 As String, strTip As String
  Dim strTip2 As String, strOrder As String, strOrder2 As String, strJoin As String
  Dim blnFirst As Boolean, blnOneField As Boolean
  Dim pacMarks As HiperTools30.Pack
  'Criado (Ridai Govinda)
  Dim intJCounter As Integer
  Dim strLookUpTable As String
  Dim strEditable As String
'
  intJCounter = 0
  strLookUpTable = ""
  '-----------

  Browse.Actual = CurrentTable

' (Se não há uma ordenação atual então coloca o campo da primeira coluna como ordenador)
  If Not Session("OrderField") > "" Then
    BeginReadOperation Session(CurrentTable & "FieldList")

    Do While IsThereOperation
      If FindField(CurrentTable, NextOperation) Then
        sstrText = ""
        Session("OrderField") = Session("CurrentField")

      End If

    Loop

  End If

' (Faz os JOIN dos campos da ordenação, se estes forem de outra tabela e o ORDER BY)
  BeginReadOperation Session("OrderField")
  strOrder = " ORDER BY"

  blnFirst = True
  blnOneField = AlphabeticIndex

  strJoin = ""

  Do While IsThereOperation
    Select Case NextOperation
       Case ",", ";", "-"

       Case Else
         Session("CurrentField") = LastOperation
         If Not blnFirst Then
           strOrder = strOrder & ", "
           blnOneField = False

         Else
           strOrder = strOrder & " "

         End If

         Select Case Session(CurrentField & "EditOption")
           Case efFldTypeLookup
              strOrder = strOrder & Session(CurrentField & "LookupTable") & "." & Session(CurrentField & "LookupFieldText")

              'Alterado (Kelson/Ridai Govinda)
              If intJCounter > 0 Then strJoin = "( " & strJoin & " )"

              strLookUpTable = Session(CurrentTable & Session("CurrentField") & "LookupTable")

              If LCase(Trim(CurrentTable)) = LCase(Trim(strLookUpTable)) Then
              strLookUpTable = strLookUpTable & "_" & intJCounter
              End If

              strJoin = strJoin & " LEFT JOIN " & Session(CurrentTable & Session("CurrentField") & "LookupTable")
              If LCase(Trim(Session(CurrentTable & Session("CurrentField") & "LookupTable"))) <> LCase(Trim(strLookUpTable)) Then
              strJoin = strJoin & " AS " & strLookUpTable
              End If

              strJoin = strJoin & " ON " & OwnerTable(CurrentTable, Session("CurrentField")) & "." & Session(CurrentTable & Session("CurrentField") & "Field") & " = " & _
              strLookUpTable & "." & Session(CurrentTable & Session("CurrentField") & "LookupField")

              intJCounter = intJCounter + 1
              '-----------------------------

           Case Else
             If Not IsExternalField Then
               'Alterado (Ridai Govinda)
               strOrder = strOrder & CurrentTable & "." & Session(CurrentField & "Field")
               '-----------------------

             Else
               strOrder = strOrder & OwnerTable(CurrentTable, Session("CurrentField")) & "." & Session(CurrentField & "Field")
               strJoin = strJoin & MakeJoin(CurrentTable, Session(CurrentField & "OwnerTable"), strJoin)

             End If

        End Select

        blnFirst = False
        If Session("OrderDesc") = "1" Then
          strOrder = strOrder & " DESC "

        End If

    End Select

  Loop

' (Monta um SELECT somente com os JOINs relacionados à ordenação)
  'Alterado (Ridai Govinda)
  Dim strKeyField As String

  blnFirst = True
  strKeyField = ""
  sql = "SELECT "

  If blnOneField Then
    strKeyField = strKeyField & Session(CurrentTable & Session("OrderField") & "Field")

  End If

  For i = 1 To Session(CurrentTable & "FieldNumber")
    If Session(CurrentTable & i & "IsKey") And (Not blnOneField Or i <> Session("OrderField")) Then
      If Not blnFirst Or blnOneField Then
        strKeyField = strKeyField & ", "

      Else
        blnFirst = False

      End If

      strKeyField = strKeyField & Session(CurrentTable & i & "Field")

    End If
  Next

  If Not strKeyField > "" Then strKeyField = Session(CurrentTable & "1Field")

  sql = sql & strKeyField & " FROM " & CurrentTable & strJoin & Filter & strOrder
  Select Case ConnectionType
  Case efConnDBDriverOracle, efConnDBProviderOracle
    sql = FixJoinForOracle(sql)
    If ConnectionType = efConnDBProviderOracle Then sobjRS.CursorLocation = adUseClient

  End Select

  Set CurrConnection = sobjConn

  If ShowSQLQuery Then ShowCurrentSQL sql, "ShowList", adOpenKeyset, adLockReadOnly

  On Error Resume Next

  sobjRS.Open sql, sobjConn, adOpenKeyset, adLockReadOnly

  If Err.number <> 0 Then
    If DebugMode Then
      ShowSQLErrors sql & strOrder, "ShowList", adOpenKeyset, adLockReadOnly
    Else
      ShowDefaultError
    End If
    Err.Clear
    Exit Sub

  End If
  On Error GoTo 0
  '----

Rem  sobjCMD.CommandType = adCmdText
Rem  sobjCMD.CommandText = sql & strOrder
Rem  sobjRS.Open sobjCMD, , adOpenKeyset, adLockReadOnly

  j = 0
  Dim blnHaveFind As Boolean
  blnHaveFind = blnOneField And sparEditFind > "" And Not Request.QueryString("BA") > ""
  sblnFind = False
  If Not sobjRS.EOF And Not sobjRS.BOF Then
    Browse.BeginBrowse sobjRS
    If blnHaveFind Then
      FindFirst sobjRS, Session(CurrentTable & Session("OrderField") & "Field"), sparEditFind

      'Alterado (Ridai)
      On Error Resume Next

      sobjRS.AbsolutePage = ((sintPosition - 1) \ Browse.PageSize) + 1

      If Err.number <> 0 Then
        If DebugMode Then
          DebugMessage LanguageLocalizer(33), 1
        Else
          ShowDefaultError
        End If

        Err.Clear
        Exit Sub

      End If
      On Error GoTo 0
      '---
      Browse.Page = ((sintPosition - 1) \ Browse.PageSize) + 1

    End If

' (Guarda os valores das chaves dos registros selecionados e paginados)
    Do While Not sobjRS.EOF And Browse.IsBrowsePage
      j = j + 1
      For i = 1 To Session(CurrentTable & "FieldNumber")
        If Session(CurrentTable & i & "IsKey") Then
          Session("Listagem" & j & "," & i & "Value") = sobjRS(Session(CurrentTable & i & "Field"))

        End If

      Next
      sobjRS.MoveNext
      Browse.IncLine

    Loop

    Browse.Status = True
    Browse.EndBrowse

' (Monta as marcas de linha)
    Set pacMarks = Server.CreateObject("HiperTools30.Pack")
    MarkLines pacMarks, j

' (Monta o SELECT da listagem)
    'Alterado/Removido (Ridai Govinda
    '----------------------
    Rem strJoin = CurrentTable
    strJoin = GetJoinForList

    For i = 1 To Session(CurrentTable & "FieldNumber")
      If Session(CurrentTable & i & "EditOption") = efFldTypeLookup Then


      End If
    Next


' (Monta a lista dos nº dos campos da listagem para acelerar sua busca na inserção do conteúdo na
' tabela e contabiliza a largura da tabela)
      BeginReadOperation Session(CurrentTable & "FieldList")

      Dim Width, strList

      Width = 0
      Do While IsThereOperation
        Select Case NextOperation
          Case ",", ";", "-"
          Case "="
            IsThereOperation

          Case Else
            If FindField(CurrentTable, LastOperation) Then
              If Session(CurrentField & "Type") <> "varchar" Then
                Width = Width + EditSize + 8

              Else
                Width = Width + EditSize

              End If
              If strList > "" Then
                strList = strList & ","

              End If
              strList = strList & Session("CurrentField")

            End If

        End Select
      Loop

      If Width < 60 Then
        Width = 60

      ElseIf Width > 95 Then
        Width = 95

      End If

    'Criado (Ridai Govinda)
    sql = "SELECT " & GetFieldsForList
    sql = sql & " FROM "
    sql = sql & FixExtRelationJoin(strJoin, Session(CurrentTable & "JoinFormUnit"))
    '---------------------

    Database.OrState = False

' (Monta o where a partir das chaves antes selecionadas)
    For k = 1 To j
      blnFirst = True
      For i = 1 To Session(CurrentTable & "FieldNumber")
        If Session(CurrentTable & i & "IsKey") Then
          If Not blnFirst Then
            'Alterado (Ridai Govinda)
            sql = sql & " AND " & CurrentTable & "." & Session(CurrentTable & i & "Field") & "=" & FormatForSQL(Session("Listagem" & k & "," & i & "Value"), Session(CurrentTable & i & "Type"))
            '----------------------

          Else
            blnFirst = False
            sql = sql & Database.PutOr & "(" & CurrentTable & "." & Session(CurrentTable & i & "Field") & "=" & FormatForSQL(Session("Listagem" & k & "," & i & "Value"), Session(CurrentTable & i & "Type"))

          End If
'          sql = sql & Database.PutOr & "(" & Session(CurrentTable & i & "Field") & "=" & FormatForSQL(Session("Listagem" & k & "," & i & "Value"), Session(CurrentTable & i & "Type"))

        End If
      Next
      If InStr(1, sql, "WHERE", vbTextCompare) > 0 Then sql = sql & ")"
    Next

    sobjShow.Nbsp

    strOrder2 = ""
    'Alterado (Ridai)
    If Session(CurrentTable & "OrderBy") > "" Then
      strOrder2 = " ORDER BY " & Session(CurrentTable & "OrderBy")
      Rem strOrder2 = strOrder2 & ", " & Replace(strOrder, " ORDER BY ", vbNullString, , 1)

      sql = sql & FixListOrder(Replace(strOrder, " ORDER BY ", vbNullString, , 1), strOrder2)
    Else
      sql = sql & strOrder
    End If

    Select Case ConnectionType
    Case efConnDBDriverOracle, efConnDBProviderOracle
      sql = FixJoinForOracle(sql)

    End Select

    Rem sobjCMD2.CommandType = adCmdText
    Set CurrConnection = sobjConn2

    If ShowSQLQuery Then ShowCurrentSQL sql, "ShowList", adOpenKeyset, adLockReadOnly

    On Error Resume Next

    sobjRS2.Open sql, sobjConn2, adOpenKeyset, adLockReadOnly

    If Err.number <> 0 Then
      If DebugMode Then
        ShowSQLErrors sql, "ShowList", adOpenKeyset, adLockReadOnly
      Else
        ShowDefaultError
      End If
      Err.Clear
      Exit Sub

    End If
    On Error GoTo 0
    '----

' (Testa a existência de registros)
    If Not sobjRS2.EOF And Not sobjRS2.BOF Then
      Dim strMessage As String

      'Alterado (Ridai Govinda)
      If Not ListTitle > "" Then
        strMessage = LanguageLocalizer(68) & Session(CurrentTable & "Caption")
      Else
        strMessage = ListTitle
      End If
      '-----------------------

      sobjShow.Center

' (Aplica o estilo das fontes do cabecalho)
     FontStyles

' (Conta as colunas da tabela)
    Dim bytColNumber As Byte
      bytColNumber = Session(CurrentTable & "ListOperationNumber") + 1
      If blnHaveFind Then
        bytColNumber = bytColNumber + 1

      End If

      Table.ColumnNoWrap = False

' (Inicio da tabela)
      'Criado (Ridai Govinda)
      If AllowEnumFields Then EnumerateFields sobjRS2
      '---------------------

      Table.BeginTable Width & "%", strMessage, bytColNumber, True

' (Testa se é necessário criar a linha dos comandos "Incluir", "Editar", "Localizar", "Outras Ordenações" e "Outras Listagens")
      If ((Session(CurrentTable & "EditOption") > 0) And (Buttons.ShowAdd Or Buttons.ShowEdit)) Or _
         Buttons.ShowFind Or Session(CurrentTable & "ListNumber") > 0 Or Session(CurrentTable & "OrderNumber") > 0 Or blnOneField Then
' (Monta o Tip das "Outras Ordenações")
        If Session(CurrentTable & "OrderNumber") > 0 Then
          For i = 1 To Session(CurrentTable & "OrderNumber")
            strAux = ""
            strAux2 = ""
' (Troca o nome dos campos pelos seus rótulos)
            BeginReadOperation Session(CurrentTable & i & "Order")
            Do While IsThereOperation
              Select Case NextOperation
                Case ",", ";", "-", "="
                  strAux = strAux & LastOperation & " "

                Case Else
                  If FindField(CurrentTable, LastOperation) Then
                    strAux = strAux & EditCaption(False)
                    If strAux2 > "" Then
                      strAux2 = strAux2 & ","

                    End If
                    strAux2 = strAux2 & Session("CurrentField")

                  End If

              End Select
            Loop
' (Testa a relação entre a ordenação montada e a utilizada atualmente, e monta os links conforme a
' relação)
            If Session("OrderField") = strAux2 Then
              strAux2 = Server.URLEncode(strAux2)
              If Session("OrderDesc") = "1" Then
                'Alterado (Ridai Govinda)
                strTip = strTip & "<IMG SRC=""" & Session("htDefaultHiperToolsWorkPath") & "/ArrowDown.gif"" BORDER=0>" & _
                "<A CLASS=HL HRef=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionList & _
                WhereHow & GetFind & "&" & efQueryStrOrderField & "=" & strAux2 & GetFilter & """>" & _
                " " & strAux & "</A><BR>"
                '-----------------------

              Else
                'Alterado (Ridai Govinda)
                strTip = strTip & "<IMG SRC=""" & Session("htDefaultHiperToolsWorkPath") & "/ArrowUp.gif"" BORDER=0>" & _
                "<A CLASS=HL HRef=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionList & _
                WhereHow & GetFind & "&" & efQueryStrOrderField & "=" & strAux2 & efQueryStrOrderDescStr & GetFilter & """>" & _
                 " " & strAux & "</A><BR>"
                '----------------------

              End If
            Else
              strAux2 = Server.URLEncode(strAux2)
              'Alterado (Ridai Govinda)
              strTip = strTip & "&nbsp;&nbsp;&nbsp;<A CLASS=HL HRef=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & _
                                                  efQSActionList & WhereHow & GetFind & "&" & efQueryStrOrderField & "=" & strAux2 & GetFilter & """>" & strAux & "</A><BR>"
              '------------------------

            End If
          Next
        End If

' (Monta o Tip das "Outras Listagens")
        For i = 1 To Session(CurrentTable & "ListNumber")
          'Alterado/Criado (Ridai Govinda)
          strTip2 = strTip2 & "&nbsp;&nbsp;&nbsp;<A CLASS=SS HRef=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & _
          efQSActionSummary & WhereHow & GetFind & "&" & efQueryStrList & "=" & Session(CurrentTable & i & "ListName") & GetFilter & """>" & Session(CurrentTable & i & "ListCaption") & "</A><BR>"
          '-----------------------------

        Next

        'Alterado (Ridai Govinda)
        If ShowListUpperBar Then
          Table.ColumnColSpan = Session(CurrentTable & "ListOperationNumber") + 1

          Table.BeginRow 1
          Table.BeginColumn
          Table.Padding = 0

          Table.BeginTable "100%"
          Table.Padding = 3
          Table.ColumnAlign = "Left"
          Table.ColumnVAlign = "Top"
          Table.BeginRow 1
          Table.BeginColumn

          ' (Se estiver em modo de edição e os comandos[incluir e editar] estiverem
          '  abilitados cria seus links)
          If (Session(CurrentTable & "EditOption") > 0) And ShowButtons Then
            If Buttons.ShowAdd Then
              'Alterado (Ridai Govinda)
              strEditable = sparEditEditable
              sparEditEditable = 1
              Response.Write "<B><A CLASS=""SS"" HREF=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionAdd & _
                         WhereHow & """>" & Buttons.AddCaption & "</A></B>"
              sparEditEditable = strEditable
              '-----------------------

              sobjShow.Nbsp
              sobjShow.Nbsp

            End If
            If Buttons.ShowEdit Then
              'Alterado (Ridai Govinda)
              Response.Write "<B><A CLASS=""SS"" HREF=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionEditor & _
                         WhereHow & GetFilter & """>" & Buttons.EditCaption & "</a></B>"
              '-----------------------
              sobjShow.Nbsp
              sobjShow.Nbsp

            End If
          End If
          ' (Se o comando localizar estivere abilitado testa que tipo de
          '  localização deve ser feita e cria seu link)
          If Buttons.ShowFind And ShowButtons Then
            'Alterado (Ridai Govinda)
            If Not AllowCommonFind Then
              If Session(CurrentTable & "FieldFind") > "" Then
                Response.Write "<B><A CLASS=""SS"" HREF=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionAdvancedFind & _
                         WhereHow & """>" & Buttons.FindCaption & "</A></B>"

              Else
                Response.Write "<B><A CLASS=""SS"" HREF=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionCommonFind & _
                           WhereHow & """>" & Buttons.FindCaption & "</A></B>"

              End If
            Else
              Response.Write "<B><A CLASS=""SS"" HREF=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionCommonFind & _
                         WhereHow & """>" & Buttons.FindCaption & "</A></B>"

            End If
            '-----------------------

          End If
          Table.EndColumn

          ' (Aplica o Tip de "Outras Listagens")
          If Session(CurrentTable & "ListNumber") > 0 And ShowReportsLink Then
            Table.ColumnAlign = "Right"
            Table.ColumnVAlign = "Top"
            Table.BeginColumn
            Response.Write "<FONT SIZE=""2""><DIV ID=""F2"" CLASS=""SS"" onMouseOver=""makeCool(F2);"" onMouseOut=""makeNormal(F2);"" onClick=""OpenDesc(T2);"" STYLE=""cursor:hand""><B>" & ListReportsCaption & "</B></DIV>"
            Response.Write "<TABLE><TR><TD ALIGN=""JUSTIFY"">"
            Response.Write "<DIV ID=""T2"" CLASS=""dbody"" STYLE=""display:'none';""><B>"
            Response.Write strTip2
            Response.Write "</B></DIV></TD></TR></TABLE>"
            Table.EndColumn

          End If
          ' (Aplica o Tip de "Outras Ordenações")
          If Session(CurrentTable & "OrderNumber") > 0 And ShowSetOfOrdersLink Then
            Table.ColumnAlign = "Right"
            Table.ColumnVAlign = "Top"
            Table.BeginColumn
            Response.Write "<FONT SIZE=""2""><DIV ID=""F1"" CLASS=""SS"" onMouseOver=""makeCool(F1);"" onMouseOut=""makeNormal(F1);"" onClick=""OpenDesc(T1);"" STYLE=""cursor:hand""><B>" & ListSetOfOrdersCaption & "</B></DIV>"
            Response.Write "<TABLE><TR><TD ALIGN=""JUSTIFY"">"
            Response.Write "<DIV ID=""T1"" CLASS=""dbody"" STYLE=""display:'none';""><B>"
            Response.Write strTip
            Response.Write "</B></DIV></TD></TR></TABLE>"
            Table.EndColumn

          End If
          Table.EndRow
          Table.EndTable

          ' (Se houver um AlphabeticIndex com um campo do tipo string selecionado
          ' então monta o alfabeto de links)
          If blnOneField Then
            Session("CurrentField") = Session("OrderField")
            Select Case Session(CurrentField & "EditOption")
              Case efFldTypeImage, efFldTypeEMail, efFldTypeHTTP, efFldTypeLookup, efFldTypeRadio, efFldTypeCheck, efFldTypeSelect
              Case Else
                Select Case Session(CurrentTable & Session("OrderField") & "Type")
                  ' (Se o campo for texto)
                  Case efDataTypeVarChar, efDataTypeChar, efDataTypeText
                    Response.Write "<SPAN CLASS=""HT"">" & EditCaption(False) & "</SPAN>"
                    sobjShow.BR
                    sobjShow.Nbsp
                    sobjShow.Nbsp
                    sobjShow.Nbsp
                    sobjShow.Nbsp
                    ' (Se o campo for texto numérico)
                    If Session(CurrentTable & Session("OrderField") & "AlphabeticIndex") = "1" Then
                      For i = 0 To 9
                        Response.Write "<A CLASS=""HT"" HREF=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & sparEditAction & WhatList & "&" & efQueryStrOrderField & "=" & Session("OrderField") & GetList & GetFilter & "&" & efQueryStrFind & "=" & i & "#find" & i & """>" & i & "</A>&nbsp;&nbsp;"

                      Next

                    ' (Se o campo for texto alfa-numérico)
                    ElseIf Session(CurrentTable & Session("OrderField") & "AlphabeticIndex") = "2" Then
                      For i = 65 To 90
                        Response.Write "<A CLASS=""HT"" HREF=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & sparEditAction & WhatList & "&" & efQueryStrOrderField & "=" & Session("OrderField") & GetList & GetFilter & "&" & efQueryStrFind & "=" & Chr(i) & "#find" & Chr(i) & """>" & Chr(i) & "</A>&nbsp;&nbsp;"

                      Next
                      For i = 0 To 9
                        Response.Write "<A CLASS=""HT"" HREF=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & sparEditAction & WhatList & "&" & efQueryStrOrderField & "=" & Session("OrderField") & GetList & GetFilter & "&" & efQueryStrFind & "=" & i & "#find" & i & """>" & i & "</A>&nbsp;&nbsp;"

                      Next

                    ' (Se o campo for texto alfabético)
                    Else
                      For i = 65 To 90
                        Response.Write "<A CLASS=""HT"" HREF=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & sparEditAction & WhatList & "&" & efQueryStrOrderField & "=" & Session("OrderField") & GetList & GetFilter & "&" & efQueryStrFind & "=" & Chr(i) & "#find" & Chr(i) & """>" & Chr(i) & "</A>&nbsp;&nbsp;"

                      Next

                    End If
               End Select
             End Select
          End If
          Table.EndColumn
          Table.EndRow
          Table.ColumnAlign = ""
          Table.ColumnColSpan = ""
        End If
        '----------------------------

      End If
      Table.BeginRow 2, True

' (Montagem dos títulos do cabeçalho)
      BeginReadOperation strList

      Dim intNumHeader, intNumArrHeaders, intAux As Integer

      intNumHeader = 0
      intCount = 1
      intNumArrHeaders = 0

' (Se houver algum super-título)
      If Session(CurrentTable & "HeaderNumber") > 0 Then
        Table.ColumnRowSpan = 2

      End If

      Do While IsThereOperation
        Select Case NextOperation
          Case ",", ";", "-", "="

          Case Else
            Session("CurrentField") = LastOperation
            intNumHeader = intNumHeader + 1
            If Session(CurrentTable & "HeaderNumber") > 0 Then
' (Se a coluna atual for a de inicio do próximo super-título então coloca o super-título e guarda quais os
' títulos foram omitidos, para depois colocá-los, senão coloca o título)
              If intNumHeader = Session(CurrentTable & intCount & "HeaderBegin") Then
                Table.ColumnRowSpan = ""
                Table.ColumnColSpan = Session(CurrentTable & intCount & "HeaderWidth")
                Table.Column "<SPAN CLASS=""HL"">" & Session(CurrentTable & intCount & "HeaderText") & "</SPAN>"
                Table.ColumnRowSpan = 2
                Table.ColumnColSpan = ""
                intAux = 1
                intNumArrHeaders = intNumArrHeaders + 1
                arrFields(intNumArrHeaders) = LastOperation
                If Session(CurrentTable & "HeaderNumber") > intCount Then
                  intCount = intCount + 1

                End If

              Else
                If intAux > 0 And intAux < Session(CurrentTable & intCount & "HeaderWidth") Then
                  intNumArrHeaders = intNumArrHeaders + 1
                  arrFields(intNumArrHeaders) = LastOperation
                  intAux = intAux + 1

                Else
                  If blnHaveFind And sblnFind And Session("CurrentField") = Session("OrderField") Then
                    Table.Column "&nbsp;"

                  End If

                  intAux = 0

                  EditTableHeader EditCaption(False), Session("CurrentField"), Session(CurrentField & "ListSize")

                 End If
              End If

            Else
              If blnHaveFind And sblnFind And Session("CurrentField") = Session("OrderField") Then
                Table.Column "&nbsp;"

              End If

              EditTableHeader EditCaption(False), Session("CurrentField"), Session(CurrentField & "ListSize")

            End If

         End Select

      Loop

' (Se títulos foram omitidos então coloca-os)
      If intNumArrHeaders > 0 Then
        Table.ColumnRowSpan = ""
        Table.EndRow
        Table.BeginRow 2
        For i = 1 To intNumArrHeaders
          Session("CurrentField") = arrFields(i)
          EditTableHeader Session(CurrentField & "Caption"), Session("CurrentField"), Session(CurrentField & "ListSize")

        Next

      End If
      Table.EndRow

' (Conteúdo da listagem)
     Dim blnFirstFind As Boolean
      blnFirstFind = True
     Dim strBGColor, strBGColor2, strFontColor, strFontColor2 As String
      Do While Not sobjRS2.EOF
        BeginReadOperation strList

' (Se este registro for um dos que entraram em alguma marcação então cria a linha da tabela aplicando
' a cor e cor da fonte da linha, senão somente cria a linha da tabela)
        j = 0
        If pacMarks.Count > 0 Then
          strAux = ""
          For i = 1 To Session(CurrentTable & "FieldNumber")
            If Session(CurrentTable & i & "IsKey") Then
              strAux = strAux & sobjRS2(Session(CurrentTable & i & "Field")) & ";"

            End If
          Next
          If pacMarks.Find(strAux) Then
            j = pacMarks.Value(strAux)
            pacMarks.Remove strAux
            If Session(CurrentTable & j & "MarkBGColor") > "" Then
              strBGColor = Session(CurrentTable & j & "MarkBGColor")

            Else
              strBGColor = Table.HeadRowColor

            End If
            If Session(CurrentTable & j & "MarkFontColor") > "" Then
              strFontColor = Session(CurrentTable & j & "MarkFontColor")

            Else
              strFontColor = Table.Style.HeaderFont.Color

            End If
            strBGColor2 = Table.RowColor
            strFontColor2 = Table.Style.InternalFont.Color
            Table.RowColor = strBGColor
            Table.Style.InternalFont.Color = strFontColor
            Table.BeginRow 1

          Else
             Table.BeginRow 1

          End If
        Else
           Table.BeginRow 1

        End If

' (Formata e apresenta cada campo da listagem)
        Do While IsThereOperation
          Select Case NextOperation
            Case ",", ";", "-", "="
            Case Else
              Session("CurrentField") = LastOperation
              If blnHaveFind And sblnFind And Session("CurrentField") = Session("OrderField") Then
                Table.BeginColumn
' (Cria a coluna que contém o indicador do registro encontrado no AlphabticIndex)
                If blnFirstFind Then
                  If sparEditFind = LCase(sobjString.ChangeChar(Left(sobjRS2(Session(CurrentTable & Session("OrderField") & "Field")), 1))) Then
                    Response.Write "<A NAME=""#find" & sparEditFind & """>"
                    sobjShow.Image Initializer.HiperToolsWorkPath & "/ArrowRight.gif", , sobjString.FormatText(LanguageLocalizer(34), sparEditFind)
                    blnFirstFind = False

                  Else
                    sobjShow.Nbsp

                  End If
                Else
                  sobjShow.Nbsp

                End If
                Table.EndColumn

              End If

' (Alinha a direita os valores de tipo numérico)
              Select Case TypeName(EditFieldValue(sobjRS2))
                Case "Long", "Integer", "Currency", "Double", "Single"
                  Table.ColumnAlign = "Right"

              End Select

              Table.ColumnWidth = Session(CurrentField & "ListSize")

' (Aplica a marca de linha na coluna, se for o caso de um registro selecionado e esta coluna for marcada
' ao invés da linha[passado no parametro do AddMark] senão cria a coluna)
              If pacMarks.Count > 0 Then
               If pacMarks.Find(strAux & Session("CurrentField")) Then
                 l = pacMarks.Value(strAux & Session("CurrentField"))
                 pacMarks.Remove strAux & Session("CurrentField")
                 If Session(CurrentTable & l & "MarkBGColor") > "" Then
                   strBGColor = Session(CurrentTable & l & "MarkBGColor")

                 Else
                   strBGColor = Table.HeadRowColor

                 End If
                 If Session(CurrentTable & l & "MarkFontColor") > "" Then
                   strFontColor = Session(CurrentTable & l & "MarkFontColor")

                 Else
                   strFontColor = Table.Style.HeaderFont.Color

                 End If
                 strFontColor2 = Table.Style.InternalFont.Color
                 Table.ColumnColor = strBGColor
                 Table.Style.InternalFont.Color = strFontColor
                 Table.BeginColumn
                 'Alterado (Ridai)
                 Table.ColumnColor = ""
                 '-------------
                 Table.Style.InternalFont.Color = strFontColor2

               Else
                Table.BeginColumn

               End If
              Else
                 Table.BeginColumn

              End If

' (Cria o ExternalLink ou InternalLink, o que for o caso)
              BeginMakeLink False

' (Preenche o dado propriamente dito, da coluna)
              Select Case Session(CurrentField & "EditOption")
                Case efFldTypeImage, efFldTypeEMail, efFldTypeHTTP
                  Response.Write EditFieldValue(sobjRS2)

                Case efFldTypeLookup, efFldTypeRadio, efFldTypeCheck, efFldTypeSelect
                  Response.Write sobjString.BreakString(EditFieldValue(sobjRS2), Session(CurrentField & "ListChars"))

                Case Else
                  Response.Write sobjString.BreakString(sobjString.Mask(FormatVariable(EditFieldValue(sobjRS2), Session(CurrentField & "Type")), Session(CurrentField & "Mask"), Session(CurrentField & "MaskCompleter")), Session(CurrentField & "ListChars"))

              End Select
              Table.ColumnAlign = ""

' (Fecha o link)
              EndMakeLink

              Table.EndColumn

          End Select

        Loop
        Table.EndRow
' (Volta a cor original, se for uma linha marcada)
        If j > 0 Then
          Table.RowColor = strBGColor2
          Table.Style.InternalFont.Color = strFontColor2

        End If
        sobjRS2.MoveNext
      Loop

      Table.EndTable

      Browse.Bar

      sobjShow.CenterEnd
    End If
    sobjRS2.Close

  Else
' (Coloca uma mensagem se não houver itens na listagem.
' O programador pode escolher a mensagem[sub EmptyListMessage] que irá aparecer neste caso senão
' aparece a mensagem padrão)
    If Session(CurrentTable & "EmptyListMessage") > "" Then
      sobjShow.MessageError "<CENTER>" & Session(CurrentTable & "EmptyListMessage") & "<P>"
      sobjShow.Paragraph

    Else
      sobjShow.MessageError "<CENTER>" & sobjString.FormatText(LanguageLocalizer(35), Session(CurrentTable & "Caption")) & "<P>"

    End If
    If Session(CurrentTable & "InsertLink") And Not Filter > "" And Session(CurrentTable & "EditOption") > 0 Then
      strEditable = sparEditEditable
      sparEditEditable = 1
      sobjShow.HTMLCR " <A HREF=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionAdd & _
                      WhereHow & """>" & LanguageLocalizer(36) & "</A><P></CENTER>"
      sparEditEditable = strEditable
    End If

  End If

  sobjRS.Close

  sobjShow.BR

  'Criado (Ridai Govinda)
  RecoverFieldsDef
  RestoreLookupTableNames
  '--------------------

' (Coloca os links de "Página Inicial" e "Última Opção")
  FPLPLink

End Sub

Private Sub EditTableHeader(ByVal strField As String, ByVal strOrderField As String, ByVal intSize As Integer)
' (Cria um título do cabecalho de listagem, com seu link conforme a ordenação)

  Table.ColumnWidth = intSize
  Table.BeginColumn
' (Se este for o campo da ordenação)
  If LCase(sobjString.ToString(Session("OrderField"))) = LCase(sobjString.ToString(strOrderField)) Then
' (Se a ordenação for decrescente)
    If Session("OrderDesc") = "1" Then
      sobjShow.Image Session("htDefaultHiperToolsWorkPath") & "/ArrowDown.gif", "", sobjString.FormatText(LanguageLocalizer(37), strField), Initializer.ScriptURL & "?" & efQueryStrAction & "=" & sparEditAction & WhatList & GetEditable & "&" & efQueryStrOrderField & "=" & Server.URLEncode(strOrderField) & GetList & GetFilter, "", 16, 9
      Response.Write " <A CLASS=HL HRef=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & sparEditAction & WhatList & "&" & efQueryStrOrderField & "=" & Server.URLEncode(strOrderField) & GetEditable & GetList & GetFilter & """>" & strField & "</A>"

    Else
      sobjShow.Image Session("htDefaultHiperToolsWorkPath") & "/ArrowUp.gif", "", sobjString.FormatText(LanguageLocalizer(38), strField), Initializer.ScriptURL & "?" & efQueryStrAction & "=" & sparEditAction & WhatList & GetEditable & "&" & efQueryStrOrderField & "=" & Server.URLEncode(strOrderField) & efQueryStrOrderDescStr & GetList & GetFilter, "", 16, 9
      Response.Write " <A CLASS=HL HRef=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & sparEditAction & WhatList & "&" & efQueryStrOrderField & "=" & Server.URLEncode(strOrderField) & GetEditable & efQueryStrOrderDescStr & GetList & GetFilter & """>" & strField & "</A>"

    End If
  Else
    Response.Write "<A CLASS=HL HRef=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & sparEditAction & WhatList & "&" & efQueryStrOrderField & "=" & Server.URLEncode(strOrderField) & GetEditable & GetList & GetFilter & """>" & strField & "</A>"

  End If
  Table.EndColumn

End Sub

Private Function EditListStr() As String
' (Recria o parametro do QueryString que indica qual a listagem atual)
 EditListStr = "&" & efQueryStrList & "=" & sparEditList

End Function

Private Function ReadFieldValues(ByVal blnEdit As Boolean) As Boolean
  ' (Le os valores dos campos para edição e visualisação de um registro)
  Dim sql, i, blnFirst

  ' (Application(CurrentTable & MakeTarget) representa o registro sendo editado atualmente pelo usuário
  ' se Application(CurrentTable & MakeTarget) contiver uma data então outro usuário o está usando no
  ' momento.
  ' Mas como isto não está funcionando perfeitamente, devido a incerteza de que o Session_End está sendo
  ' chamado, estou esvaziando este application aqui)

  'Alterado (Ridai Govinda)
  'Application.Lock
  'Application(CurrentTable & MakeTarget) = ""
  'Application.Unlock

  'Criado (Ridai Govinda)
  'Desfaz o lock se já tiver decorrido LockTimeOut minutos:
  RecordLocked efLockClearLock

  If (Not RecordLocked()) Or (Not blnEdit) Then
  Rem If (Not Application(CurrentTable & MakeTarget) > "") Or (Not blnEdit) Then
  '-----------------------
' (Monta o SELECT dos dados a serem editados ou visualizados apresentando somente os campos chave)
    blnFirst = True
    sql = "SELECT "

    For i = 1 To Session(CurrentTable & "FieldNumber")
      If Session(CurrentTable & i & "IsKey") Or InStr(LCase(sparEditFind), LCase(Session(CurrentTable & i & "Field"))) > 0 Then
        If Not blnFirst Then
          sql = sql & ", "

        Else
          blnFirst = False

        End If

        sql = sql & Session(CurrentTable & i & "Field")

      End If
    Next

    sql = sql & " FROM " & CurrentTable & Session(CurrentTable & "JoinFormUnit") & Filter

    sobjConn.Close

    'Alterado (Ridai Govinda)
    Rem sobjConn.Provider = "SQLOLEDB.1"
    '-----------------------

    sobjConn.Open ConnectionString, UserID, Password

    Rem sobjCMD.ActiveConnection = sobjConn
    Rem sobjCMD.CommandText = sql
    Rem sobjCMD.CommandType = adCmdText

    Rem    sobjShow.MessageError sql
    Select Case ConnectionType
    Case efConnDBProviderOracle, efConnDBDriverSQL
      sobjRS.CursorLocation = adUseClient
    End Select
    'Alterado/Criado (Ridai)
    Set CurrConnection = sobjConn

    If ShowSQLQuery Then ShowCurrentSQL sql, "ReadFieldValues", adOpenStatic, adLockReadOnly

    On Error Resume Next

    sobjRS.Open sql, sobjConn, adOpenStatic, adLockReadOnly

    If Err.number <> 0 Then
      If DebugMode Then
        ShowSQLErrors sql, "ReadFieldValues", adOpenStatic, adLockReadOnly
      Else
        ShowDefaultError
      End If
      Err.Clear
      Exit Function

    End If
    On Error GoTo 0
    '----

' (Se foi encontrado algum registro então localiza o registro procurado pelo parametro sparEditFind)
    If Not sobjRS.EOF And Not sobjRS.BOF Then
      sobjRS.MoveLast
      sobjRS.MoveFirst
      If sparEditFind > "" Then
        sobjRS.Find sparEditFind
        If sobjRS.EOF Or sobjRS.BOF Then
          sobjRS.MoveFirst

        End If
      Else
        If (sparEditMove <> conMoveFirst) And (sparEditMove > "") Then
          If (sparEditMove = conMoveLast) Then
            sobjRS.MoveLast

          Else
            sobjRS.Move CInt(sparEditMove)

          End If
        End If
      End If

      For i = 1 To Session(CurrentTable & "FieldNumber")
        If Session(CurrentTable & i & "IsKey") Then
          Session(CurrentTable & i & "Value") = sobjRS(Session(CurrentTable & i & "Field"))

        End If
      Next

' (Armazena a posição do registro atual e o nº total de registros)
      Session("AbsolutePosition") = sobjRS.AbsolutePosition
      Session("RecordCount") = sobjRS.RecordCount

      sobjRS.Close

      Database.AndState = False

' (Monta o SELECT do registro selecionado, já contendo todos os campos e seus JOINs, se for campo
' externo)
      sql = "SELECT "
      For i = 1 To Session(CurrentTable & "FieldNumber")
        If i > 1 Then
          sql = sql & ", "

        End If

        sql = sql & OwnerTable(CurrentTable, i) & "." & Session(CurrentTable & i & "Field")

      Next

      sql = sql & " FROM " & CurrentTable & Session(CurrentTable & "JoinFormUnit")

      For i = 1 To Session(CurrentTable & "FieldNumber")
        If Session(CurrentTable & i & "IsKey") Then
          sql = sql & Database.PutAnd & OwnerTable(CurrentTable, i) & "." & Session(CurrentTable & i & "Field") & "=" & FormatForSQL(Session(CurrentTable & i & "Value"), Session(CurrentTable & i & "Type"))

        End If
      Next

      sobjConn.Close
      sobjConn.Open ConnectionString, UserID, Password

      Rem sobjCMD.ActiveConnection = sobjConn
      Rem sobjCMD.CommandText = sql
      Rem sobjCMD.CommandType = adCmdText

      sobjRS.CursorLocation = adUseServer

      'Alterado/Criado (Ridai)
      Set CurrConnection = sobjConn

      If ShowSQLQuery Then ShowCurrentSQL sql, "ReadFieldValues", adOpenKeyset, adLockReadOnly

      On Error Resume Next

      sobjRS.Open sql, sobjConn, adOpenStatic, adLockReadOnly

      If Err.number <> 0 Then
        If DebugMode Then
          ShowSQLErrors sql, "ReadFieldValues", adOpenStatic, adLockReadOnly
        Else
          ShowDefaultError
        End If
        Err.Clear
        Exit Function

      End If
      On Error GoTo 0
      '----

' (Guarda o valor de cada campo em dois Sessions um é o valor que poderá ser alterado pelo usuário
' [Value] e o outro é guardado[OldValue] para ser utilizado na localização do registro entre outros
' motivos)
      For i = 1 To Session(CurrentTable & "FieldNumber")
' (Se o campo for text transforma o valor para o tipo string, a intenção é que todos os valores sejam
' assim transformados para diminuir o nº de formatações de campos e fazer o INSERT e o UPDATE
' em SQL)
        If Session(CurrentTable & i & "Type") = "text" Then
          Session(CurrentTable & i & "OldValue") = sobjString.ToString(sobjRS(Session(CurrentTable & i & "Field")).Value)
          Session(CurrentTable & i & "Value") = sobjString.ToString(sobjRS(Session(CurrentTable & i & "Field")).Value)

        Else
          Session(CurrentTable & i & "OldValue") = sobjRS(Session(CurrentTable & i & "Field"))
          Session(CurrentTable & i & "Value") = sobjRS(Session(CurrentTable & i & "Field"))

        End If
      Next
      sobjRS.Close

    Else
      sobjRS.Close
      If Not blnEdit Then
' (Se não for encontrado nenhum registro e não houver filtro então faz uma inserção)
        If Not sparEditFilter > "" Then
          Session("RecordAction") = "1"
          sparEditAction = efQSActionAdd
          EditClear sparEditAction

        Else
          If Session(CurrentTable & "ClientValidation") Then
            Message = "<SCRIPT LANGUAGE=JAVASCRIPT>alert(""" & LanguageLocalizer(39) & """);</SCRIPT>"

          End If
          Message = Message & "<FONT COLOR=RED>" & "<B>" & LanguageLocalizer(39) & "</B></FONT>"
          sparEditFilter = ""
          ReadFieldValues = ReadFieldValues(blnEdit)

        End If

      Else
        Message = LanguageLocalizer(40)
        ReadFieldValues = False

      End If

   End If
   ReadFieldValues = True

 Else
   'Alterado (Ridai Govinda)
   ShowLockMessage
   ReadFieldValues = False
   '-----------------------

 End If

End Function

Rem ============================
Rem Criado (Ridai Govinda)
Rem Mais uma coisa _faltando_!!!
Rem ----------------------------
Private Sub EditClear2()
  Dim strSQL As String, strAux As String
  Dim strTable As String, strField As String
  Dim i As Integer, intPos As Integer, intPos2 As Integer, intPos3 As Integer

  strTable = ""
  strField = ""

  For i = 1 To Session(CurrentTable & "FieldNumber")
    Select Case Left(Session(CurrentTable & i & "DefaultValue"), 8)
      ' (Campo contador) - Não faz nada.
      Case constrProximo
      Case Else
        If sparEditDefaults > "" Then
          intPos3 = InStr(LCase(sparEditDefaults), LCase(Session(CurrentTable & i & "Field")) & "=")
          If intPos3 > 0 Then
            intPos = intPos3 + Len(Session(CurrentTable & i & "Field"))
            strAux = Right(sparEditDefaults, Len(sparEditDefaults) - intPos)
            intPos2 = InStr(strAux, ";")
            If intPos2 > 0 Then
              If Session(CurrentTable & i & "PersistentDefValue") Then Session(CurrentTable & i & "Value") = FormatField(Left(strAux, intPos2 - 1), Session(CurrentTable & i & "Type"), Session(CurrentTable & i & "Size"), Session(CurrentTable & i & "CharCase"))
              sparEditDefaults = Left(sparEditDefaults, intPos3 - 1) & Right(strAux, Len(strAux) - intPos2)

            Else
              If Session(CurrentTable & i & "PersistentDefValue") Then Session(CurrentTable & i & "Value") = FormatField(strAux, Session(CurrentTable & i & "Type"), Session(CurrentTable & i & "Size"), Session(CurrentTable & i & "CharCase"))
              sparEditDefaults = Left(sparEditDefaults, intPos3 - 1)

            End If

          Else
            If Session(CurrentTable & i & "PersistentDefValue") Then Session(CurrentTable & i & "Value") = Session(CurrentTable & i & "DefaultValue")

          End If

        Else
          ' (Se o campo é de data de atualização então o valor recebe hoje)
          ' -- Implementar os outros (Update/Delete/Insert).
          If Session(CurrentTable & i & "EditOption") = efFldTypeAtualization And _
             Session(CurrentTable & i & "EditOption") = efFldTypeDateOfUpdate And _
             Session(CurrentTable & i & "EditOption") = efFldTypeDateOfDelete And _
             Session(CurrentTable & i & "EditOption") = efFldTypeDateOfInsert Then
            Session(CurrentTable & i & "Value") = Now()

          Else
            If Session(CurrentTable & i & "PersistentDefValue") Then Session(CurrentTable & i & "Value") = Session(CurrentTable & i & "DefaultValue")

          End If
        End If
    End Select
  Next

End Sub
Rem ----------------------------

Private Sub EditClear(ByVal strClearOption As String)
' (Coloca os valores default nos campos Sessions, para fazer uma inclusão)
  Dim strSQL As String, strAux As String
  Dim strTable As String, strField As String
  Dim i As Integer, intPos As Integer, intPos2 As Integer, intPos3 As Integer

  strTable = ""
  strField = ""

  'Criado (Ridai Govinda)

  For i = 1 To Session(CurrentTable & "FieldNumber")
    Select Case Left(Session(CurrentTable & i & "DefaultValue"), 8)
    'Alterado/Removido (Ridai Govinda Pombo)
' (Campo contador)
      Case constrProximo
        ' (Campo da tabela principal buscando o valor em um campo de outra tabela)
        If Len(Session(CurrentTable & i & "DefaultValue")) > 8 Then
          strTable = Mid(Session(CurrentTable & i & "DefaultValue"), 9, InStr(Session(CurrentTable & i & "DefaultValue"), ".") - 9)
          strField = Mid(Session(CurrentTable & i & "DefaultValue"), InStr(Session(CurrentTable & i & "DefaultValue"), ".") + 1)

        ' (Campo de outra tabela)
        ElseIf Session(CurrentTable & i & "ExternalAddress") Then
          strField = Session(CurrentTable & i & "Field")
          strTable = Session(CurrentTable & i & "OwnerTable")

        ' (Campo da tabela principal)
        Else
          strField = Session(CurrentTable & i & "Field")
          strTable = CurrentTable

        End If

        'Criado (Ridai Govinda)
        Session(CurrentTable & i & "Value") = 0
        NextIdentityRequestTable = strTable
        NextIdentityRequestField = strField
        NextIdentityTarget = MakeTarget
        '---------------------

      Case Else
' (Se está sendo feita uma inclusão[pq a copia tb chama esta função para dar valor aos campos
' contadores])
        If strClearOption = efQSActionAdd Then
' (Se está há algum valor default vindo pelo parametro sparEditDefaults)
          If sparEditDefaults > "" Then
            intPos3 = InStr(LCase(sparEditDefaults), LCase(Session(CurrentTable & i & "Field")) & "=")
            If intPos3 > 0 Then
              intPos = intPos3 + Len(Session(CurrentTable & i & "Field"))
              strAux = Right(sparEditDefaults, Len(sparEditDefaults) - intPos)
              intPos2 = InStr(strAux, ";")
              If intPos2 > 0 Then
                Session(CurrentTable & i & "Value") = FormatField(Left(strAux, intPos2 - 1), Session(CurrentTable & i & "Type"), Session(CurrentTable & i & "Size"), Session(CurrentTable & i & "CharCase"))
                sparEditDefaults = Left(sparEditDefaults, intPos3 - 1) & Right(strAux, Len(strAux) - intPos2)

              Else
                Session(CurrentTable & i & "Value") = FormatField(strAux, Session(CurrentTable & i & "Type"), Session(CurrentTable & i & "Size"), Session(CurrentTable & i & "CharCase"))
                sparEditDefaults = Left(sparEditDefaults, intPos3 - 1)

              End If

            Else
              Session(CurrentTable & i & "Value") = Session(CurrentTable & i & "DefaultValue")

            End If

          Else
' (Se o campo é de data de atualização então o valor recebe hoje)
            If Session(CurrentTable & i & "EditOption") = efFldTypeAtualization And _
               Session(CurrentTable & i & "EditOption") = efFldTypeDateOfUpdate And _
               Session(CurrentTable & i & "EditOption") = efFldTypeDateOfDelete And _
               Session(CurrentTable & i & "EditOption") = efFldTypeDateOfInsert Then
              Session(CurrentTable & i & "Value") = Now

            Else
              Session(CurrentTable & i & "Value") = Session(CurrentTable & i & "DefaultValue")

            End If
          End If
        End If

    End Select
  Next

End Sub

Private Sub GetForm(ByVal strFields As String)
' (Busca os valores dos campos nos inputs, já formatando-os)
Dim i As Integer

  BeginReadOperation strFields
  Dim Width

  Width = 0
  Do While IsThereOperation
    Select Case NextOperation
       Case ",", ";", "-", ""
       Case "="
         IsThereOperation

       Case Else

         'Alterado (Ridai Govinda)
         If FindField(CurrentTable, LastOperation) And _
            Not (IdentityCheck(False) And Not AllowIdentityUpdate) Then
         '----------------
           Select Case Session(CurrentField & "EditOption")
' (Campo de data de atualização)
             Case efFldTypeAtualization, efFldTypeDateOfUpdate, efFldTypeDateOfDelete, efFldTypeDateOfInsert
               Session(CurrentField & "Value") = sobjString.ZTILongDate(Now)

' (Campo de data separada)
             Case efFldTypeSeparateDate
                Session(CurrentField & "Value") = sobjString.ImportDate(Request.Form(Session(CurrentField & "Field") & _
                                                "DayValue") & Request.Form(Session(CurrentField & "Field") & _
                                                "MonthValue") & Request.Form(Session(CurrentField & "Field") & _
                                                "YearValue"), stDateTypeDDMMAAAA)

' (Campos de senha e arquivo)
              Case efFldTypePassword, efFldTypeFile
                If Request.Form(Session(CurrentField & "Field")) > "" Then
                  If LCase(Request.Form(Session(CurrentField & "Field"))) = "nada" Then
                    Session(CurrentField & "Value") = ""

                  Else
                    Session(CurrentField & "Value") = FormatField(Request.Form(Session(CurrentField & "Field")), Session(CurrentField & "Type"), Session(CurrentField & "Size"), Session(CurrentField & "CharCase"))

                  End If
                End If

' (Campo imagem)
              Case efFldTypeImage

' (Campo CheckBox)
              Case efFldTypeCheck
                If LCase(Request.Form(Session(CurrentField & "Field"))) = LCase(sobjString.ToString(Session(CurrentField & ",1Value"))) Then
                  Session(CurrentField & "Value") = Session(CurrentField & ",1Value")

                Else
                  Session(CurrentField & "Value") = Session(CurrentField & ",2Value")

                End If

' (Outros campos)
              Case Else
                Select Case Session(CurrentField & "Type")
' (Tipo de dado numérico ou data e hora)
                  Case efDataTypeFloat, efDataTypeReal, efDataTypeMoney, efDataTypeDateTime
                    Session(CurrentField & "Value") = FormatField(Request.Form(Session(CurrentField & "Field")), Session(CurrentField & "Type"), Session(CurrentField & "Size"), Session(CurrentField & "CharCase"))

' (Outros tipos de dado)
                  Case Else
                    Session(CurrentField & "Value") = Request.Form(Session(CurrentField & "Field"))

                End Select
           End Select
           'Criado (Ridai Govinda Pombo)
           If Not CheckValueAfterGet() Then
             ShowError efErrorDuplicateKey, "EditForm", "GetForm", LanguageLocalizer(6)
           End If
           '---------------------------

         End If
     End Select
  Loop

End Sub

Private Sub EditSave()
  ' (Grava as alterações no BD. É feito separadamente para inserção e alteração, mas o objetivo é uní-las
  ' numa só rotina que o fará diretamente nm comando SQL)
  Dim i As Integer, j As Integer, intField As Integer
  Dim strAnotherTable As String, strSaved As String
  Dim strCurrentTable As String, strTable As String, strFileName As String
  Dim strSQL As String
  Dim lngIdentity As Long
  Dim vntValue

  lngIdentity = 0

  ' (Modo de inserção)
  If Session("RecordAction") = "1" Then
    Rem sobjCMD.CommandText = CurrentTable
    Rem sobjCMD.CommandType = adCmdTable
    strSQL = CurrentTable

    'Alterado/Criado (Ridai)
    If ConnectionType = efConnDBProviderOracle Then sobjRS.CursorLocation = adUseClient

    Set CurrConnection = sobjConn

    If ShowSQLQuery Then ShowCurrentSQL strSQL, "EditSave", adOpenDynamic, adLockOptimistic

    On Error GoTo SQLErrorHandler

    sobjRS.Open strSQL, sobjConn, adOpenDynamic, adLockOptimistic, adCmdTable

    sobjRS.AddNew

    On Error GoTo 0
    '---

    strAnotherTable = ""
    For i = 1 To Session(CurrentTable & "FieldNumber")
      If Session(CurrentTable & i & "ShowOption") <> 1 Then
        If Not Session(CurrentTable & i & "ExternalAddress") Then
          'Alterado/Criado (Ridai Govinda)
          If Left(Session(CurrentTable & i & "DefaultValue"), 8) = constrProximo Then
            lngIdentity = GetIdentity(CurrentTable)
            Session(CurrentTable & i & "Value") = lngIdentity
          End If
          SaveFile lngIdentity, i
          '------------------------------

          ' (Aplica os valores, formatados, nos campos da tabela principal)

          vntValue = FormatFieldtoSave(Session(CurrentTable & i & "Value"), _
                     Session(CurrentTable & i & "Type"), _
                     Session(CurrentTable & i & "Size"), _
                     Session(CurrentTable & i & "CharCase"))

          If Not Session(CurrentTable & i & "Null") And IsNull(vntValue) Then
            ShowError efErrorNullRequiredField, "EditForm", "EditSave", sobjString.FormatText(LanguageLocalizer(7), Replace(Session(CurrentTable & i & "Caption"), "&", ""))
          End If

          sobjRS(Session(CurrentTable & i & "Field")) = vntValue

          Session(CurrentTable & i & "OldValue") = Session(CurrentTable & i & "Value")

        Else
' (Guarda o nome de tabela de um campo externo)
          If Not strAnotherTable > "" Then
            strAnotherTable = Session(CurrentTable & i & "OwnerTable")

          End If
        End If
      End If

    Next
    'Alterado/Criado (Ridai)

    If ShowSQLQuery Then ShowCurrentSQL strSQL, "EditSave", adOpenDynamic, adLockOptimistic

    On Error GoTo SQLErrorHandler

    sobjRS.Update

    On Error GoTo 0
    '---
    sobjRS.Close

    strSaved = CurrentTable & "&"

' (Enquanto houver alguma tabela externa salva todos os seus campos)
    Do While strAnotherTable <> ""
      strCurrentTable = strAnotherTable
      strSaved = strSaved & "(" & strCurrentTable & ")"
      strAnotherTable = ""

      Rem sobjCMD.CommandText = strCurrentTable
      Rem sobjCMD.CommandType = adCmdTable
      strSQL = strCurrentTable


      'Alterado/Criado (Ridai)
      If ShowSQLQuery Then ShowCurrentSQL strSQL, "EditSave", adOpenDynamic, adLockOptimistic

      On Error GoTo SQLErrorHandler

      sobjRS.Open strSQL, sobjConn, adOpenDynamic, adLockOptimistic, adCmdTable

      sobjRS.AddNew

      On Error GoTo 0
      '---

      strTable = ""
      If FindRelation(CurrentTable, strCurrentTable) Then
         strTable = CurrentTable & CurrentRelation
         intField = 1

      ElseIf FindRelation(strCurrentTable, CurrentTable) Then
         strTable = strCurrentTable & CurrentRelation
         intField = 2

      End If

      If strTable > "" Then
        For j = 1 To Session(strTable & "FieldtoFieldNumber")
          If Session(CurrentTable & i & "ShowOption") <> 1 Then
            If FindField(CurrentTable, Session(strTable & "," & j & "RelationField" & intField)) Then

              'Alterado/Criado (Ridai Govinda)
              If Left(Session(strTable & j & "DefaultValue"), 8) = constrProximo Then
                lngIdentity = GetIdentity(strTable)
                Session(CurrentField & "Value") = lngIdentity
              End If
              SaveFile lngIdentity, i
              '------------------------------

              sobjRS(Session(strTable & "," & j & "RelationField" & (3 - intField))) = FormatFieldtoSave(Session(CurrentField & "Value"), _
                                                                       Session(CurrentField & "Type"), EditSize, _
                                                                       Session(CurrentTable & i & "CharCase"))
              Session(CurrentTable & i & "OldValue") = Session(CurrentTable & i & "Value")

            End If
          End If
        Next
      End If

      For i = 1 To Session(CurrentTable & "FieldNumber")
        If Session(CurrentTable & i & "ShowOption") <> 1 Then
          If Session(CurrentTable & i & "OwnerTable") = strCurrentTable Then

            'Alterado/Criado (Ridai Govinda)
            If Left(Session(CurrentTable & i & "DefaultValue"), 8) = constrProximo Then
              Session(CurrentTable & i & "Value") = GetIdentity(CurrentTable)
            End If
            '------------------------------

            sobjRS(Session(CurrentTable & i & "Field")) = FormatFieldtoSave(Session(CurrentTable & _
                                                                     i & "Value"), Session(CurrentTable & i & "Type"), _
                                                                     Session(CurrentTable & i & "Size"), _
                                                                     Session(CurrentTable & i & "CharCase"))
            Session(CurrentTable & i & "OldValue") = Session(CurrentTable & i & "Value")

          Else
            If (Not strAnotherTable > "") And (InStr(strSaved, Session(CurrentTable & i & "OwnerTable")) = 0) Then
              strAnotherTable = Session(CurrentTable & i & "OwnerTable")

            End If
          End If
        End If
      Next

      sobjRS.Update
      sobjRS.Close

    Loop

' (Modo de alteração)
  Else
    Dim sql

    sql = "SELECT * FROM " & CurrentTable

    If sparEditFind > "" Then
      sql = sql & " WHERE " & sparEditFind

    End If

    Rem sobjCMD.CommandText = sql
    Rem sobjCMD.CommandType = adCmdText
    strSQL = sql


    'Alterado/Criado (Ridai)
    If ShowSQLQuery Then ShowCurrentSQL strSQL, "EditSave", adOpenDynamic, adLockOptimistic

    On Error GoTo SQLErrorHandler

    sobjRS.Open strSQL, sobjConn, adOpenDynamic, adLockPessimistic

    On Error GoTo 0
    '---

    strAnotherTable = ""
    For i = 1 To Session(CurrentTable & "FieldNumber")
      If Session(CurrentTable & i & "ShowOption") <> 1 Then
        If Not Session(CurrentTable & i & "ExternalAddress") Then
' (Aplica os valores, formatados, nos campos da tabela principal)
          sobjRS(Session(CurrentTable & i & "Field")) = FormatFieldtoSave(Session(CurrentTable & _
                                                                   i & "Value"), Session(CurrentTable & i & "Type"), _
                                                                   Session(CurrentTable & i & "Size"), _
                                                                   Session(CurrentTable & i & "CharCase"))
          Session(CurrentTable & i & "OldValue") = Session(CurrentTable & i & "Value")

        Else
' (Guarda o nome de tabela de um campo externo)
          If Not strAnotherTable > "" Then
            strAnotherTable = Session(CurrentTable & i & "OwnerTable")

          End If
        End If
      End If
    Next
    sobjRS.Update
    sobjRS.Close

    strSaved = CurrentTable & "&"

' (Enquanto houver alguma tabela externa salva todos os seus campos)
    Do While strAnotherTable <> ""
      strCurrentTable = strAnotherTable
      strSaved = strSaved & "(" & strCurrentTable & ")"
      strAnotherTable = ""

      strTable = ""
      If FindRelation(CurrentTable, strCurrentTable) Then
         strTable = CurrentTable & CurrentRelation
         intField = 1

      ElseIf FindRelation(strCurrentTable, CurrentTable) Then
         strTable = strCurrentTable & CurrentRelation
         intField = 2

      End If

      sql = "SELECT * FROM " & strCurrentTable

      Database.AndState = False

      If strTable > "" Then
        For j = 1 To Session(strTable & "FieldtoFieldNumber")
          If Session(CurrentTable & i & "ShowOption") <> 1 Then
            If FindField(CurrentTable, Session(strTable & "," & j & "RelationField" & intField)) Then
              sql = sql & Database.PutAnd & Session(strTable & "," & j & "RelationField" & (3 - intField)) & " = " & _
                        FormatForSQL(Session(CurrentField & "OldValue"), Session(CurrentField & "Type"))

            End If
          End If
        Next

      End If

      strSQL = sql
      Rem sobjCMD.CommandType = adCmdText

      'Alterado/Criado (Ridai)
      If ShowSQLQuery Then ShowCurrentSQL strSQL, "EditSave", adOpenDynamic, adLockOptimistic

      On Error GoTo SQLErrorHandler

      sobjRS.Open strSQL, sobjConn, adOpenDynamic, adLockOptimistic

      On Error GoTo 0
      '---

      If sobjRS.EOF Or sobjRS.BOF Then
        Rem sobjCMD.CommandType = adCmdTable
        strSQL = strCurrentTable

        sobjRS.Close

        'Alterado/Criado (Ridai)
        If ShowSQLQuery Then ShowCurrentSQL strSQL, "EditSave", adOpenDynamic, adLockOptimistic

        On Error GoTo SQLErrorHandler

        sobjRS.Open strSQL, sobjConn, adOpenDynamic, adLockOptimistic, adCmdTable

        sobjRS.AddNew

        On Error GoTo 0
        '---

      End If

      For i = 1 To Session(CurrentTable & "FieldNumber")
        If Session(CurrentTable & i & "ShowOption") <> 1 Then
          If Session(CurrentTable & i & "OwnerTable") = strCurrentTable Then
              sobjRS(Session(CurrentTable & i & "Field")) = FormatFieldtoSave(Session(CurrentTable & _
                                                                       i & "Value"), Session(CurrentTable & i & "Type"), _
                                                                       Session(CurrentTable & i & "Size"), _
                                                                       Session(CurrentTable & i & "CharCase"))
              Session(CurrentTable & i & "OldValue") = Session(CurrentTable & i & "Value")

          Else
            If (Not strAnotherTable > "") And (InStr(strSaved, Session(CurrentTable & i & "OwnerTable")) = 0) Then
              strAnotherTable = Session(CurrentTable & i & "OwnerTable")

            End If
          End If
        End If
      Next

      If strTable > "" Then
        For j = 1 To Session(strTable & "FieldtoFieldNumber")
          If Session(CurrentTable & i & "ShowOption") <> 1 Then
            If FindField(CurrentTable, Session(strTable & "," & j & "RelationField" & intField)) Then
              sobjRS(Session(strTable & "," & j & "RelationField" & (3 - intField))) = FormatFieldtoSave(Session(CurrentField & "Value"), _
                                                                       Session(CurrentField & "Type"), EditSize, _
                                                                       Session(CurrentTable & i & "CharCase"))
              Session(CurrentTable & i & "OldValue") = Session(CurrentTable & i & "Value")

            End If
          End If
        Next

      End If

      sobjRS.Update
      sobjRS.Close

    Loop

  End If
  Exit Sub

SQLErrorHandler:
  If Err.number <> 0 Then
    If DebugMode Then
      ShowSQLErrors strSQL, "EditSave", adOpenDynamic, adLockOptimistic
    Else
      ShowDefaultError
    End If
    Err.Clear

  End If
End Sub


Private Sub EditDelete()
' (Inicia o processo de exclusão de um registro)
  DeleteFrom CurrentTable, "", "", "", ""

End Sub

Private Sub DeleteFrom(ByVal strTable As String, ByVal strLastTable As String, ByVal strSelect As String, ByVal strLinkField As String, ByVal strLinkField2 As String)
' (Sub recursiva que exclui um registro numa tabela, e todos os registro dependentes, nas tabelas
' interligadas a esta. Parametros: Tabela da exclusão, tabela relacionada que ordenou a exclusão nesta,
' SELECT a ser utilizado na exclusão com sub-query, campo de ligação na primeira tabela, campo de
' ligação na segunda tabela)
Dim strSelectAux, sql As String
Dim i, j As Integer

' (Para cada relacionamento da tabela, monta o WHERE do DELETE e chama novamente a sub)
  For i = 1 To Session(strTable & "RelationNumber")
    strSelectAux = strSelect

    If Not strSelectAux > "" Then
        strSelectAux = strTable & " WHERE "

    End If

' (Monta os relacionamento dos campos das tabelas relacionadas, no WHERE)
    For j = 1 To Session(strTable & i & "FieldtoFieldNumber")
      If strSelectAux > (strTable & " WHERE ") Then
        strSelectAux = strSelectAux & " AND "

      End If
      strSelectAux = strSelectAux & Session(strTable & i & "RelationTable") & "." & Session(strTable & i & "," & j & "RelationField2") & " = " & strTable & "." & Session(strTable & i & "," & j & "RelationField1")

    Next

    If Not strSelect > "" Then
      If sparEditFind > "" Then
        strSelectAux = strSelectAux & " AND " & sparEditFind

      End If
    End If

    strSelectAux = Session(strTable & i & "RelationTable") & ", " & strSelectAux

' (Nova chamada da sub)
    DeleteFrom Session(strTable & i & "RelationTable"), strTable, strSelectAux, Session(strTable & i & ",1RelationField1"), Session(strTable & i & ",1RelationField2")

  Next

' (Exclusão com o select montado)
  If strSelect > "" Then
    sql = "DELETE FROM " & strTable & " WHERE " & strTable & "." & strLinkField2
    sql = sql & " IN (SELECT " & strLastTable & "." & strLinkField & " FROM " & strSelect & ")"

    Rem sobjCMD.CommandType = adCmdText

    'Alterado/Criado (Ridai)
    Set CurrConnection = sobjConn

    If ShowSQLQuery Then ShowCurrentSQL sql, "DeleteFrom", adOpenForwardOnly, adLockReadOnly

    On Error Resume Next

    sobjRS.Open sql, sobjConn, adOpenForwardOnly, adLockReadOnly

    If Err.number <> 0 Then
      If DebugMode Then
        ShowSQLErrors sql, "DeleteFrom", adOpenForwardOnly, adLockReadOnly
      Else
        ShowDefaultError
      End If
      Err.Clear
      Exit Sub

    End If
    On Error GoTo 0
    '----

' (Exclusão com o select vindo do parametro)
  Else
    sql = "DELETE FROM " & strTable

    If sparEditFind > "" Then
      sql = sql & " WHERE " & sparEditFind

    End If

    Rem sobjCMD.CommandText = sql
    Rem sobjCMD.CommandType = adCmdText

    'Alterado/Criado (Ridai)
    Set CurrConnection = sobjConn

    If ShowSQLQuery Then ShowCurrentSQL sql, "DeleteFrom", adOpenForwardOnly, adLockReadOnly

    On Error Resume Next

    sobjRS.Open sql, sobjConn, adOpenForwardOnly, adLockReadOnly

    If Err.number <> 0 Then
      If DebugMode Then
        ShowSQLErrors sql, "DeleteFrom", adOpenForwardOnly, adLockReadOnly
      Else
        ShowDefaultError
      End If
      Err.Clear
      Exit Sub

    End If
    On Error GoTo 0
    '----

  End If

End Sub

Private Sub JavaScriptDefaults()
' (Inicialização dos Javacripts da tela de edição, define todo o escopo básico de funções existentes em
' todas as tela de edição)

' (Abre o script)
  If Not Validate.JSStart Then
    Validate.JSBegin

  End If

  'Alterado (Ridai Govinda)
  If ShowButtons Then
    ' (Faz os testes relacionados aos botões como vaidação de campos, e confirmação de exclusão e
    ' cancelamento)
    'Criado (Ridai Govinda)
    If AllowJavaScriptRefresh > "" And Buttons.ShowEdit Then
      sobjShow.HTMLCR LanguageLocalizer(41)
      sobjShow.HTMLCR "window.setTimeout(""document.frmEditForm.btmAltera.click();"", " & _
          CStr(IIf(sbytLockRemainType > 0, LockTimeOutRemaining * 60000, LockTimeOutRemaining * 1000)) & ");"
      AllowJavaScriptRefresh = ""
    End If
    '----------------------

    sobjShow.HTMLCR "function ButtonsTest(field){"
    sobjShow.HTMLCR "document.frmEditForm.btmEditor.value = field.value;"
    sobjShow.HTMLCR "  switch (field.value){"
    ' (Se houver a opção de cancelar)
    If Session("RecordAction") > "" Then
      sobjShow.HTMLCR "  case """ & Buttons.CancelCaption & """ :"
      ' (Se algum valor de campo tiver sido alterado)
      sobjShow.HTMLCR "    if (blnChanged)"
      sobjShow.HTMLCR "      return confirm(""" & LanguageLocalizer(42) & """);"
      sobjShow.HTMLCR "    else"
      sobjShow.HTMLCR "      return true;"
      sobjShow.HTMLCR "  Break;"

    ' (Se houver a opção de excluir)
    Else
      sobjShow.HTMLCR "  case """ & Buttons.DeleteCaption & """ :"
      sobjShow.HTMLCR "  return confirm(""" & LanguageLocalizer(43) & """); Break;"

    End If

    ' (Se houver a opção de salvar e a validação for feita no cliente então chama a função que organiza as
    ' validações[ValidateForm criada na função ClientValidates])
    If Session(CurrentTable & "ClientValidation") And Session("RecordAction") > "" Then
      sobjShow.HTMLCR "  case """ & Buttons.SaveCaption & """ :"
      sobjShow.HTMLCR "  case """ & Buttons.SaveAddCaption & """ :"
      sobjShow.HTMLCR "  case """ & Buttons.SaveCopyCaption & """ :"
      sobjShow.HTMLCR "  return ValidateForm(); Break;"

    End If

    sobjShow.HTMLCR "  default : return true;"
    sobjShow.HTMLCR "}}"


    ' (Desabilita os botões apresentados no evento onClick)
    sobjShow.HTMLCR "function OnClickAndSubmit(){"
    If Session("RecordAction") > "" Then
      If Buttons.ShowSave Then
        sobjShow.HTMLCR "document.frmEditForm.btmSalva.disabled = true;"
      End If
      If Buttons.ShowSaveAdd Then
        sobjShow.HTMLCR "document.frmEditForm.btmSalvaInc.disabled = true;"
      End If
      If Buttons.ShowSaveCopy Then
        sobjShow.HTMLCR "document.frmEditForm.btmSalvaCop.disabled = true;"
      End If
      If Buttons.ShowCancel Then
        sobjShow.HTMLCR "document.frmEditForm.btmCancela.disabled = true;"
      End If
    Else
      If Buttons.ShowFirst Then
        sobjShow.HTMLCR "document.frmEditForm.btmPrimeiro.disabled = true;"
      End If
      If Buttons.ShowPrevious Then
        sobjShow.HTMLCR "document.frmEditForm.btmAnterior.disabled = true;"
      End If
      If (Session(CurrentTable & "EditOption") > 0) And (sparEditEditable = 1) Then
        If Buttons.ShowAdd Then
          sobjShow.HTMLCR "document.frmEditForm.btmInclui.disabled = true;"
        End If
        If Buttons.ShowCopy Then
          sobjShow.HTMLCR "document.frmEditForm.btmCopia.disabled = true;"
        End If
        If Buttons.ShowEdit Then
          sobjShow.HTMLCR "document.frmEditForm.btmAltera.disabled = true;"
        End If
        If Buttons.ShowDelete Then
          sobjShow.HTMLCR "document.frmEditForm.btmExclui.disabled = true;"
        End If

      End If
      If Buttons.ShowFind Then
        sobjShow.HTMLCR "document.frmEditForm.btmLocaliza.disabled = true;"
      End If
      If Buttons.ShowList Then
        sobjShow.HTMLCR "document.frmEditForm.btmListagem.disabled = true;"
      End If
      If Buttons.ShowNext Then
        sobjShow.HTMLCR "document.frmEditForm.btmProximo.disabled = true;"
      End If
      If Buttons.ShowLast Then
        sobjShow.HTMLCR "document.frmEditForm.btmUltimo.disabled = true;"
      End If

    End If
    sobjShow.HTMLCR "}"
  End If
  '--------------------

' (Faz a desabilitação criadas no FieldDisableValue, não está funcionando quando um campo de um tab
' desabilita um campo de outro tab)
  If Session("RecordAction") > "" Then
    sobjShow.HTMLCR "var blnChanged = false;"
  Dim i, j, k As Integer
  Dim blnFirst As Boolean
    blnFirst = True
  Dim strDisableFields As String
    BeginReadOperation Session("LastTab")

    Dim strList As String

    Do While IsThereOperation
      Select Case NextOperation
        Case ",", ";", "-"
        Case "="
          IsThereOperation

        Case Else
          If FindField(CurrentTable, LastOperation) Then
            If Session(CurrentField & "DisableNumber") > 0 Then

' (No primeiro campo que tiver FieldDisableValue inicializa a função)
              If blnFirst Then
                sobjShow.HTMLCR "function DisableFields(strField)"
                sobjShow.HTMLCR "{var i;"
                blnFirst = False

              End If

' (Se o campo atual for este então desabilite desta forma)
              sobjShow.HTMLCR "if(strField=='" & Session(CurrentField & "Field") & "'){"

' (Para cada desabilitação que pode ser feita a partir deste campo monta-a)
              For j = 1 To Session(CurrentField & "DisableNumber")
                Select Case Session(CurrentField & "EditOption")
' (Campo checkbox)
                  Case efFldTypeCheck
                    If sobjString.ToString(Session(CurrentField & ",1Value")) = sobjString.ToString(Session(CurrentField & "," & j & ",1DisableValue")) Xor Session(CurrentField & "," & j & "DisableCondition") Then
                      sobjShow.HTMLCR "if(document.frmEditForm." & Session(CurrentField & "Field") & ".checked)"

                    Else
                      sobjShow.HTMLCR "if(!document.frmEditForm." & Session(CurrentField & "Field") & ".checked)"

                    End If

' (Outros campos)
                  Case Else
                    Response.Write "if("
' (Para cada valor da desabilitação)
                    For k = 1 To Session(CurrentField & "," & j & "DisableValueNumber")
                      If k > 1 Then
' (Se for uma condição de desabilitação então utiliza o operador "ou" [||] senão utiliza o operador "e"[&])
                        If Session(CurrentField & "," & j & "DisableCondition") Then
                          Response.Write "||"

                        Else
                          Response.Write "&&"

                        End If

                      End If

' (Se o campo for do tipo radio então o seu acesso não é pelo valor e sim pela ordem do item selecionado)
                      If Session(CurrentField & "EditOption") = efFldTypeRadio Then
                        If Session(CurrentField & "," & j & "DisableCondition") Then
                          Response.Write "(document.frmEditForm." & Session(CurrentField & "Field") & ".item(" & Session(CurrentField & "," & j & "," & k & "DisableValue") & ").checked)"

                        Else
                          Response.Write "(!document.frmEditForm." & Session(CurrentField & "Field") & ".item(" & Session(CurrentField & "," & j & "," & k & "DisableValue") & ").checked)"

                        End If

' (Se o campo não for do tipo radio então o seu acesso é pelo valor)
                      Else
                        Response.Write "(document.frmEditForm." & Session(CurrentField & "Field") & ".value"
                        If Session(CurrentField & "," & j & "DisableCondition") Then
                          Response.Write "="

                        Else
                          Response.Write "!"

                        End If
                        Response.Write "=""" & Session(CurrentField & "," & j & "," & k & "DisableValue") & """)"

                      End If
                    Next
                    sobjShow.HTMLCR "){"

                End Select
' (Desabilitação propriamente dita)
              Dim strActualField As String
                strActualField = Session("CurrentField")
                FindField CurrentTable, Session(CurrentField & "," & j & "DisableField")
                sobjShow.HTMLCR "document.frmEditForm." & Session(CurrentField & "Field") & ".disabled=true;}"
                sobjShow.HTMLCR "else{"
                sobjShow.HTMLCR "document.frmEditForm." & Session(CurrentField & "Field") & ".disabled=false;}"

                CurrentField = strActualField

              Next
              sobjShow.HTMLCR "}"
            End If

          End If

      End Select

    Loop
    If Not blnFirst Then
       sobjShow.HTMLCR "}"

    End If

  End If

' (Testa se o formulário está em estado de edição)
  If Form.Enable Then
' (Ao serem alterados os campos guardam a presença desta alteração na variável blnChanged que avisa
' por exemplo ao botão de cancelar, que deve ser pedida a confirmação de cancelamento)
    sobjShow.HTMLCR "function FieldChanged(strField){"
    sobjShow.HTMLCR "blnChanged = true;"
    If Not blnFirst Then
      sobjShow.HTMLCR "DisableFields(strField);"

    End If
    sobjShow.HTMLCR "}"
  End If

' (Ao clicar num link de um tab, seu valor é enviado a um Input Hidden[hidEditor criado na função
' ShowScreen] e o form é submetido)
  sobjShow.HTMLCR "function TabCtrlsClick(intTab){"
  sobjShow.HTMLCR "  document.frmEditForm.hidEditor.value = intTab;"
  sobjShow.HTMLCR "  document.frmEditForm.submit();"
  sobjShow.HTMLCR "}"

' (Fecha o Script)
  Validate.JSEnd

End Sub

' (O formulário de edição é montado em uma tabela que distribue os campos. Esta operação é feita a
' partir da contagem de campos por linha de tabela e seleção do colspan de cada célula a partir desta
' contagem)

' (Inicializa a variável que conta as colunas por linha da tabela do formulário de edição)
Sub StartRow()
    sintNumCol = 1

End Sub

' (Faz a contagem das colunas por linha da tabela do formulário de edição)
Sub IncNumCol()
    sintNumCol = sintNumCol + 1

End Sub

Private Sub ColumnColSpan()
' (Muda os colspans das colunas da tabela do formulário de edição conforme a necessidade dos campos,
' descritas no Session(CurrentTable & "UnitOperations") em forma de lista de nº colunas na linha)

Dim intColSpan As Integer
' (Se for a primeira coluna da linha e houver algum item na lista de colspans)
  If sintNumCol = 1 And IsThereNumber Then
    intColSpan = Session(CurrentTable & "UnitOperationNumber") \ NextNumber

  Else
    intColSpan = Session(CurrentTable & "UnitOperationNumber") \ sintLastNumber

  End If

 If sintNumCol <= (Session(CurrentTable & "UnitOperationNumber") Mod sintLastNumber) Then
    intColSpan = intColSpan + 1

  End If

  If intColSpan > 1 Then
    Table.ColumnColSpan = intColSpan

  Else
    Table.ColumnColSpan = ""

  End If

End Sub

Private Sub ShowScreen()
  ' (Monta o formulário de edição)
  Dim blnNotStarting As Boolean, blnFirstEditable As Boolean
  Dim FontSize As Integer, i As Integer, intPos As Integer
  Dim strURL As String, strLastOperation As String
  'Criado (Ridai Govinda)
  Dim blnAux As Boolean, blnDoIdCheck As Boolean
  '--------------------

  ' (Os checkboxes aparecem normalmente como checks e não como radioboxes)
  NoCheckAsRadio

  sobjShow.BR
  sobjShow.Center

  ' (Desabillitado por não haver a Sub Execute no VB, como há no ASP)
  ' Implentar utilizando o controle ScriptControl da MS (Kelson/Ridai)
  If Session(CurrentTable & sparEditTab & "Task") > "" Then
    '    Execute Session(CurrentTable & sparEditTab & "Task")

  End If

  ' (Cria o estilo dos botões)
  sobjShow.HTMLCR "<STYLE> .BTM{font: 8pt Arial}</STYLE>"

  blnFirstEditable = True

  ' (Apresenta as mensagens de erro que estiverem no Message)
  If Message > "" Then
    Table.Padding = 2
    Table.BeginTable "65%", "", 1, True
    Table.BeginRow 2
    Table.BeginColumn
    Response.Write Message
    Table.EndColumn
    Table.EndRow
    Table.EndTable

    Message = ""

  End If

  strURL = Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionExeEdition & WhatWhereHow & GetTab

  'Criado pelo (Ruben Zevallos)
  'Inclusão de propriedades no <FORM>
  '<FORM ACTION="http://espm.totem.com.br/admin/Publicador.Asp?EA=h15&EW=11.1&EE=1&ET=0" METHOD="POST" NAME="frmEditForm" onSubmit="OnClickAndSubmit()">
  '<FORM method=post action="submit.asp" AUTOCOMPLETE="OFF" Style = "style" Class = "Class">

  Form.BeginForm strURL, "frmEditForm", , , "OnClickAndSubmit()", , , Class, StyleSheet, AutoComplete

  ' ==============================================================================
  ' De: Kelson
  ' Para: Ridai
  ' ------------------------------------------------------------------------------
  '
  ' Olha aqui aonde você pode alterar o código para entrar permitir o configuração
  ' das fontes do formulário.
  '
  ' O ideal seria que somente fosse utilizado este código quando o tamanho da
  ' fonte não fosse definida. Ou algo assemelhado, só não escrevo o código porque
  ' não tenho certeza da funcionalidade deste troço aqui a baixo.

  If Session(CurrentTable & "EditOption") <> 0 Then
    FontSize = 1

  Else
    FontSize = 2

  End If

  ' ==============================================================================

  Dim strMessage As String

' (Hidden que guarda o tab clicado)
  Form.Hidden "hidEditor"

' (Hidden que guarda o botão clicado)
  Form.Hidden "btmEditor"

' (Muda o título da tabela e a habilitação dos campos conforme a operação a ser feita com o
' registro)
  If Not Session("RecordAction") > "" Then
    strMessage = Session(CurrentTable & "Caption")
    Form.Enable = False

  ElseIf Session("RecordAction") = "1" Then
    strMessage = Session(CurrentTable & "Caption") & " - " & LanguageLocalizer(45)
    Form.Enable = True

  ElseIf Session("RecordAction") = "2" Then
    strMessage = Session(CurrentTable & "Caption") & " - " & LanguageLocalizer(46)
    Form.Enable = True

  End If

  'Criado/Alterado (Ridai Govinda)
  If UnitTitle > "" Then strMessage = UnitTitle
  '-----------------------------

' (Inicio da tabela externa)
  Table.Padding = 2
  Table.BeginTable "65%", strMessage, 1, True
  Table.Padding = 0

  blnNotStarting = False

' (Cabecalho)
  If ShowUnitUpperBar Then
    'Alterado (Ridai Govinda)
    Table.BeginRow 1, True
    Table.ColumnAlign = "LEFT"
    Table.BeginColumn
  ' (Indicador de existência de filtro nos dados)

    'Alterado (Ridai)
    If sparEditFilter > "" And ShowRecordCount Then
      Response.Write LanguageLocalizer(44) & "<BR>&nbsp;&nbsp;"

    End If

    'Alterado (Ridai)
    If ShowRecordCount And Session("RecordCount") > "" Then
      'Alterado (Flavio/Ridai)
      If Session("RecordAction") = "1" Then
              Response.Write LanguageLocalizer(47) & Session("RecordCount")
      Else
        Response.Write LanguageLocalizer(48) & IIf(Session("AbsolutePosition") > -1, Session("AbsolutePosition"), "...") & _
                       LanguageLocalizer(49) & Session("RecordCount")
      End If
      Response.Write "<BR>"

    Else
      sobjShow.Nbsp 2

    End If
    '----------------------
    Table.EndColumn
    Table.ColumnAlign = ""
    Table.EndRow
  End If
  '-------------------

  Session("LastTab") = Session(CurrentTable & "FieldUnit")

  ' (Localização de tab)
  For i = 1 To sparEditTab
    intPos = InStr(Session("LastTab"), "|")
    If intPos > 0 Then
      ' (Toma a parte da string que vai do inicio do tab ao fim desta)
      Session("LastTab") = Right(Session("LastTab"), Len(Session("LastTab")) - intPos)

    End If

  Next

  intPos = InStr(Session("LastTab"), "|")
  ' (Corta a string no inicio do próximo tab)
  If intPos > 0 Then
    Session("LastTab") = Left(Session("LastTab"), intPos - 1)

  End If

' (Monta a string com os nº de colunas por linha da tabela)
  TotalOperation Session("LastTab"), "UnitOperationNumber", True, "UnitOperations"

' (Se a opção de validação for no cliente então monta a validação)
  If Session(CurrentTable & "ClientValidation") And Session("RecordAction") > "" Then
    ClientValidates

  End If

' (Cria os JavaScripts genéricos)
  JavaScriptDefaults

' (Atribui a todos os inputs a Class 'BTM')
  Form.Class = "BTM"

  Table.BeginRow 1
  Table.BeginColumn

' (Montagem dos tabs)
  If HaveTabs Then
    Dim blnHeader As Boolean

    FontStyles

    BeginReadOperation Session(CurrentTable & "UnitTabs")

    Table.Spacing = 0
    Table.Padding = 1
    Table.BeginTable "100%"
    Table.BeginRow
    i = 0
    Do While IsThereOperation
      Select Case NextOperation
        Case "|"

        Case Else
          If sobjString.ToString(i) <> sobjString.ToString(sparEditTab) Then
            'Alterado (Ridai Govinda)
            Table.ColumnColor = Table.Style.HeaderColor
            '----------------------

          End If

          Table.BeginColumn
          sobjShow.Center
          Response.Write "<DIV CLASS=HL ID=Editor" & i & " onMouseOver=""makeCool(Editor" & i & ");"" onMouseOut=""makeNormal(Editor" & i & ");"" onClick=""TabCtrlsClick(" & i & ");"" STYLE=""cursor:hand"" STYLE=""cursor:hand""><B>" & LastOperation & "</B></DIV>"
          i = i + 1
          Table.EndColumn
          Table.ColumnColor = ""

      End Select

    Loop

    Table.EndRow
    Table.EndTable

  End If

' (Inicio da colocação dos campos)
  BeginReadNumbers Session(CurrentTable & "UnitOperations")

' (Se o FormUnit for iniciado com um sub-título[texto que segue o '=' no FormUnit] então monta sub-título)
  If Left(sstrText, 1) <> "=" Then
    StartRow
    Table.Padding = 2
    Table.BeginTable "100%"
    blnNotStarting = True
    ColumnColSpan
    Table.BeginRow FontSize
    Table.BeginColumn

  End If

' (Utiliza-se em todo o EditForm uma string de manipulação onde se testa os separadores de campos das
' strings FormUnit, FormList, etc)
' (Coloca o conteúdo do tab a ser apresentado na string de manipulação)
  BeginReadOperation Session("LastTab")

  Do While IsThereOperation
    Select Case NextOperation
' (Quebra de coluna)
      Case ","
        IncNumCol
        ColumnColSpan
        Table.EndColumn
        Table.BeginColumn

' (Quebra de linha)
      Case ";"
        StartRow
        Table.EndColumn
        Table.EndRow
        ColumnColSpan
        Table.BeginRow FontSize
        Table.BeginColumn

' (Linha separadora de formulário)
      Case "-"
        StartRow
        Table.EndColumn
        Table.EndRow
        Table.EndTable
        Table.EndColumn
        Table.EndRow
        Table.ColumnColSpan = ""
        Table.BeginRow 1
        Table.BeginColumn
        Table.BeginTable "100%"
        ColumnColSpan
        Table.BeginRow FontSize
        Table.BeginColumn

' (Sub-título)
      Case "="
        StartRow
        If blnNotStarting Then
          Table.EndColumn
          Table.EndRow
          Table.EndTable
          Table.EndColumn
          Table.EndRow

        End If

        blnNotStarting = True

        Table.BeginRow 2, True
        Table.ColumnAlign = "CENTER"
        Table.Column NextOperation
        Table.ColumnAlign = ""
        Table.EndRow
        Table.ColumnColSpan = ""
        Table.BeginRow 1
        Table.BeginColumn
        Table.BeginTable "100%"
        ColumnColSpan
        Table.BeginRow FontSize
        Table.BeginColumn
        NextOperation

' (Campo de edição ou HTML para formatação)
      Case Else
        If FindField(CurrentTable, LastOperation) Then
           If (Session(CurrentTable & "EditOption") > 0) And (sparEditEditable = 1) Then
              If Form.Enable And Session("RecordAction") > "" Then
                 Form.onChange = "FieldChanged('" & Session(CurrentField & "Field") & "')"

              End If

              'Alterado/Criado (Ridai Govinda)
              blnDoIdCheck = True

              If Session("RecordAction") = "2" And AllowIdentityUpdate Then blnDoIdCheck = False

              blnAux = Form.Enable

              If IdentityCheck(False) And blnDoIdCheck Then Form.Enable = False

              blnFirstEditable = ShowField(blnFirstEditable And Session("RecordAction") > "", False, Form.Enable, False)

              Form.Enable = blnAux
              '-----------------------------

           Else
             ShowFieldAsText

           End If

        Else
          If LastOperation = " " Then
            Response.Write "&nbsp;"

          ElseIf FindGrid(LastOperation) Then
            sobjShow.BR
            sobjShow.HTMLCR "<IFRAME FRAMEBORDER=NO HEIGHT=100 WIDTH=450 SRC=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionGrid & "&" & efQueryStrWhat & "=" & Session(CurrentGrid & "WhatList") & """></IFRAME>"

          Else
            Response.Write LastOperation

          End If

        End If

    End Select
  Loop

' (Roda a desabilitação de campos dos FieldDisableValue com os valores atuais)
  If Session("RecordAction") > "" Then
  Dim blnFirst As Boolean
    blnFirst = True

    BeginReadOperation Session("LastTab")

    Do While IsThereOperation
      Select Case NextOperation
        Case ",", ";", "-", "="

        Case Else
          If FindField(CurrentTable, LastOperation) Then
            If Session(CurrentField & "DisableNumber") > 0 Then
              If blnFirst Then
                sobjShow.HTMLCR "<SCRIPT LANGUAGE=""JavaScript"">"
                blnFirst = False

              End If
              sobjShow.HTMLCR "DisableFields('" & Session(CurrentField & "Field") & "');"

            End If
          End If

      End Select
    Loop

    If Not blnFirst Then
      sobjShow.HTMLCR "</SCRIPT>"
    End If

  End If

  Form.onChange = ""

  Table.EndColumn
  Table.EndRow
  Table.EndTable

  Table.EndColumn
  Table.EndRow

  Table.BeginRow
  Table.BeginColumn

  'Alterado (Ridai Govinda)
  If ShowButtons Then ShowBar
  '-----------------------

  Table.EndColumn
  Table.EndRow

  Table.EndTable

  Form.EndForm

  FPLPLink

End Sub

Private Sub ShowBar()
  'Criado (Ridai Govinda)
  Const lconCallBttnTest = "return ButtonsTest(this)"
  '--------------------
' (Cria a barra de botões do formulário de edição)
  sobjShow.Center

' (Se estiver em modo de visualização)
  If Not Session("RecordAction") > "" Then
' (Botão 'Primeiro Registro')
    If Buttons.ShowFirst Then
      If Session("AbsolutePosition") > 1 Then
        Form.Enable = True
        Form.onClick = lconCallBttnTest
        Rem Form.onClick = "return ButtonsTest(document.frmEditForm.btmPrimeiro)"

      Else
        Form.Enable = False

      End If
      'Alterado (Ridai Govinda)
      Form.Title = Buttons.FirstTitle
      Form.AccessKey = Buttons.FirstAccessKey

      If Buttons.FirstClass > "" Then Form.Class = Buttons.FirstClass

      Form.StyleSheet = Buttons.FirstStyleSheet
      If Buttons.FirstImage > "" Then
        Form.Image "btmPrimeiro", Buttons.FirstImage, Buttons.FirstCaption, Buttons.FirstCaption
      Else
        Form.Submit "btmPrimeiro", Buttons.FirstCaption
      End If
      '-----------------------
    End If

' (Botão 'Registro Anterior')
    If Buttons.ShowPrevious Then
      If Session("AbsolutePosition") > 1 Then
        Form.Enable = True
        Form.onClick = lconCallBttnTest
        Rem Form.onClick = "return ButtonsTest(document.frmEditForm.btmAnterior)"

      Else
        Form.Enable = False

      End If
      'Alterado (Ridai Govinda)
      Form.Title = Buttons.PreviousTitle
      Form.AccessKey = Buttons.PreviousAccessKey

      If Buttons.PreviousClass > "" Then Form.Class = Buttons.PreviousClass

      Form.StyleSheet = Buttons.PreviousStyleSheet

      If Buttons.PreviousImage > "" Then
        Form.Image "btmAnterior", Buttons.PreviousImage, Buttons.PreviousCaption, Buttons.PreviousCaption
      Else
        Form.Submit "btmAnterior", Buttons.PreviousCaption
      End If
      '-----------------------
    End If

  End If

  Form.Enable = True

  If sparEditEditable = 1 Then
    Select Case Session(CurrentTable & "EditOption")
      Case "1"
' (Se estiver em modo de visualização)
        If Not Session("RecordAction") > "" Then
' (Botão 'Incluir')
          If Buttons.ShowAdd Then
            Form.onClick = lconCallBttnTest
            Rem Form.onClick = "return ButtonsTest(document.frmEditForm.btmInclui)"

            'Alterado (Ridai Govinda)
            Form.Title = Buttons.AddTitle
            Form.AccessKey = Buttons.AddAccessKey

            If Buttons.AddClass > "" Then Form.Class = Buttons.AddClass

            Form.StyleSheet = Buttons.AddStyleSheet

            If Buttons.AddImage > "" Then
              Form.Image "btmInclui", Buttons.AddImage, Buttons.AddCaption, Buttons.AddCaption
            Else
              Form.Submit "btmInclui", Buttons.AddCaption
            End If
            '-----------------------
          End If
' (Botão 'Copiar')
        If Buttons.ShowCopy Then
            Form.onClick = lconCallBttnTest
            Rem Form.onClick = "return ButtonsTest(document.frmEditForm.btmCopia)"
            'Alterado (Ridai Govinda)
            Form.Title = Buttons.CopyTitle
            Form.AccessKey = Buttons.CopyAccessKey

            If Buttons.CopyClass > "" Then Form.Class = Buttons.CopyClass

            Form.StyleSheet = Buttons.CopyStyleSheet

            If Buttons.CopyImage > "" Then
              Form.Image "btmCopia", Buttons.CopyImage, Buttons.CopyCaption, Buttons.CopyCaption
            Else
              Form.Submit "btmCopia", Buttons.CopyCaption
            End If
            '-----------------------
          End If
' (Botão 'Alterar')
          If Buttons.ShowEdit Then
            Form.onClick = lconCallBttnTest
            Rem Form.onClick = "return ButtonsTest(document.frmEditForm.btmAltera)"
            'Alterado (Ridai Govinda)
            Form.Title = Buttons.EditTitle
            Form.AccessKey = Buttons.EditAccessKey

            If Buttons.EditClass > "" Then Form.Class = Buttons.EditClass

            Form.StyleSheet = Buttons.EditStyleSheet

            If Buttons.EditImage > "" Then
              Form.Submit "btmAltera", Buttons.EditCaption
              Form.Image "btmAltera", Buttons.EditImage, Buttons.EditCaption, Buttons.EditCaption
            Else
              Form.Submit "btmAltera", Buttons.EditCaption
            End If
            '-----------------------
          End If
' (Botão 'Excluir')
          If Buttons.ShowDelete Then
            Form.onClick = lconCallBttnTest
            Rem Form.onClick = "return ButtonsTest(document.frmEditForm.btmExclui)"
            'Alterado (Ridai Govinda)
            Form.Title = Buttons.DeleteTitle
            Form.AccessKey = Buttons.DeleteAccessKey

            If Buttons.DeleteClass > "" Then Form.Class = Buttons.DeleteClass

            Form.StyleSheet = Buttons.DeleteStyleSheet

            If Buttons.DeleteImage > "" Then
              Form.Image "btmExclui", Buttons.DeleteImage, Buttons.DeleteCaption, Buttons.DeleteCaption
            Else
              Form.Submit "btmExclui", Buttons.DeleteCaption
            End If
            '-----------------------
          End If
          Form.AccessKey = ""

        End If

' (Se estiver em modo de edição ou inclusão)
        If Session("RecordAction") > "" Then
' (Botão 'Salvar')
          If Buttons.ShowSave Then
            Form.onClick = lconCallBttnTest
            Rem Form.onClick = "return ButtonsTest(document.frmEditForm.btmSalva)"
            'Alterado (Ridai Govinda)
            Form.Title = Buttons.SaveTitle
            Form.AccessKey = Buttons.SaveAccessKey

            If Buttons.SaveClass > "" Then Form.Class = Buttons.SaveClass

            Form.StyleSheet = Buttons.SaveStyleSheet

            If Buttons.SaveImage > "" Then
              Form.Image "btmSalva", Buttons.SaveImage, Buttons.SaveCaption, Buttons.SaveCaption
            Else
              Form.Submit "btmSalva", Buttons.SaveCaption
            End If
            '-----------------------
          End If
' (Botão 'Salvar e Incluir')
          If Buttons.ShowSaveAdd Then
            Form.onClick = lconCallBttnTest
            Rem Form.onClick = "return ButtonsTest(document.frmEditForm.btmSalvaInc)"
            'Alterado (Ridai Govinda)
            Form.Title = Buttons.SaveAddTitle
            Form.AccessKey = Buttons.SaveAddAccessKey

            If Buttons.SaveAddClass > "" Then Form.Class = Buttons.SaveAddClass

            Form.StyleSheet = Buttons.SaveAddStyleSheet

            If Buttons.SaveAddImage > "" Then
              Form.Image "btmSalvaInc", Buttons.SaveAddImage, Buttons.SaveAddCaption, Buttons.SaveAddCaption
            Else
              Form.Submit "btmSalvaInc", Buttons.SaveAddCaption
            End If
            '-----------------------
          End If
' (Botão 'Salvar e Copiar')
          If Buttons.ShowSaveCopy Then
            Form.onClick = lconCallBttnTest
            Rem Form.onClick = "return ButtonsTest(document.frmEditForm.btmSalvaCop)"
            'Alterado (Ridai Govinda)
            Form.Title = Buttons.SaveCopyTitle
            Form.AccessKey = Buttons.SaveCopyAccessKey

            If Buttons.SaveCopyClass > "" Then Form.Class = Buttons.SaveCopyClass

            Form.StyleSheet = Buttons.SaveCopyStyleSheet

            If Buttons.SaveCopyImage > "" Then
              Form.Image "btmSalvaCop", Buttons.SaveCopyImage, Buttons.SaveCopyCaption, Buttons.SaveCopyCaption
            Else
              Form.Submit "btmSalvaCop", Buttons.SaveCopyCaption
            End If
            '-----------------------
          End If
' (Botão 'Cancelar')
          If Buttons.ShowCancel Then
            Form.AccessKey = "C"
            Form.onClick = lconCallBttnTest
            Rem Form.onClick = "return ButtonsTest(document.frmEditForm.btmCancela)"
            Form.Title = "Cancela todas as alterações efetuadas"
            'Alterado (Ridai Govinda)
            Form.Title = Buttons.CancelTitle
            Form.AccessKey = Buttons.CancelAccessKey

            If Buttons.CancelClass > "" Then Form.Class = Buttons.CancelClass

            Form.StyleSheet = Buttons.CancelStyleSheet

            If Buttons.CancelImage > "" Then
              Form.Image "btmCancela", Buttons.CancelImage, Buttons.CancelCaption, Buttons.CancelCaption
            Else
              Form.Submit "btmCancela", Buttons.CancelCaption
            End If
            '-----------------------
          End If
          Form.AccessKey = ""

        End If

    End Select
  End If

' (Se estiver em modo de visualização)
  If Not Session("RecordAction") > "" Then
' (Botão 'Localizar')
    If Buttons.ShowFind Then
      Form.onClick = lconCallBttnTest
      Rem Form.onClick = "return ButtonsTest(document.frmEditForm.btmLocaliza)"
      'Alterado (Ridai Govinda)
      Form.Title = Buttons.FindTitle
      Form.AccessKey = Buttons.FindAccessKey

      If Buttons.FindClass > "" Then Form.Class = Buttons.FindClass

      Form.StyleSheet = Buttons.FindStyleSheet

      If Buttons.FindImage > "" Then
        Form.Image "btmLocaliza", Buttons.FindImage, Buttons.FindCaption, Buttons.FindCaption
      Else
        Form.Submit "btmLocaliza", Buttons.FindCaption
      End If
      '-----------------------
    End If
' (Botão 'Listagem')
    If Buttons.ShowList Then
      Form.onClick = lconCallBttnTest
      Rem Form.onClick = "return ButtonsTest(document.frmEditForm.btmListagem)"
      'Alterado (Ridai Govinda)
      Form.Title = Buttons.ListTitle
      Form.AccessKey = Buttons.ListAccessKey

      If Buttons.ListClass > "" Then Form.Class = Buttons.ListClass

      Form.StyleSheet = Buttons.ListStyleSheet

      If Buttons.ListImage > "" Then
        Form.Image "btmListagem", Buttons.ListImage, Buttons.ListCaption, Buttons.ListCaption
      Else
        Form.Submit "btmListagem", Buttons.ListCaption
      End If
      '----------------------
    End If
' (Botão 'Próximo Registro')
    If Buttons.ShowNext Then
      If Session("AbsolutePosition") = Session("RecordCount") Then
        Form.Enable = False

      Else
        Form.Enable = True
        Form.onClick = lconCallBttnTest
        Rem Form.onClick = "return ButtonsTest(document.frmEditForm.btmProximo)"

      End If
      'Alterado (Ridai Govinda)
      Form.Title = Buttons.NextTitle
      Form.AccessKey = Buttons.NextAccessKey

      If Buttons.NextClass > "" Then Form.Class = Buttons.NextClass

      Form.StyleSheet = Buttons.NextStyleSheet

      If Buttons.NextImage > "" Then
        Form.Image "btmProximo", Buttons.NextImage, Buttons.NextCaption, Buttons.NextCaption
      Else
        Form.Submit "btmProximo", Buttons.NextCaption
      End If
      '-----------------------
    End If
' (Botão 'Último Registro')
    If Buttons.ShowLast Then
      If Session("AbsolutePosition") = Session("RecordCount") Then
        Form.Enable = False

      Else
        Form.Enable = True
        Form.onClick = lconCallBttnTest
        Rem Form.onClick = "return ButtonsTest(document.frmEditForm.btmUltimo)"

      End If
      'Alterado (Ridai Govinda)
      Form.Title = Buttons.LastTitle
      Form.AccessKey = Buttons.LastAccessKey

      If Buttons.LastClass > "" Then Form.Class = Buttons.LastClass

      Form.StyleSheet = Buttons.LastStyleSheet

      If Buttons.LastImage > "" Then
        Form.Image "btmUltimo", Buttons.LastImage, Buttons.LastCaption, Buttons.LastCaption
      Else
        Form.Submit "btmUltimo", Buttons.LastCaption
      End If
      '-----------------------
    End If
  End If
  Form.onClick = ""
  Form.Enable = True

End Sub

Private Function EditSizeToShow() As Integer
' (Função que retorna o tamanho de um campo para visualização)
  If Session(CurrentField & "ShowSize") > "" Then
    EditSizeToShow = Session(CurrentField & "ShowSize")

  Else
    If EditSize <= 80 Then
      EditSizeToShow = EditSize

    Else
      EditSizeToShow = 80

    End If

  End If

End Function

Private Function EditSize() As Long
' (Função que retorna o tamanho de um campo para visualização)

' (Por tipo)
  Select Case Session(CurrentField & "Type")
    Case efDataTypeInt, efDataTypeTinyInt, efDataTypeSmallInt
      EditSize = Len(sobjString.ToString(256 ^ Session(CurrentField & "Size"))) - 1

    Case efDataTypeFloat
      EditSize = Len(sobjString.ToString(256 ^ Session(CurrentField & "Size"))) + 1

    Case efDataTypeMoney
      EditSize = Len(sobjString.ToString(256 ^ Session(CurrentField & "Size"))) + 5

    Case efDataTypeDateTime
      Select Case Session(CurrentField & "EditOption")
        Case efValOptTime
          EditSize = 5

        Case Else
          EditSize = 10

      End Select

    Case Else
      EditSize = CLng(Session(CurrentField & "Size"))

  End Select

End Function

Private Function EditSizeToGet() As Integer
' (Função que retorna o nº máximo caracteres de um campo)
  If Session(CurrentField & "GetSize") > "" Then
    EditSizeToGet = Session(CurrentField & "GetSize")

  Else
    EditSizeToGet = EditSize

  End If

End Function

Private Sub ShowFieldImages()
  Dim strPopupLink As String
' (Apresenta as imagens de campo requerido e erro de preenchimento)

' (Campo requerido)
  If Session(CurrentField & "RequireMessage") > "" Or Not Session(CurrentField & "Null") Then
    Response.Write " "
    sobjShow.Image Session("htDefaultHiperToolsWorkPath") & "/CR.gif", "", LanguageLocalizer(50), "", "", 5, 5

  End If

  If Session(CurrentField & "LookupFieldPopup") Then

    Response.Write " "
    sobjShow.Image Session("htDefaultHiperToolsWorkPath") & "/img/icone/externallink.gif", "", Session(CurrentField & "LookupFieldPopupTitle"), strPopupLink, "", , , , "_blank"

  End If

' (Erro de preenchimento)
  If Session(CurrentField & "ValidationError") > "" Then
    Response.Write " "
    sobjShow.Image Session("htDefaultHiperToolsWorkPath") & "/CE.gif", "", Session(CurrentField & "ValidationError"), "", "", 5, 5
    Session(CurrentField & "ValidationError") = ""

  End If

End Sub

Private Sub ShowCaption(ByVal blnForm As Boolean)
' (Apresenta o rótulo do CurrentField, se estiver em modo de edição coloca sublinhado na tecla
' de acesso ao campo)
Dim i As Integer

  i = InStr(Session(CurrentField & "Caption"), "&")

  Rem Session(CurrentField & "CaptionVAlign") = Value


  If i > 0 Then
    If blnForm Then
      Response.Write FormatShowFieldCaption(Left(Session(CurrentField & "Caption"), i - 1) & "<U>" & Mid(Session(CurrentField & "Caption"), i + 1, 1) & "</U>" & Mid(Session(CurrentField & "Caption"), i + 2))
      Form.AccessKey = sobjString.ChangeChar(Mid(Session(CurrentField & "Caption"), i + 1, 1))

    Else
      Response.Write Left(Session(CurrentField & "Caption"), i - 1) & Mid(Session(CurrentField & "Caption"), i + 1)

    End If

  Else
    If blnForm Then
      Response.Write FormatShowFieldCaption(Session(CurrentField & "Caption"))

    Else
      Response.Write Session(CurrentField & "Caption")

    End If

  End If

End Sub

' (Flag que indica se um campo checkbox será apresentado como radiobox)
Private Sub CheckAsRadio()
  sstrCheckAsRadio = "1"

End Sub
Private Sub NoCheckAsRadio()
  sstrCheckAsRadio = ""

End Sub

Private Function FileSize(ByVal FileName As String) As String
' (Transforma o tamanho do arquivo de bytes para a unidade mais apropriada)
Dim objFile As Object
Dim objFileSystem As Object

    Set objFileSystem = Server.CreateObject("Scripting.FileSystemObject")
    If objFileSystem.FileExists(FileName) Then
      Set objFile = objFileSystem.GetFile(FileName)
      FileSize = " (" & FileSizeFormat(objFile.Size) & ")"

    Else
      FileSize = " (" & LanguageLocalizer(51) & ")"

    End If

End Function

Private Function FileSizeFormat(ByVal dblSize As Double) As String
' (Transforma o tamanho do arquivo de bytes para a unidade mais apropriada)
Dim i As Byte
Dim strSizeUnit As String

  i = 0
  Do While dblSize > 512
    dblSize = dblSize / 1024
    i = i + 1

  Loop

  Select Case i
    Case 0
      strSizeUnit = "bytes"

    Case 1
      strSizeUnit = "KB"

    Case 2
      strSizeUnit = "MB"

    Case 3
      strSizeUnit = "GB"

    Case 4
      strSizeUnit = "TB"

    Case 5
      strSizeUnit = "HB"

    Case 6
      strSizeUnit = "EB"

  End Select
  FileSizeFormat = sobjString.ZTINumber(Round(dblSize, 2)) & " " & strSizeUnit

End Function

Private Function ShowField(ByVal blnFocus As Boolean, ByVal blnFindField As Boolean, ByVal blnEnable As Boolean, ByVal blnEmpty As Boolean) As Boolean
  ' (Apresenta um campo de formulário, parametro: Algum campo já possui o foco ?, O campo vai
  ' ser apresentado sem o seu rótulo ?, Está habilitado ?, Deve aparecer vazio ?)
  Dim sql As String, strFName As String
  Dim i As Integer
  Dim Value
  Dim blnFirst As Boolean
  Dim arrAux As Variant
  Dim intCurrentField As Long

  ' (Se o campo permite edição)
   If Session(CurrentField & "ShowOption") > -1 Then
    Form.Title = Session(CurrentField & "Hint")

    'Incluido Ruben Zevallos Jr.
    Form.AutoComplete = Session(CurrentField & "AutoComplete")
    Form.ReadOnly = Session(CurrentField & "ReadOnly")

    If Not IsNull(Session(CurrentField & "Disabled")) And Not IsEmpty(Session(CurrentField & "Disabled")) Then
      Form.Enable = Session(CurrentField & "Disabled") = False

    End If

    Form.TabIndex = Session(CurrentField & "TabIndex")
    Form.StyleSheet = Session(CurrentField & "StyleSheet")

    If Not IsNull(Session(CurrentField & "Class")) And Not IsEmpty(Session(CurrentField & "Class")) Then
      Form.Class = Session(CurrentField & "Class")

    End If

    If blnEmpty Then
      Value = ""

    Else
      If Not Session(CurrentField & "Value") > "" Then
        Value = ""

      Else
        If Session(CurrentField & "EditOption") = efFldTypeText Then
          Value = FormatVariable(Session(CurrentField & "Value"), Session(CurrentField & "Type"))

        Else
          Value = Session(CurrentField & "Value")

        End If
      End If
    End If

    Select Case Session(CurrentField & "EditOption")
' (Campos tipo Text, HTTP, E-Mail)
      Case efFldTypeText, efFldTypeHTTP, efFldTypeEMail
        If Not blnFindField Then
          ShowCaption True
          ShowFieldImages
          sobjShow.BR
          Form.text Session(CurrentField & "Field"), EditSizeToShow, sobjString.ToString(Value), EditSizeToGet

        Else
          Form.text Session(CurrentField & "Field"), 30

        End If

' (Campo tipo Cor)
      Case efFldTypeColor
        If Not blnFindField Then
          ShowCaption True

          'Criado (Ridai Govinda) (Ativar depois para se poder apagar o valor do campo)
          'sobjShow.BR
          'Form.Text Session(CurrentField & "Field"), EditSizeToShow, _
          'sobjString.ToString(Value), EditSizeToGet
          '---------------------
          Form.Palette Session(CurrentField & "Field"), sobjString.ToString(Value)

        Else
          Form.text Session(CurrentField & "Field"), 30

        End If

' (Campo tipo Atualização)
      Case efFldTypeAtualization, efFldTypeDateOfUpdate, efFldTypeDateOfDelete, efFldTypeDateOfInsert
        If Not blnFindField Then
          ShowCaption True
          ShowFieldImages
          sobjShow.BR

        End If
        Response.Write Value

' (Campo tipo Arquivo)
      Case efFldTypeFile
        'Verifica o estado do Session, se estiver vazio não apresenta nenhum link
        'Alterado (Ridai Govinda)

        If Session("RecordAction") > "" Then
          Rem Response.Write "<A HREF=""\SendFile.asp?O=1&Folder=" & Server.URLEncode(Session(CurrentField & "FileRoot")) & "&Field=" & Session(CurrentField & "Field") & "&FileName=" & Session(CurrentField & "FileNamePrefix") & Session.SessionID
          Response.Write "<A HREF=""" & Initializer.ScriptName & "?" & efQueryStrAction & "=" & efQSActionGetFile & "&" & _
            efQueryStrFileOpt & "=" & efFileOptChoose & "&" & _
            efQueryStrFileFolder & "=" & Server.URLEncode(Session(CurrentField & "FileRoot")) & "&" & _
            efQueryStrFileField & "=" & Session(CurrentField & "Field") & "&"

          If Session("RecordAction") = 1 Then
            Response.Write efQueryStrFileName & "=" & Session(CurrentField & "FileNamePrefix") & "___EF"
          Else
              If Session(CurrentField & "Value") > "" Then
                Response.Write efQueryStrFileName & "="
                arrAux = Split(Session(CurrentField & "Value"), ".")

                strFName = ""
                If UBound(arrAux) > -1 Then
                  For i = 0 To UBound(arrAux) - 1
                    If strFName > "" Then strFName = strFName & "."
                    strFName = strFName & arrAux(i)
                  Next
                End If

                Response.Write strFName
              Else
                Response.Write efQueryStrFileName & "=" & Session(CurrentField & "FileNamePrefix")
                For i = 1 To Session(CurrentTable & "FieldNumber")
                  If Session(CurrentTable & i & "IsKey") Then
                    Response.Write Session(CurrentTable & i & "Value")
                    Exit For
                  End If
                Next

              End If
          End If

          Response.Write WhatList

          'Criado (Ridai Govinda)
          'sobjShow.Nbsp 2
          'Form.Text Session(CurrentField & "Field"), EditSizeToShow, _
          'sobjString.ToString(Value), EditSizeToGet
          '---------------------

        End If

'        For i = 1 To Session(CurrentTable & "FieldNumber")
'          If Session(CurrentTable & i & "IsKey") Then
'             Response.Write Session(CurrentTable & i & "Value")
'          End If
'        Next
        'Verifica o estado do Session, se estiver vazio não apresenta nenhum link
        If Session("RecordAction") > "" Then
            Response.Write """ TARGET=""__top"" CLASS=""SS"">"
        End If

        ShowCaption True
        ShowFieldImages

        'Verifica o estado do Session, se estiver vazio não apresenta nenhum link
        If Session("RecordAction") > "" Then
            Response.Write "</A>"
        End If

        If Value > "" Then
          sobjShow.HTMLCR " - <A HREF=""" & Session(CurrentField & "FileRoot") & "/" & Value & """ TARGET=__blank>" & Value & FileSize(Replace(Initializer.SiteRootDir & Session(CurrentField & "FileRoot") & "\" & Value, "\\", "\")) & "</A>"

        End If

' (Campo tipo Imagem)
      Case efFldTypeImage
        If Not blnFindField Then
          'Alterado (Ridai Govinda)
          Rem sobjShow.JSAlert "#" & Session("RecordAction") & "#"
          Rem sobjShow.JSAlert "#" & (Session("RecordAction") > "") & "#"

          If Session(CurrentField & "ImageNamePrefix") > "" And Session("RecordAction") > "" Then
            Rem Response.Write "<A HREF=""\SendFile.asp?O=1&Folder=" & Server.URLEncode(Session(CurrentField & "ImageRoot")) & "&Field=" & Session(CurrentField & "Field") & "&FileName=" & Session(CurrentField & "ImageNamePrefix") & Session.SessionID
            Response.Write "<A HREF=""" & Initializer.ScriptName & "?" & efQueryStrAction & "=" & efQSActionGetFile & "&" & _
              efQueryStrFileOpt & "=" & efFileOptChoose & "&" & _
              efQueryStrFileFolder & "=" & Server.URLEncode(Session(CurrentField & "ImageRoot")) & "&" & _
              efQueryStrFileField & "=" & Session(CurrentField & "Field") & "&"

            If Session("RecordAction") = 1 Then
              Response.Write efQueryStrFileName & "=" & Session(CurrentField & "ImageNamePrefix") & "___EF"
            Else
              If Session(CurrentField & "Value") > "" Then
                Response.Write efQueryStrFileName & "="
                arrAux = Split(Session(CurrentField & "Value"), ".")

                strFName = ""
                If UBound(arrAux) > -1 Then
                  For i = 0 To UBound(arrAux) - 1
                    If strFName > "" Then strFName = strFName & "."
                    strFName = strFName & arrAux(i)
                  Next
                End If

                Response.Write strFName

              Else
                Response.Write efQueryStrFileName & "=" & Session(CurrentField & "ImageNamePrefix")
                For i = 1 To Session(CurrentTable & "FieldNumber")
                  If Session(CurrentTable & i & "IsKey") Then
                    Response.Write Session(CurrentTable & i & "Value")
                    Exit For
                  End If
                Next

              End If
            End If
            Response.Write WhatList

            Response.Write """ target=__top Class=""SS"">"
            ShowCaption True
            Response.Write "</A>"

          Else
            ShowCaption True

          End If

          ShowFieldImages

          'Criado (Ridai Govinda)
          'sobjShow.BR
          'Form.Text Session(CurrentField & "Field"), EditSizeToShow, _
          'sobjString.ToString(Value), EditSizeToGet
          '---------------------
          sobjShow.BR

        End If
        '-------------------------

        If Value > "" Then
          Rem sobjShow.Image "/" & Session(CurrentField & "ImageRoot") & "/" & Value, "", "/" & Session(CurrentField & "ImageRoot") & "/" & Value, "'/" & Session(CurrentField & "ImageRoot") & "/" & Value & "' TARGET=__blank", "", Session(CurrentField & "ImageWidth"), ""
          Rem --- Alterado ---

          sobjShow.Image Session(CurrentField & "ImageRoot") & "/" & Value, , Session(CurrentField & "ImageRoot") & "/" & Value, Session(CurrentField & "ImageRoot") & "/" & Value, "", IIf(Session(CurrentField & "ImageWidth") > "", Session(CurrentField & "ImageWidth"), ""), "", 0, "_blank"

        End If

' (Campo tipo Senha)
      Case efFldTypePassword
        ShowCaption True
        ShowFieldImages
        sobjShow.BR
        Form.Password Session(CurrentField & "Field"), EditSizeToShow

' (Campo tipo Caixa de preenchimento)
      Case efFldTypeTextArea, efFldTypeHTML, efFldTypeTextAreaCounter
        If Not blnFindField Then
          ShowCaption True
          ShowFieldImages
          sobjShow.BR

        End If

        Form.TextArea Session(CurrentField & "Field"), CInt(Session(CurrentField & "TextAreaHeight")), CInt(EditSizeToShow), sobjString.ToString(Value), Session(CurrentField & "TextAreaAlign"), Session(CurrentField & "TextAreaWrap")

' (Campo tipo UF)
      Case efFldTypeUF
        If Not blnFindField Then
          ShowCaption True
          ShowFieldImages
          sobjShow.BR

        End If
          If Not blnEnable Then
            Form.SelectBegin Session(CurrentField & "Field"), 1, False
            If Value > "" Then
              Form.SelectOption Value, Value, True

            Else
              Form.SelectOption "", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", True

            End If
            Form.SelectEnd

          Else
            Form.SelectState Session(CurrentField & "Field"), 1, False, sobjString.ToString(Value)

          End If

' (Campo tipo Lookup)
      Case efFldTypeLookup
' (Testa se o lookup está definido)
        If (Session(CurrentField & "LookupTable") > "") And (Session(CurrentField & "LookupField") > "") _
        And (Session(CurrentField & "LookupFieldText") > "") Then

          If Not blnFindField Then
            ShowCaption True
            ShowFieldImages
            sobjShow.BR

          End If

' (Se o campo não estiver habilitado então apresenta apenas o item selecionado)
          If Not blnEnable Then
            If Value > "" Then
              'Alterado (Ridai Govinda)
              sql = "SELECT " & Session(CurrentField & "LookupField") & ", " & _
                    Session(CurrentField & "LookupFieldText") & " " & _
                    " FROM " & Session(CurrentField & "LookupTable")
              sql = sql & " WHERE " & Session(CurrentField & "LookupField") & " = " & FormatForSQL(Value, Session(CurrentField & "Type"))
              '-----------------------

            End If

' (Senão monta o select completo)
          Else
            'Alterado (Ridai Govinda)

            sql = "SELECT " & Session(CurrentField & "LookupField") & ", " & _
                  Session(CurrentField & "LookupFieldText") & " " & _
                  " FROM " & Session(CurrentField & "LookupTable")
            '-----------------------

            ' (Monta filtro por valor)
            If Session(CurrentField & "LookupFilterValue") > "" Then
              sql = sql & " WHERE " & Session(CurrentField & "DependentField") & " = "
              sql = sql & FormatForSQL(Session(CurrentField & "LookupFilterValue"), Session(CurrentField & "FilterFieldType"))

                ' (Monta filtro por dependência de um campo, ex.:num lookup de estados deve haver um filtro
                ' destes estados pelo país)
            Else
              If Session(CurrentField & "LookupFilterField") > "" Then
                intCurrentField = Session("CurrentField")
                If FindField(CurrentTable, Session(CurrentField & "LookupFilterField")) Then
                  sql = sql & " WHERE " & Session(CurrentTable & intCurrentField & "DependentField") & " = "
                  sql = sql & FormatForSQL(Session(CurrentField & "Value"), Session(CurrentField & "Type"))

                End If
                Session("CurrentField") = intCurrentField
              End If
            End If

            If sql > "" Then
              sql = sql & " ORDER BY " & Session(CurrentField & "LookupFieldText")

            End If
          End If

'          If sql > "" Then
'            sobjCMD.CommandText = sql
'            sobjCMD.CommandType = adCmdText
'
'          End If

          Form.SelectBegin Session(CurrentField & "Field"), 1, False
          If (Not Value > "" Or IsNull(Value)) Then
            Form.SelectOption "", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", True

          Else
            If blnEnable Then
              Form.SelectOption "", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", False

            End If

          End If

          If ((Value > "" And Not IsNull(Value)) Or blnEnable) And sql > "" And Not Session(CurrentField & "LookupFieldPopup") Then

            'Alterado/Criado (Ridai)
            Set CurrConnection = sobjConn

            If ShowSQLQuery Then ShowCurrentSQL sql, "ShowField", adOpenForwardOnly, adLockReadOnly

            On Error Resume Next

            sobjRS.Open sql, sobjConn, adOpenForwardOnly, adLockReadOnly

            If Err.number <> 0 Then
              If DebugMode Then
                ShowSQLErrors sql, "ShowField", adOpenForwardOnly, adLockReadOnly
              Else
                ShowDefaultError
              End If
              Err.Clear
              Exit Function

            End If
            On Error GoTo 0
            '----

            Do While Not sobjRS.EOF
              Form.SelectOption sobjRS(Session(CurrentField & "LookupField")), _
                                    sobjRS(Session(CurrentField & "LookupFieldText")), _
                                    (sobjString.ToString(sobjRS(Session(CurrentField & "LookupField"))) = sobjString.ToString(Value))
              sobjRS.MoveNext

            Loop
            sobjRS.Close
            Form.SelectEnd

          End If
        Else
          ShowError efErrorNoLookUp, "EditForm", "ShowField", sobjString.FormatText(LanguageLocalizer(8), Session(CurrentField & "Field"), CurrentTable)

        End If

' (Campo tipo Checkbox)
      Case efFldTypeCheck
        If Not sstrCheckAsRadio > "" Then
          Form.CheckBox Session(CurrentField & "Field"), Session(CurrentField & ",1Value"), _
          "left", BitBln(sobjString.ToString(Value) = sobjString.ToString(Session(CurrentField & ",1Value")))
          If Not blnFindField Then
            ShowCaption True
            ShowFieldImages

          End If

        Else
          Response.Write Session(CurrentField & ",1Caption")
          Form.Radio Session(CurrentField & "Field"), Session(CurrentField & ",1Value"), "left", BitBln(sobjString.ToString(Value) = sobjString.ToString(Session(CurrentField & ",1Value")))
          Response.Write " "
          Response.Write Session(CurrentField & ",2Caption")
          Form.Radio Session(CurrentField & "Field"), Session(CurrentField & ",2Value"), "left", BitBln(sobjString.ToString(Value) = sobjString.ToString(Session(CurrentField & ",2Value")))

        End If

' (Campo tipo Radio)
      Case efFldTypeRadio
        If Not blnFindField Then
          ShowCaption True
          ShowFieldImages

        End If
      Dim j
' (Monta uma tabela que organiza os radios)
        Table.BeginTable "100%", "", 1
        Table.ColumnVAlign = "Middle"
        Table.BeginRow 1
        blnFirst = True
' (Para cada valor do radio)
        For i = 1 To Session(CurrentField & "ValueNumber")
' (Faz as quebras de linha conforme o nº de itens por coluna definido no
' Session(CurrentField & "RadioColumns"))
          If (i - 1) Mod Session(CurrentField & "RadioColumns") = 0 Then
            If Not blnFirst Then
              Table.EndColumn
              Table.EndRow
              Table.BeginRow 1

            End If
            blnFirst = False
          End If

' (Coloca o input Radio)
          Table.ColumnWidth = 10
          Table.BeginColumn
          Form.Radio Session(CurrentField & "Field"), Session(CurrentField & "," & i & "Value"), "left", BitBln(sobjString.ToString(Value) = sobjString.ToString(Session(CurrentField & "," & i & "Value")))
          Table.EndColumn

' (Coloca o texto do item)
          Table.ColumnWidth = ""
          Table.BeginColumn
          Response.Write "&nbsp;" & Session(CurrentField & "," & i & "Caption")
          Table.EndColumn

        Next
        Table.EndRow
        Table.EndTable

' (Campo tipo Select)
      Case efFldTypeSelect
        If Not blnFindField Then
          ShowCaption True
          ShowFieldImages
          sobjShow.BR

        End If

        Form.SelectBegin Session(CurrentField & "Field"), 1, False
        If Not Value > "" Then
          Form.SelectOption "", "", True

        End If
        For i = 1 To Session(CurrentField & "ValueNumber")
          Form.SelectOption Session(CurrentField & "," & i & "Value"), Session(CurrentField & "," & i & "Caption"), _
                                           (sobjString.ToString(Value) = sobjString.ToString(Session(CurrentField & "," & i & "Value")))

        Next
        Form.SelectEnd

' (Campo tipo Data separada em Dia, Mês e Ano)
      Case efFldTypeSeparateDate
        If Not blnFindField Then
          ShowCaption True
          ShowFieldImages
          sobjShow.BR

        End If

' (Dia)
        Form.SelectBegin Session(CurrentField & "Field") & "DayValue", 1, False
        For i = 1 To 31
          Form.SelectOption sobjString.LeadingZeroes(i, 2), sobjString.LeadingZeroes(i, 2), i = Day(Value)
        Next
        Form.SelectEnd

        Response.Write "<B>/</B>"

' (Mês)
        Form.SelectBegin Session(CurrentField & "Field") & "MonthValue", 1, False
        For i = 1 To 12
          Form.SelectOption sobjString.LeadingZeroes(i, 2), sobjString.LeadingZeroes(i, 2), i = Month(Value)
        Next
        Form.SelectEnd

        Response.Write "<B>/</B>"

' (Ano)
' (As variáveis BeginYear e EndYear indicam o primeiro e o último ano do select)
      Dim BeginYear, EndYear
        BeginYear = Session(CurrentField & "BeginYear")
        EndYear = Session(CurrentField & "EndYear")
        If Not BeginYear > "" Then
          BeginYear = Year(Now) - 20

        End If
        If Not EndYear > "" Then
          EndYear = Year(Now) + 20

        End If

        Form.SelectBegin Session(CurrentField & "Field") & "YearValue", 1, False
        For i = BeginYear To EndYear
          Form.SelectOption i, i, i = Year(Value)
        Next
        Form.SelectEnd

    End Select

' (Põe o foco no primeiro input da tela)
    If blnFocus Then
      Select Case Session(CurrentField & "EditOption")
        Case efFldTypeImage

        Case efFldTypeRadio
          blnFocus = False
          sobjShow.HTMLCR "<SCRIPT LANGUAGE=""JavaScript"">"
          sobjShow.HTMLCR "document.frmEditForm." & Session(CurrentField & "Field") & "(0).focus();"
          sobjShow.HTMLCR "</SCRIPT>"

        Case Else
          'Alterado (Ridai Govinda)
          If Not IdentityCheck(False) Then
            blnFocus = False
            sobjShow.HTMLCR "<SCRIPT LANGUAGE=""JavaScript"">"
            sobjShow.HTMLCR "document.frmEditForm." & Session(CurrentField & "Field") & ".focus();"
            sobjShow.HTMLCR "</SCRIPT>"

          End If
          '----------------------

      End Select

    End If

' (Se o campo não permite edição apresenta o valor como texto)
  Else
    ShowFieldAsText

  End If

' (Retorna Se algum campo já recebeu o foco)
  ShowField = blnFocus

End Function

Private Sub ShowFieldLookupPopup()
    'Criado por Ridai Govinda & Kelson Medeiros


End Sub

Private Function BitBln(ByVal blnValor As Boolean) As Byte
' (Troca um valor booleano para 0 ou 1, parametro: Valor booleano)
  If blnValor Then
    BitBln = 1

  Else
    BitBln = 0

  End If

End Function

Private Sub ShowFieldAsText()
' (Apresenta um em forma textual)
Dim sql As String
Dim i As Integer
Dim Value

  If FindField(CurrentTable, LastOperation) Then
    If Not Session(CurrentField & "Value") > "" Then
      Value = ""

    Else
      Value = FormatVariable(Session(CurrentField & "Value"), Session(CurrentField & "Type"))

    End If

    Rem Response.Write "<FONT SIZE=-1>"

    Select Case Session(CurrentField & "EditOption")
' (Campos tipos Texto, UF, Area de texto e Data de Atualização)

      Case efFldTypeText, efFldTypeUF, efFldTypeTextArea, efFldTypeAtualization, efFldTypeDateOfUpdate, efFldTypeDateOfDelete, efFldTypeDateOfInsert, efFldTypeHTML, efFldTypeTextAreaCounter
        Response.Write FormatShowFieldCaption(EditCaption(False))
        Response.Write " : "

        If (Session(CurrentField & "ExternalLink") > "") Then
          BeginMakeLink True

        End If

        If Session(CurrentField & "EditOption") = efFldTypeTextArea Or _
           Session(CurrentField & "EditOption") = efFldTypeTextAreaCounter Then Value = Replace(Value, vbCrLf, "<BR>")

        Response.Write FormatShowField(sobjString.Mask(FormatVariable(Value, Session(CurrentField & "Type")), Session(CurrentField & "Mask"), Session(CurrentField & "MaskCompleter")))
        EndMakeLink

' (Campo tipo Cor)
      Case efFldTypeColor
        Response.Write EditCaption(False)
        Response.Write " : <SPAN style='background:" & Value & ";width:31;height:25'></SPAN>"

' (Campo tipo Endereço HTTP)
      Case efFldTypeHTTP
        Response.Write FormatShowFieldCaption(EditCaption(False) & " : ")
        'Alterado/Criado (Ridai)
        Response.Write "<A HRef=""http://" & Value & """"

        If LinkTarget > "" Then Response.Write " TARGET=""" & LinkTarget & """"

        Response.Write ">" & FormatShowField(Value) & "</A>"
        '----------------------

      'Criado (Ridai)
      ' (Campo tipo Endereço FTP)
      Case efFldTypeFTP
        Response.Write FormatShowFieldCaption(EditCaption(False) & " : ")
        Response.Write "<A HRef=""ftp://" & Value & """"
        If LinkTarget > "" Then Response.Write " TARGET=""" & LinkTarget & """"
        Response.Write ">" & FormatShowField(Value) & "</A>"

      ' (Campo tipo Endereço Gopher)
      Case efFldTypeGopher
        Response.Write FormatShowFieldCaption(EditCaption(False) & " : ")
        Response.Write "<A HRef=""gopher://" & Value & """"
        If LinkTarget > "" Then Response.Write " TARGET=""" & LinkTarget & """"
        Response.Write ">" & FormatShowField(Value) & "</A>"
      '------------------------

      ' (Campo tipo E-Mail)
      Case efFldTypeEMail
        Response.Write FormatShowFieldCaption(EditCaption(False) & " : ")
        Response.Write "<A HRef=mailto:" & Value & ">" & FormatShowField(Value) & "</A>"

' (Campo tipo Arquivo de Imagem)
      Case efFldTypeImage
        Response.Write FormatShowFieldCaption(EditCaption(False))
        sobjShow.BR
        If Value > "" Then
          sobjShow.Image Session(CurrentField & "ImageRoot") & "/" & Value, "", Session(CurrentField & "ImageRoot") & "/" & Value, Session(CurrentField & "ImageRoot") & "/" & Value, "", Session(CurrentField & "ImageWidth"), "", , "__blank"

        End If

' (Campo tipo Arquivo)
      Case efFldTypeFile
        Response.Write FormatShowFieldCaption(EditCaption(False) & " : ")
        If Value > "" Then
          sobjShow.HTMLCR "<A HREF=""/" & Session(CurrentField & "FileRoot") & "/" & Value & """ TARGET=""__blank"">" & Value & FileSize(Initializer.SiteRootDir & Session(CurrentField & "FileRoot") & "\" & Value) & "</A>"

        End If

' (Campo tipo Lookup)
      Case efFldTypeLookup
        Response.Write FormatShowFieldCaption(EditCaption(False))
        Response.Write " : "

        If Value > "" Then
          sql = "SELECT * FROM " & Session(CurrentField & "LookupTable")
          sql = sql & " WHERE " & Session(CurrentField & "LookupField") & "="
          sql = sql & FormatForSQL(Value, Session(CurrentField & "Type"))

          Rem sobjCMD.CommandText = sql

          'Alterado/Criado (Ridai)
          Set CurrConnection = sobjConn

          If ShowSQLQuery Then ShowCurrentSQL sql, "ShowFieldAsText", adOpenForwardOnly, adLockReadOnly

          On Error Resume Next

          sobjRS.Open sql, sobjConn, adOpenForwardOnly, adLockReadOnly

          If Err.number <> 0 Then
            If DebugMode Then
              ShowSQLErrors sql, "ShowFieldAsText", adOpenForwardOnly, adLockReadOnly
            Else
              ShowDefaultError
            End If
            Err.Clear
            Exit Sub

          End If
          On Error GoTo 0
          '----

          If Not sobjRS.EOF And Not sobjRS.BOF Then
            BeginMakeLink True
            Response.Write FormatShowField(sobjRS(Session(CurrentField & "LookupFieldText")))
            EndMakeLink

          End If
          sobjRS.Close

        End If

' (Campo tipo Checkbox)
      Case efFldTypeCheck
        Response.Write FormatShowFieldCaption(EditCaption(False) & " ")
        BeginMakeLink True
        If sobjString.ToString(Value) = sobjString.ToString(Session(CurrentField & ",1Value")) Then
          Response.Write FormatShowField(Session(CurrentField & ",1Caption"))

        Else
          Response.Write FormatShowField(Session(CurrentField & ",2Caption"))

        End If
        EndMakeLink

' (Campos tipos Select e Radio)
      Case efFldTypeRadio, efFldTypeSelect
        Response.Write FormatShowFieldCaption(EditCaption(False) & " ")

        BeginMakeLink True

        For i = 1 To Session(CurrentField & "ValueNumber")
          If sobjString.ToString(Value) = sobjString.ToString(Session(CurrentField & "," & i & "Value")) Then
            Response.Write FormatShowField(Session(CurrentField & "," & i & "Caption"))

          End If

        Next
        EndMakeLink

    End Select
Rem    Response.Write "</FONT>"

  End If

End Sub
Private Function FormatShowField(ByVal strField)
  Dim strSpan As String
  Dim strFont As String

  FormatShowField = ""

  strSpan = ""

  If Session(CurrentField & "StyleSheet") > "" Then strSpan = strSpan & " STYLE=""" & Session(CurrentField & "StyleSheet") & """"
  If Session(CurrentField & "Class") > "" Then strSpan = strSpan & " CLASS=""" & Session(CurrentField & "Class") & """"
  If Session(CurrentField & "Align") > "" Then strSpan = strSpan & " ALIGN=""" & Session(CurrentField & "Align") & """"

  strFont = ""

  If Session(CurrentField & "FontFace") > "" Then strFont = strFont & " FACE=""" & Session(CurrentField & "FontFace") & """"
  If Session(CurrentField & "FontSize") > "" Then strFont = strFont & " SIZE=""" & Session(CurrentField & "FontSize") & """"
  If Session(CurrentField & "FontColor") > "" Then strFont = strFont & " COLOR=""" & Session(CurrentField & "FontColor") & """"
  If Session(CurrentField & "FontBGColor") > "" Then strFont = strFont & " BGCOLOR=""" & Session(CurrentField & "FontFace") & """"

  If strSpan > "" Then FormatShowField = FormatShowField & "<SPAN" & strSpan & ">"
  If strFont > "" Then FormatShowField = FormatShowField & "<FONT" & strFont & ">"

  If Session(CurrentField & "FontBold") Then FormatShowField = FormatShowField & "<B>"
  If Session(CurrentField & "FontItalic") Then FormatShowField = FormatShowField & "<I>"
  If Session(CurrentField & "FontUnderline") Then FormatShowField = FormatShowField & "<U>"

  FormatShowField = FormatShowField & strField

  If Session(CurrentField & "FontUnderline") Then FormatShowField = FormatShowField & "</U>"
  If Session(CurrentField & "Italic") Then FormatShowField = FormatShowField & "</I>"
  If Session(CurrentField & "FontBold") Then FormatShowField = FormatShowField & "</B>"

  If strFont > "" Then FormatShowField = FormatShowField & "</FONT>"
  If strSpan > "" Then FormatShowField = FormatShowField & "</SPAN>"

End Function

Private Function FormatShowFieldCaption(ByVal strCaption)
  Dim strSpan As String
  Dim strFont As String

  FormatShowFieldCaption = ""

  strSpan = ""

  If Session(CurrentField & "CaptionStyleSheet") > "" Then strSpan = strSpan & " STYLE=""" & Session(CurrentField & "CaptionStyleSheet") & """"
  If Session(CurrentField & "CaptionClass") > "" Then strSpan = strSpan & " CLASS=""" & Session(CurrentField & "CaptionClass") & """"
  If Session(CurrentField & "CaptionAlign") > "" Then strSpan = strSpan & " ALIGN=""" & Session(CurrentField & "CaptionAlign") & """"

  strFont = ""

  If Session(CurrentField & "CaptionFontFace") > "" Then strFont = strFont & " FACE=""" & Session(CurrentField & "CaptionFontFace") & """"
  If Session(CurrentField & "CaptionFontSize") > "" Then strFont = strFont & " SIZE=""" & Session(CurrentField & "CaptionFontSize") & """"
  If Session(CurrentField & "CaptionFontColor") > "" Then strFont = strFont & " COLOR=""" & Session(CurrentField & "CaptionFontColor") & """"
  If Session(CurrentField & "CaptionFontBGColor") > "" Then strFont = strFont & " BGCOLOR=""" & Session(CurrentField & "CaptionFontFace") & """"

  If strSpan > "" Then FormatShowFieldCaption = FormatShowFieldCaption & "<SPAN" & strSpan & ">"
  If strFont > "" Then FormatShowFieldCaption = FormatShowFieldCaption & "<FONT" & strFont & ">"

  If Session(CurrentField & "CaptionFontBold") Then FormatShowFieldCaption = FormatShowFieldCaption & "<B>"
  If Session(CurrentField & "CaptionFontItalic") Then FormatShowFieldCaption = FormatShowFieldCaption & "<I>"
  If Session(CurrentField & "CaptionFontUnderline") Then FormatShowFieldCaption = FormatShowFieldCaption & "<U>"

  FormatShowFieldCaption = FormatShowFieldCaption & strCaption

  If Session(CurrentField & "CaptionFontUnderline") Then FormatShowFieldCaption = FormatShowFieldCaption & "</U>"
  If Session(CurrentField & "CaptionFontItalic") Then FormatShowFieldCaption = FormatShowFieldCaption & "</I>"
  If Session(CurrentField & "CaptionFontBold") Then FormatShowFieldCaption = FormatShowFieldCaption & "</B>"

  If strFont > "" Then FormatShowFieldCaption = FormatShowFieldCaption & "</FONT>"
  If strSpan > "" Then FormatShowFieldCaption = FormatShowFieldCaption & "</SPAN>"

End Function

Private Function EditFieldValue(ByVal objRS2 As ADODB.RecordSet) As Variant
' (Retorna o valor do campo atual formatado para listagem, parametro: RecordSet)
Dim sql, Value, i

  If sobjString.ToString(objRS2(Session(CurrentField & "Field")).Value) > "" Then
    Select Case Session(CurrentField & "EditOption")
' (Campos tipos Texto, UF, Area de texto e Data de Atualização)

      Case efFldTypeText, efFldTypeUF, efFldTypeTextArea, efFldTypeAtualization, efFldTypeDateOfUpdate, efFldTypeDateOfDelete, efFldTypeDateOfInsert, efFldTypeHTML, efFldTypeTextAreaCounter
        If Session(CurrentField & "Type") = efDataTypeText Then
          Value = sobjString.ToString(objRS2(Session(CurrentField & "Field")).Value)

        Else
          Value = objRS2(Session(CurrentField & "Field")).Value

        End If

' (Campo tipo Cor)
      Case efFldTypeColor
        Value = "<SPAN STYLE=""background:" & objRS2(Session(CurrentField & "Field")).Value & ";width:31;height:25""></SPAN>"

' (Campo tipo Arquivo de Imagem)
      Case efFldTypeImage
        Value = "<CENTER><IMG SRC=""" & Session(CurrentField & "ImageRoot") & "/" & objRS2(Session(CurrentField & "Field")).Value & """ WIDTH=""" & Session(CurrentField & "ImageWidth") & """></CENTER>"

' (Campo tipo Arquivo)
      Case efFldTypeFile
        Value = "<A HREF=""" & Session(CurrentField & "FileRoot") & "/" & objRS2(Session(CurrentField & "Field")).Value & """ TARGET=""__blank"">" & objRS2(Session(CurrentField & "Field")).Value & FileSize(Initializer.SiteRootDir & Session(CurrentField & "FileRoot") & "\" & objRS2(Session(CurrentField & "Field")).Value) & "</A>"

' (Campo tipo Endereço HTTP)
      Case efFldTypeHTTP
        Value = "<A HREF=""http://" & sobjString.ToString(objRS2(Session(CurrentField & "Field"))) & """>" & sobjString.ToString(objRS2(Session(CurrentField & "Field"))) & "</A>"

' (Campo tipo E-Mail)
      Case efFldTypeEMail
        Value = "<A HREF=""mailto:" & sobjString.ToString(objRS2(Session(CurrentField & "Field"))) & """>" & sobjString.ToString(objRS2(Session(CurrentField & "Field"))) & "</A>"

' (Campo tipo Lookup)
      Case efFldTypeLookup
        Value = sobjString.ToString(objRS2(Session(CurrentField & "LookupFieldText")))

' (Campo tipo Checkbox)
      Case efFldTypeCheck
        If sobjString.ToString(objRS2(Session(CurrentField & "Field"))) > "" Then
          If sobjString.ToString(objRS2(Session(CurrentField & "Field"))) = sobjString.ToString(Session(CurrentField & ",1Value")) Then
            Value = Session(CurrentField & ",1Caption")

          Else
            Value = Session(CurrentField & ",2Caption")

          End If
        Else
          Value = Session(CurrentField & ",2Caption")

        End If

' (Campos tipos Select e Radio)
      Case efFldTypeRadio, efFldTypeSelect
        For i = 1 To Session(CurrentField & "ValueNumber")
          If sobjString.ToString(objRS2(Session(CurrentField & "Field"))) = Session(CurrentField & "," & i & "Value") Then
            Value = Session(CurrentField & "," & i & "Caption")

          End If

        Next

    End Select

  Else
    Value = ""

  End If
  EditFieldValue = Value

End Function

' (Funções que tratam as strings com coleções de campos, como FormUnit e FomList)
Private Sub BeginReadOperation(ByVal strData As String)
' (Inicializa a variável com a string a ser tratada)
  sstrText = strData

End Sub

Private Function IsThereOperation() As Boolean
 ' (Se a variável com a string não está vazia)
  IsThereOperation = (sstrText > "")

End Function

Public Function PrepareForm(ByVal strText As String) As String
' (Codifica um HTML a ser inserido numa string de coleção, para não haver confusão entre os
' operadores do HTML com os do EditForm)
Dim strAux As String
  strAux = strText
  strAux = Replace(strAux, ".", "@~@")
  strAux = Replace(strAux, ";", "@!@")
  strAux = Replace(strAux, "=", "@?@")
  strAux = Replace(strAux, "-", "@@@")
  strAux = Replace(strAux, ",", "@%@")
  PrepareForm = strAux

End Function

Private Function NextOperation() As String
' (Retorna o próximo valor ou operação da string)
Dim strSubText As String, strChar As String
Dim blnEnd As Integer

  strChar = Left(sstrText, 1)
  blnEnd = False

  Do
    Select Case strChar
      Case "."
        If strSubText = "" Then
          CurrentTable = LastOperation

        End If
        sstrText = Mid(sstrText, 2)
        strChar = Left(sstrText, 1)

      Case ",", ";", "-", "=", "(", ")", "|", "#", ""
        If strSubText = "" Then
          strSubText = strChar
          sstrText = Mid(sstrText, 2)

        End If
        blnEnd = True

      Case Else
        strSubText = strSubText & strChar
        sstrText = Mid(sstrText, 2)
        strChar = Left(sstrText, 1)

    End Select

   Loop While Not blnEnd

' (Decodifica um HTML inserido na string de coleção)
   strSubText = Replace(strSubText, "@~@", ".")
   strSubText = Replace(strSubText, "@!@", ";")
   strSubText = Replace(strSubText, "@?@", "=")
   strSubText = Replace(strSubText, "@@@", "-")
   strSubText = Replace(strSubText, "@%@", ",")

   NextOperation = strSubText
   sstrSubText = strSubText

End Function

Private Function LastOperation() As String
' (Retorna o último valor encontrado pelo NextOperation)
  LastOperation = sstrSubText

End Function

Private Sub BeginReadNumbers(ByVal strData As String)
' (Inicializa a variável com a string de nº a ser tratada)
  sstrText2 = strData

End Sub

Private Function IsThereNumber() As Boolean
' (Se a variável com a string de nº não está vazia)
  IsThereNumber = (sstrText2 > "")

End Function

Private Function NextNumber() As Integer
' (Retorna o próximo nº da string)
Dim strSubText, strChar As String
Dim blnEnd As Boolean
Dim i As Integer

  i = 0
  strChar = Left(sstrText2, 1)
  blnEnd = False

  Do
    Select Case strChar
      Case ",", ""
        sstrText2 = Mid(sstrText2, 2)
        If strSubText <> "" Then
          blnEnd = True

        End If

      Case Else
        strSubText = strSubText & strChar
        sstrText2 = Mid(sstrText2, 2)
        strChar = Left(sstrText2, 1)

    End Select
    i = i + 1

  Loop While Not blnEnd And i < 20
  If strSubText > "" Then
    NextNumber = CInt(strSubText)
    sintLastNumber = CInt(strSubText)

  Else
    NextNumber = 1
    sintLastNumber = 1

  End If

End Function

Private Sub TotalOperation(ByVal strText As String, ByVal strTotal As String, Optional ByVal blnKeepParts As Boolean = False, Optional ByVal strParts As String = "")
' (Conta os campos ou colunas de tabela indicados numa string, parametros: String, total de campos, deve guardar a
' lista de nº de colunas por linha ?, Onde deve guardá-los ?)
Dim intGreat As Integer, i As Integer

  BeginReadOperation strText
  Session(CurrentTable & strParts) = ""

  i = 0
  intGreat = 0

  sstrSubText = ""

  Do While IsThereOperation
    Select Case NextOperation

      Case ","
        If intGreat < i Then
          intGreat = i

        End If

      Case "="
        If intGreat < i Then
          intGreat = i

        End If
        NextOperation

      Case ";", "-", "|"
        If intGreat < i Then
          intGreat = i

        End If

' (Guardar a lista de nº de colunas por linha)
        If blnKeepParts Then
          If Session(CurrentTable & strParts) > "" Then
            Session(CurrentTable & strParts) = Session(CurrentTable & _
             strParts) + "," + sobjString.ToString(i)

          Else
            Session(CurrentTable & strParts) = sobjString.ToString(i)

          End If
        End If
        i = 0

      Case Else
        i = i + 1

    End Select

    If LastOperation = "=" Then
      NextOperation

    End If
   Loop

   If intGreat < i Then
     intGreat = i

   End If

' (Guardar a lista de nº de colunas por linha)
   If blnKeepParts Then
     If Session(CurrentTable & strParts) > "" Then
       Session(CurrentTable & strParts) = Session(CurrentTable & _
        strParts) + "," + sobjString.ToString(i)

     Else
       Session(CurrentTable & strParts) = sobjString.ToString(i)

     End If
   End If

' (Total de campos ou colunas)
   Session(CurrentTable & strTotal) = intGreat

End Sub

Private Sub MakeServerValidation()
' (Se houver algum campo com erro de prrenchimento então volta para a edição)
  If Not ServerValidates Then
    Response.Redirect Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionScreen & _
    WhatWhereHow & GetTab

  End If

End Sub

Private Function ServerValidates() As Boolean
' (Chama a validação de cada campo, retorna a existência de erro e monta a mensagem com os
' campos errados)
Dim i As Integer, blnIsValid As Boolean

' (Testa se algum campo requerido está vazio)
  blnIsValid = True
  BeginReadOperation Session("LastTab")
  Do While IsThereOperation
    Select Case NextOperation
      Case ",", ";", "-", ""
      Case "="
        IsThereOperation

      Case Else
        If FindField(CurrentTable, LastOperation) Then
          'Criado/Alterado (Ridai Govinda)
          If Not IdentityCheck(False) Then

            If (Session(CurrentField & "RequireMessage") > "") Then
              If Not Session(CurrentField & "Value") > "" Then
                Message = Message & ErrorImage & Session(CurrentField & "RequireMessage") & "<BR>"
                Session(CurrentField & "ValidationError") = Session(CurrentField & "RequireMessage")
                blnIsValid = False

              End If

            ElseIf Not Session(CurrentField & "Null") Then
              If Not Session(CurrentField & "Value") > "" Then
                Message = Message & ErrorImage & sobjString.FormatText(LanguageLocalizer(52), "<B>" & EditCaption(False) & "</B>") & "<BR>"
                Session(CurrentField & "ValidationError") = sobjString.FormatText(LanguageLocalizer(52), EditCaption(False))
                blnIsValid = False

              End If
            End If

          End If
          '-------------------------------

        End If
    End Select
  Loop

' (Testa se algum campo foi preenchido de forma errada)
  BeginReadOperation Session("LastTab")
  Do While IsThereOperation
    Select Case NextOperation
      Case ",", ";", "-", ""
      Case "="
        IsThereOperation

      Case Else
        If FindField(CurrentTable, LastOperation) Then
          If Not ValidField(CurrentField) Then
            blnIsValid = False

          End If
        End If

    End Select
  Loop

  ServerValidates = blnIsValid

End Function

Private Function ErrorImage() As String
' (Retorna o HTML que monta a imagem que indica erro de preenchimento)
  ErrorImage = "<IMG SRC=""" & Session("htDefaultHiperToolsWorkPath") & "/CE.gif"" width=5 height=5> "

End Function

Private Function ValidField(ByVal strField As String) As Boolean
' (Faz a validação de um campo observando sua opção de validação)
  Dim strSQL As String
  Dim blnIsValid As Boolean
  blnIsValid = True

  Select Case Session(strField & "ValidationOption")
    Case efValOptCGC
      If Not Validate.IsCGC(Session(strField & "Value")) Then
        blnIsValid = False

      End If

    'Criado (Ridai Govinda)
    Case efValOptCGCCPF
      blnIsValid = Validate.IsCPF_CGC(Session(strField & "Value"), vlCp_CGCFirst)

    Case efValOptCPFCGC
      blnIsValid = Validate.IsCPF_CGC(Session(strField & "Value"), vlCp_CPFFirst)

    Case efValOptCEP
      blnIsValid = Validate.IsCEP(Session(strField & "Value"))
    '---------------------

    Case efValOptCPF
      If Not Validate.IsCPF(Session(strField & "Value")) Then
        blnIsValid = False

      End If

    Case efValOptDate
      If TypeName(Session(strField & "Value")) = "string" Then
        If Not Validate.IsValidDate(Session(strField & "Value")) Then
          blnIsValid = False

        End If
      End If

    Case efValOptCompareDates
      If FindField(CurrentTable, Session(strField & "DoubleValidateField")) Then
        Select Case Validate.CompareDate(Session(strField & "Value"), Session(CurrentField & "Value"))
          Case 1
            blnIsValid = False
            If FindField(CurrentTable, Session(strField & "Field")) Then
              Message = Message & ErrorImage & sobjString.FormatText(LanguageLocalizer(53), "<B>" & EditCaption(False) & "</B>") & "<BR>"
              Session(CurrentField & "ValidationError") = sobjString.FormatText(LanguageLocalizer(53), EditCaption(False))

            End If

          Case 2
            blnIsValid = False
            Message = Message & ErrorImage & sobjString.FormatText(LanguageLocalizer(53), "<B>" & EditCaption(False) & "</B>") & "<BR>"
            Session(CurrentField & "ValidationError") = sobjString.FormatText(LanguageLocalizer(53), EditCaption(False))

          Case 3
            blnIsValid = False
            If FindField(CurrentTable, Session(strField & "Field")) Then
              Message = Message & ErrorImage & sobjString.FormatText(LanguageLocalizer(53), "<B>" & EditCaption(False) & "</B>") & "<BR>"
              Session(CurrentField & "ValidationError") = sobjString.FormatText(LanguageLocalizer(53), EditCaption(False))
              If FindField(CurrentTable, Session(CurrentField & "DoubleValidateField")) Then
                Message = Message & ErrorImage & sobjString.FormatText(LanguageLocalizer(53), "<B>" & EditCaption(False) & "</B>") & "<BR>"
                Session(CurrentField & "ValidationError") = sobjString.FormatText(LanguageLocalizer(53), EditCaption(False))

              End If

            End If

          Case 4
            blnIsValid = False
            Message = Message & ErrorImage & Session(strField & "DoubleValidateMessage")
            Session(CurrentField & "ValidationError") = Session(strField & "DoubleValidateMessage")

        End Select
      End If

    Case efValOptTime
      If Not Validate.IsValidTime(Session(strField & "Value")) Then
        blnIsValid = False

      Else
        Session(strField & "Value") = Validate.FormatTime(Session(strField & "Value"))

      End If

    Case efValOptSepDate
Rem      FormJavaValidSepDate Session(CurrentTable & i & "Field") & "DayValue", Session(CurrentTable & i & "Field") & "MonthValue", Session(CurrentTable & i & "Field") & "YearValue", "V" & Session(CurrentTable & i & "Field"), False, ""

    Case efValOptDateMToday
      If Validate.IsValidDate(Session(strField & "Value")) Then
        If CDate(Session(strField & "Value")) > Now + 1 Then
          blnIsValid = False

        End If

      Else
        blnIsValid = False

      End If

    Case efValOptSepDateMToday
Rem      FormJavaValidSepDate Session(CurrentTable & i & "Field") & "DayValue", Session(CurrentTable & i & "Field") & "MonthValue", Session(CurrentTable & i & "Field") & "YearValue", "V" & Session(CurrentTable & i & "Field"), True, Session(CurrentTable & i & "Caption") & " deve ser maior que hoje"

    Case efValOptEmail
      If Not Validate.IsEmail(Session(strField & "Value")) Then
        blnIsValid = False

      End If

  End Select

  If Not blnIsValid And Session(strField & "ValidationOption") <> efValOptCompareDates Then
    If FindField(CurrentTable, strField) Then
      Message = Message & ErrorImage & sobjString.FormatText(LanguageLocalizer(53), "<B>" & EditCaption(False) & "</B>") & "<BR>"
      Session(CurrentField & "ValidationError") = sobjString.FormatText(LanguageLocalizer(53), EditCaption(False))

    End If

  End If

' (Faz a validação de unicidade de um campo)
  If blnIsValid And Session(CurrentField & "IsUnique") > "" Then
    strSQL = "Select Count(*) As RecNumber From " & CurrentTable & " Where " & Session(CurrentField & "Field") & "=" & FormatForSQL(Session(CurrentField & "Value"), Session(CurrentField & "Type"))
    Rem sobjCMD2.CommandType = adCmdText

    'Alterado/Criado (Ridai)
    Set CurrConnection = sobjConn2

    If ShowSQLQuery Then ShowCurrentSQL strSQL, "ValidField", adOpenKeyset, adLockReadOnly

    On Error Resume Next

    sobjRS2.Open strSQL, sobjConn2, adOpenKeyset, adLockReadOnly

    If Err.number <> 0 Then
      If DebugMode Then
        ShowSQLErrors strSQL, "ValidField", adOpenKeyset, adLockReadOnly
      Else
        ShowDefaultError
      End If
      Err.Clear
      Exit Function

    End If
    On Error GoTo 0
    '----

    If sobjRS2("RecNumber") > 0 Then
      blnIsValid = False
      Message = Message & ErrorImage & sobjString.FormatText(LanguageLocalizer(54), "<B>" & EditCaption(False) & "</B>") & "<BR>"
      Session(CurrentField & "ValidationError") = sobjString.FormatText(LanguageLocalizer(54), EditCaption(False))

    End If
  End If

  ValidField = blnIsValid

End Function

Private Sub ClientValidates()
' (Montas as validações em JavaScript no Browser)

Dim i As Integer
  sstrLastValidate = ""

' (Validação de erro de preenchimento)
  BeginReadOperation Session("LastTab")
  Do While IsThereOperation
    Select Case NextOperation
      Case ",", ";", "-", ""
      Case "="
        IsThereOperation

      Case Else
        If FindField(CurrentTable, LastOperation) Then
          Select Case Session(CurrentField & "ValidationOption")
             'Criado (Ridai Govinda)
           Case efValOptCGCCPF
              Validate.FormJavaValidCPFCGC Session(CurrentField & "Field"), "frmEditForm", "V" & Session(CurrentField & "Field"), vlCp_CGCFirst

           Case efValOptCPFCGC
              Validate.FormJavaValidCPFCGC Session(CurrentField & "Field"), "frmEditForm", "V" & Session(CurrentField & "Field"), vlCp_CPFFirst

           Case efValOptCGC
              Validate.FormJavaValidCGC Session(CurrentField & "Field"), "frmEditForm", "V" & Session(CurrentField & "Field")

            Case efValOptCPF
              Validate.FormJavaValidCPF Session(CurrentField & "Field"), "frmEditForm", "V" & Session(CurrentField & "Field")

            Case efValOptCEP
              Validate.FormJavaValidCEP Session(CurrentField & "Field"), "frmEditForm", "V" & Session(CurrentField & "Field")

            Case efValOptDate
              Validate.FormJavaValidDate Session(CurrentField & "Field"), "frmEditForm", "V" & Session(CurrentField & "Field"), False, ""

            Case efValOptCompareDates
              Validate.FormJavaCompareDate Session(CurrentField & "Field"), Session(CurrentField & "DoubleValidateField"), "frmEditForm", "V" & Session(CurrentField & "Field"), Session(CurrentField & "DoubleValidateMessage")

            Case efValOptTime
Rem              FormJavaValidTime Session(CurrentField & "Field"), "V" & Session(CurrentField & "Field"), False, ""

            Case efValOptSepDate
              Validate.FormJavaValidSepDate Session(CurrentField & "Field") & "DayValue", Session(CurrentField & "Field") & "MonthValue", Session(CurrentField & "Field") & "YearValue", "frmEditForm", "V" & Session(CurrentField & "Field"), False, ""

            Case efValOptDateMToday
              Validate.FormJavaValidDate Session(CurrentField & "Field"), "frmEditForm", "V" & Session(CurrentField & "Field"), True, EditCaption(False) & " deve ser maior que hoje"

            Case efValOptSepDateMToday
              Validate.FormJavaValidSepDate Session(CurrentField & "Field") & "DayValue", Session(CurrentField & "Field") & "MonthValue", Session(CurrentField & "Field") & "YearValue", "frmEditForm", "V" & Session(CurrentField & "Field"), True, EditCaption(False) & " deve ser maior que hoje"

            Case efValOptEmail
              Validate.FormJavaValidEMail Session(CurrentField & "Field"), "frmEditForm", "V" & Session(CurrentField & "Field")

          End Select
          If Session(CurrentField & "ValidationOption") <> efValOptNone Then
            If sstrLastValidate = "" Then
              sstrLastValidate = sstrLastValidate & "V" & Session(CurrentField & "Field") & "()"

            Else
              sstrLastValidate = sstrLastValidate & " && V" & Session(CurrentField & "Field") & "()"

            End If
          End If
        End If

    End Select

  Loop

' (Cria a função que chama todas as validações)
  Validate.JavaEmpty
  Validate.JavaBegin "ValidateForm", "frmEditForm"

' (Validação de campo requerido vazio)
  BeginReadOperation Session("LastTab")
  Do While IsThereOperation
    Select Case NextOperation
      Case ",", ";", "-", ""
      Case "="
        IsThereOperation

      Case Else
        If FindField(CurrentTable, LastOperation) Then
          If (Session(CurrentField & "EditOption") <> efFldTypePassword Or _
              Not sobjString.ToString(Session(CurrentField & "Value")) > "") And _
              Session(CurrentField & "EditOption") <> efFldTypeCheck And _
              Session(CurrentField & "EditOption") <> efFldTypeImage And _
              Session(CurrentField & "EditOption") <> efFldTypeFile Then

            If Session(CurrentField & "RequireMessage") > "" Then
              Validate.JavaRequired Session(CurrentField & "Field"), Session(CurrentField & "RequireMessage")

            ElseIf Not Session(CurrentField & "Null") And Session(CurrentField & "EditOption") <> efFldTypeRadio Then
              Validate.JavaRequired Session(CurrentField & "Field"), "O campo " & EditCaption(False) & " deve ser preenchido"

            End If
          End If
        End If

    End Select

  Loop
  If sstrLastValidate > "" Then
    sobjShow.HTMLCR "return (" & sstrLastValidate & ");"

  Else
    sobjShow.HTMLCR "return true;"

  End If
  Response.Write "}"

End Sub

Private Sub BeginMakeLink(ByVal blnPrefforExtenal As Boolean)
' (Cria a ancora dos campos que possuem ExternalLink ou InternalLink, parametro: O link
' ExternalLink tem preferência de aparição sobre o InternalLink ?)

  If blnPrefforExtenal Then
    If Session(CurrentField & "ExternalLink") > "" Then
      'Alterado/Criado (Ridai Govinda)
      MakeExternalLink
      '------------------------------

    Else
      If Session(CurrentField & "InternalLink") > "" Then
        MakeInternalLink

      End If
    End If

  Else
    If Session(CurrentField & "InternalLink") > "" Then
      MakeInternalLink

    Else
      If Session(CurrentField & "ExternalLink") > "" Then
      'Alterado/Criado (Ridai Govinda)
      MakeExternalLink
      '------------------------------

      End If
    End If

  End If

End Sub
'Criado (Ridai Govinda)
Private Sub MakeExternalLink()
  Dim strLink As String, strFieldFound As String
  strLink = "<A HREF=" & Session(CurrentField & "ExternalLink")
  strFieldFound = CurrentField

  If Session(CurrentField & "ExtLinkQSParam") > "" Then
    If FindField(CurrentTable, Session(CurrentField & "ExtLinkField")) Then
      If InStr(strLink, "?") > 0 Then
        strLink = strLink & "&"
      Else
        strLink = strLink & "?"
      End If

      strLink = strLink & Session(strFieldFound & "ExtLinkQSParam") & "=" & _
                sobjRS2(Session(CurrentField & "Field"))
    End If

    CurrentField = sobjString.TotalNormalizeString(strFieldFound)

    Rem sobjRS2 (Session(CurrentField & "Field"))
    Rem Server.URLEncode (FieldValue(CurrentTable, Session(CurrentField & "ExtLinkField")))
  End If

  Response.Write strLink & ">"
End Sub
'---------------------

Private Sub MakeInternalLink()
' (Monta o InternalLink)
Dim strURL As String

  strURL = "<A HREF=" & Initializer.ScriptName & "?" & efQueryStrAction & "=" & efQSActionEditor & "&" & _
           efQueryStrWhat & "=" & Session(CurrentField & "InternalLink")

  Rem --- Alterado --- (Ueslei/Ridai)
  If sparEditFilter > "" And Session(CurrentField & "InternalLink") = parWhat Then strURL = strURL & GetFilter
  Rem ---

  If Session(CurrentTable & "WhatList") = Session(CurrentField & "InternalLink") Or Session(CurrentField & "ILEditable") Then
    strURL = strURL & efQueryStrEditableStr

  End If

  strURL = strURL & TargetOfLinks & ">"

  Response.Write strURL

End Sub

Private Function TargetOfLinks() As String
' (Cria o find do InternalLink dependendo dos ILTargets[Internal Link Targets])
Dim strAux As String
Dim i, intCurrentField As Integer

  For i = 1 To Session(CurrentField & "ILTargetNumber")
    If strAux > "" Then
      strAux = strAux & " And "

    End If
    If sobjRS2.State = 0 Then
      strAux = strAux & Session(CurrentField & "," & i & "ILTarget") & "=" & FormatForSQL(Session(CurrentField & "Value"), Session(CurrentField & "Type"))

    Else
      intCurrentField = Session("CurrentField")
      If FindField(CurrentTable, Session(CurrentField & "," & i & "ILTarget")) Then
        strAux = strAux & Session(CurrentField & "Field") & "=" & FormatForSQL(sobjRS2(Session(CurrentField & "Field")), Session(CurrentField & "Type"))
        Session("CurrentField") = intCurrentField

      Else
        Session("CurrentField") = intCurrentField
        strAux = strAux & Session(CurrentField & "," & i & "ILTarget") & "=" & FormatForSQL(sobjRS2(Session(CurrentField & "Field")), Session(CurrentField & "Type"))

      End If
      Session("CurrentField") = intCurrentField
    End If

  Next
  If strAux > "" Then
    TargetOfLinks = "&" & efQueryStrFind & "=" & Server.URLEncode(strAux)

  End If

End Function

Private Sub EndMakeLink()
' (Cria o find do registro atual)
  If Session(CurrentField & "ExternalLink") > "" Or Session(CurrentField & "InternalLink") > "" Then
    sobjShow.HTMLCR "</A>"

  End If

End Sub

Private Function MakeTarget() As String
' (Cria o find do registro atual)
Dim strAux As String, i As Integer

  strAux = ""
  For i = 1 To Session(CurrentTable & "FieldNumber")
    If Session(CurrentTable & i & "IsKey") Then
      If strAux > "" Then
        strAux = strAux & " And "

      End If
      If Session(CurrentTable & i & "OldValue") > "" Then
        strAux = strAux & Session(CurrentTable & i & "Field") & "=" & FormatForSQL(Session(CurrentTable & i & "OldValue"), Session(CurrentTable & i & "Type"))

      End If
    End If

  Next

  If strAux > "" Then
    MakeTarget = "&" & efQueryStrFind & "=" & Server.URLEncode(strAux)

  End If

End Function

Private Function GetFind() As String
' (Repete o parametro find se este não for vazio)
  If sparEditFind > "" Then
    GetFind = "&" & efQueryStrFind & "=" & sparEditFind

  End If

End Function

Private Function GetTab() As String
' (Repete o parametro tab se este não for vazio)
  If sparEditTab > "" Then
    GetTab = "&" & efQueryStrTab & "=" & Server.URLEncode(sparEditTab)

  End If

End Function

Private Function GetEditable() As String
' (Repete o parametro modo de edição se este não for vazio)
  If sparEditEditable > "" Then
    GetEditable = "&" & efQueryStrEditable & "=" & Server.URLEncode(sparEditEditable)

  End If

End Function

Private Function GetFilter() As String
' (Repete o parametro filtro se este não for vazio)
  If sparEditFilter > "" Then
    GetFilter = "&" & efQueryStrFilter & "=" & Server.URLEncode(sparEditFilter)

  End If

End Function

Private Function GetDefaults() As String
' (Repete o parametro valores default para inclusão se este não for vazio)
  If sparEditDefaults > "" Then
    GetDefaults = "&" & efQueryStrDefaults & "=" & Server.URLEncode(sparEditDefaults)

  End If

End Function

Private Function GetList() As String
' (Repete o parametro relatório se este não for vazio)
  If sparEditList > "" Then
    GetList = "&" & efQueryStrList & "=" & Server.URLEncode(sparEditList)

  End If

End Function

Private Function MakeFilter(ByVal objRS As ADODB.RecordSet, ByVal strFields As String) As String
'(Cria um filtro a partir dos valores dos campos, parametro: RecordSet, campos)
Dim strAux As String

  BeginReadOperation strFields

  Do While IsThereOperation
    Select Case NextOperation
      Case ",", ";", "-", "="

      Case Else
        If FindField(CurrentTable, LastOperation) Then
          If strAux > "" Then
            strAux = strAux + " AND "

          End If
          strAux = strAux & Session(CurrentField & "Field") & "=" & FormatForSQL(objRS(Session(CurrentField & "Field")), Session(CurrentField & "Type"))

        End If
    End Select

  Loop

  MakeFilter = Server.URLEncode(strAux)

End Function

Private Function WhatList() As String
' (Repete o parametro tabela atual se este não for vazio)
 If sparEditWhat > "" Then
    WhatList = "&" & efQueryStrWhat & "=" & sparEditWhat

  ElseIf Session(CurrentTable & "WhatList") > "" Then
    WhatList = "&" & efQueryStrWhat & "=" & Session(CurrentTable & "WhatList")

  End If

End Function

Private Function WhereHow() As String
  'Alterado/Criado (Ridai Govinda)
  WhereHow = WhatList & GetEditable & GetDefaults
  '---------------------

End Function

Private Function WhatWhereHow() As String
  'Criado (Ridai Govinda)
  WhatWhereHow = MakeTarget & WhatList & GetFilter & GetEditable & GetDefaults
  '---------------------

End Function

'Criado (Ridai Govinda)
Public Property Let FirstPageLink(ByVal strLocal As String)
  FPLPLinkSet strLocal
End Property
'---------------------

Public Sub FPLPLinkSet(ByVal strLocal)
' (Altera o link de Página Inicial)
  Session("FPLPLink") = strLocal

End Sub

'Alterado/Criado (Ridai)
Public Property Let DoFPLPLink(ByVal blnParam As Boolean)
  ShowFirstLastPageLink = blnParam
End Property

Public Property Get DoFPLPLink() As Boolean
  DoFPLPLink = ShowFirstLastPageLink
End Property
'---------

Private Sub FPLPLink()
' (Monta o link de Página Inicial)
  If DoFPLPLink Then
    sobjShow.Paragraph
    sobjShow.Center

    Response.Write "<FONT SIZE=2>"

    If Session("FPLPLink") > "" Then
      Response.Write "<A HREF=" & Session("FPLPLink")

    Else
      Response.Write "<A HREF=" & Initializer.ScriptName

    End If

    Response.Write ">" & LanguageLocalizer(55) & "</a> "

    If QSLast > "" Then
      Response.Write "<A HREF=" & QSLast & ">" & LanguageLocalizer(56) & "</a>"

    End If

    Response.Write "</FONT>"

    sobjShow.CenterEnd
  End If

End Sub
'Criado Alterado (Ridai)
Public Property Let DebugMode(ByVal bolParam As Boolean)
  sblnShowMessage = bolParam

End Property

Public Property Get DebugMode() As Boolean
  DebugMode = sblnShowMessage
End Property

Private Sub ShowSQLErrors(ByVal strSQL As String, ByVal strRoutine As String, _
                          ByVal intOpen As Integer, ByVal intLock As Integer)
  Dim objError As ADODB.Error, strOpen As String, strLock As String

  sobjShow.MessageError sobjString.FormatText(LanguageLocalizer(57), strRoutine)

  For Each objError In Errors
    sobjShow.MessageError sobjString.FormatText(LanguageLocalizer(58), objError.Description)
    sobjShow.MessageError sobjString.FormatText(LanguageLocalizer(59), objError.Source)
    sobjShow.MessageError sobjString.FormatText(LanguageLocalizer(60), objError.number)
    sobjShow.MessageError "+ SQLState: {" & objError.SQLState & "}"
    sobjShow.MessageError "+ NativeError: {" & objError.NativeError & "}"
  Next

  sobjShow.MessageError LanguageLocalizer(61)
  sobjShow.MessageError strSQL

  strOpen = Database.ADOOpenType(intOpen)
  strLock = Database.ADOLock(intLock)

  sobjShow.MessageError LanguageLocalizer(62) & strOpen
  sobjShow.MessageError LanguageLocalizer(63) & strLock

End Sub

Public Property Get Errors() As ADODB.Errors
  Set Errors = CurrConnection.Errors
End Property

'Criado (Ridai)
'Expõe qual é a conexão atual sendo utilizado pelo EditForm:
Public Property Get CurrConnection() As ADODB.Connection
  Set CurrConnection = sobjCurrConn
End Property

Private Property Set CurrConnection(objParam As ADODB.Connection)
  Set sobjCurrConn = objParam
End Property

Private Sub ShowCurrentSQL(ByVal strSQL As String, ByVal strRoutine As String, _
                           ByVal intOpen As Integer, ByVal intLock As Integer)
  Dim strOpen As String, strLock As String

  sobjShow.MessageAlert "HiperTools30.EditForm." & strRoutine & LanguageLocalizer(64)
  sobjShow.MessageAlert strSQL

  strOpen = Database.ADOOpenType(intOpen)
  strLock = Database.ADOLock(intLock)

  sobjShow.MessageAlert LanguageLocalizer(62) & strOpen
  sobjShow.MessageAlert LanguageLocalizer(63) & strLock

End Sub

Public Property Let ShowSQLQuery(ByVal bolParam As Boolean)
  sblnShowSQL = bolParam
End Property

Public Property Get ShowSQLQuery() As Boolean
  ShowSQLQuery = sblnShowSQL
End Property
'---

Private Sub DebugMessage(ByVal strText As String, Optional ByVal intMessageType As Integer = 0)
' (Se estiver em modo de debug então dá uma mensagem, parametros: Mensagem, tipo da
' mensagem[0 - Alert, 1- MessageError])
  If sblnShowMessage Then
    If (intMessageType = 1) Or (Len(strText) > 30 And intMessageType = 0) Then
      sobjShow.MessageError strText

    Else
      sobjShow.JSAlert strText

    End If
  End If

End Sub

Private Function Filter() As String
' (Monta o filtro apartir do QueryString sparEditFilter e dos filtros do AddFilter)
Dim strFilter As String
Dim i As Integer

  Database.AndState = False

  If sparEditFilter > "" Then
    strFilter = Database.PutAnd & sparEditFilter

  End If

  For i = 1 To Session(CurrentTable & "FilterNumber")
    If FindField(CurrentTable, Session(CurrentTable & i & "FilterField")) Then
      '--- Alterado (Uelsei/Ridai)
      Select Case Session(CurrentTable & i & "FilterBooleanOperator")
      Case efBooleanOperatorOR
        strFilter = strFilter & Database.PutOr
      Case Else
        strFilter = strFilter & Database.PutAnd
      End Select
      strFilter = strFilter & OwnerTable(CurrentTable, Session("CurrentField")) & _
                  "." & Session(CurrentTable & i & "FilterField") & _
                  " " & IIf(Session(CurrentTable & i & "FilterOperator") > "", Session(CurrentTable & i & "FilterOperator"), "=") & " " & _
                  FormatForSQL(Session(CurrentTable & i & "FilterValue"), Session(CurrentField & "Type"))
      '---

    End If
  Next

  Filter = strFilter

End Function

Private Sub ClearSession(ByVal strSession As String)
' (Limpa um Session se houver algum dado. O Session não é limpo direto pois se não houver
' valor, este pode não ter sido instanciado, e esvaziá-lo vai instanciá-lo)
  If sobjString.ToString(Session(strSession)) > "" Then
    Session(strSession) = ""

  End If

End Sub

Private Sub FontStyles()
' (Monta os estilos dos links e o JavaScript dos Tips)
Dim strLink, strALink As String

  If Session("BodyLink") > "" Then
    strLink = Session("BodyLink")

  Else
    strLink = "Black"

  End If
  If Session("BodyALink") > "" Then
    strALink = Session("BodyALink")

  Else
    strALink = "Purple"

  End If
  sobjShow.HTMLCR "<STYLE>"
  sobjShow.HTMLCR ".SS{text-decoration:none;color:" & strLink & ";font:bold 8pt}"
  sobjShow.HTMLCR ".SS:HOVER{color:" & strALink & "}"
  sobjShow.HTMLCR ".HL{text-decoration:none;color:"
  If Style.HeaderFont.Color > "" Then
    Response.Write Style.HeaderFont.Color

  Else
    Response.Write strLink

  End If
  sobjShow.HTMLCR ";font:bold 8pt}"
  If Not Style.HeaderFont.Color > "" Then
    sobjShow.HTMLCR ".HL:HOVER{color:" & strALink & "}"

  End If
  sobjShow.HTMLCR ".HT{text-decoration:none;color:"
  If Style.HeaderFont.Color > "" Then
    Response.Write Style.InternalFont.Color

  Else
    Response.Write strLink

  End If
  sobjShow.HTMLCR ";font:bold 8pt}"
  If Not Style.InternalFont.Color > "" Then
    sobjShow.HTMLCR ".HT:HOVER{color:" & strALink & "}"

  End If
  sobjShow.HTMLCR "</STYLE>"

  sobjShow.HTMLCR "<script type=""text/javascript"">"
  sobjShow.HTMLCR "function OpenDesc(text) {"
  sobjShow.HTMLCR "  if (text.style.display =="""") {"
  sobjShow.HTMLCR "    text.style.display = ""none"";"
  sobjShow.HTMLCR "  }"
  sobjShow.HTMLCR "  else {"
  sobjShow.HTMLCR "  text.style.display="""";"
  sobjShow.HTMLCR "  }"
  sobjShow.HTMLCR "}"
  sobjShow.HTMLCR "function makeCool(text) {"
  sobjShow.HTMLCR "  text.style.color = """ & strALink & """"
  sobjShow.HTMLCR "}"
  sobjShow.HTMLCR "function makeNormal(text) {"
  sobjShow.HTMLCR "  text.style.color = """ & strLink & """"
  sobjShow.HTMLCR "}"
  sobjShow.HTMLCR "</script>"

End Sub

Private Sub QSSave()
' (Salva a URL atual)
  Session("QSLast") = Session("QSCurrent")
  Session("QSCurrent") = Initializer.ScriptName & "?" & Initializer.QueryString

End Sub

Private Function QSLast() As String
' (Retorna a URL da tela anterior)
  QSLast = Session("QSLast")

End Function

Private Sub SummaryPars(ByVal strText As String)
' (Processa a string de formatação de um relatório, parametro: string de formatação)
  Dim blnFirst As Boolean, blnFirst2 As Boolean
  Dim blnParameter As Boolean
  Dim i As Integer, j As Integer

  blnParameter = False

  Session("FieldNumber") = 0
  Session("TotalNumber") = 0
  sblnHaveSum = False
  sblnHaveGraph = False

  BeginReadOperation strText
  Do While IsThereOperation
    Select Case NextOperation
      Case ",", ";", "=", "-"

' (Inicio dos parametros de um campo)
      Case "("
        blnParameter = True

' (Fim dos parametros de um campo)
      Case ")"
        blnParameter = False

' (Campo calculado)
      Case "#"
        Session("FieldNumber") = Session("FieldNumber") + 1
        Session(Session("FieldNumber") & "IsCalcField") = "1"
        Session(Session("FieldNumber") & "ShowField") = NextOperation
        Session(Session("FieldNumber") & "Field") = LastOperation
        Session(Session("FieldNumber") & "Value") = 0
        ClearSession Session("FieldNumber") & "Type"
        ClearSession Session("FieldNumber") & "IsGraph"
        ClearSession Session("FieldNumber") & "IsSum"
        ClearSession Session("FieldNumber") & "IsCount"
        ClearSession Session("FieldNumber") & "Operation"
        ClearSession Session("FieldNumber") & "Have1stValue"

      Case Else
' (Parametros)
        If blnParameter Then
          Select Case LCase(LastOperation)
' (Soma)
            Case "s"
              Session(Session("FieldNumber") & "IsSum") = "1"
              Session(Session("FieldNumber") & "OrderSQL") = "SUM(" & Session(Session("FieldNumber") & "Field") & ")"
              sblnHaveSum = True

' (Quantidade)
            Case "c"
              Session(Session("FieldNumber") & "IsCount") = "1"
              Session(Session("FieldNumber") & "OrderSQL") = "COUNT(*)"

' (Gráfico)
            Case "g"
              Session(Session("FieldNumber") & "IsGraph") = "1"
              sblnHaveGraph = True

' (Gráfico de soma)
            Case "gs"
              Session(Session("FieldNumber") & "IsGraph") = "1"
              Session(Session("FieldNumber") & "IsSum") = "1"
              Session(Session("FieldNumber") & "OrderSQL") = "SUM(" & Session(Session("FieldNumber") & "Field") & ")"
              sblnHaveSum = True
              sblnHaveGraph = True

' (Gráfico de Quantidade)
            Case "gc"
              Session(Session("FieldNumber") & "IsGraph") = "1"
              Session(Session("FieldNumber") & "IsCount") = "1"
              Session(Session("FieldNumber") & "OrderSQL") = "COUNT(*)"
              sblnHaveGraph = True

' (Agrupador)
            Case "a"
              sstrGroupField = Session("FieldNumber")

' (Total)
            Case "t"
              Session(Session("FieldNumber") & "IsTotal") = "1"

' (Tipo do campo para formatação)
            Case "f"
              NextOperation
              Session(Session("FieldNumber") & "Type") = NextOperation

' (Operação)
            Case "o"
              NextOperation
              i = InStr(sstrText, ",")
              j = InStr(sstrText, ")")
              If i > j Or i = 0 Then
                i = j

              End If
              If i > 0 Then
                Session(Session("FieldNumber") & "Operation") = Left(sstrText, i - 1)
                sstrText = Right(sstrText, Len(sstrText) - i + 1)

              End If

' (Valor padrão)
            Case "v"
              NextOperation
              i = InStr(sstrText, ",")
              j = InStr(sstrText, ")")
              If i > j Or i = 0 Then
                i = j

              End If
              If i > 0 Then
                Session(Session("FieldNumber") & "Value") = Left(sstrText, i - 1)
                sstrText = Right(sstrText, Len(sstrText) - i + 1)
                Session(Session("FieldNumber") & "Have1stValue") = "1"

              End If

' (Rótulo da coluna)
            Case "l"
              NextOperation
              Session(Session("FieldNumber") & "CaptionField") = NextOperation

          End Select

' (Nome do campo)
        Else
          Session("FieldNumber") = Session("FieldNumber") + 1
          If FindField(CurrentTable, LastOperation) Then
            Session(Session("FieldNumber") & "ShowField") = Session("CurrentField")
            If Session(CurrentField & "EditOption") = efFldTypeLookup Then
              Session(Session("FieldNumber") & "Field") = Session(CurrentField & "LookupFieldText")

            Else
              Session(Session("FieldNumber") & "Field") = Session(CurrentField & "Field")

            End If
            Session(Session("FieldNumber") & "CaptionField") = EditCaption(False)
            Session(Session("FieldNumber") & "OrderSQL") = Session("CurrentField")
            Session(Session("FieldNumber") & "Type") = Session(CurrentField & "Type")
            Session(Session("FieldNumber") & "Value") = 0
            ClearSession Session("FieldNumber") & "IsGraph"
            ClearSession Session("FieldNumber") & "IsSum"
            ClearSession Session("FieldNumber") & "IsCount"
            ClearSession Session("FieldNumber") & "IsCalcField"
            ClearSession Session("FieldNumber") & "Operation"

          Else
            ShowError efErrorCalcFieldNotFound, "EditForm", "ShowSummary", sobjString.FormatText(LanguageLocalizer(9), LastOperation, Session(CurrentList & "ListCaption"))

          End If

        End If

    End Select

  Loop

' (Monta o SELECT[sstrFields] e o GROUP BY[sstrFields2] do relatório)
  sstrFields = ""
  sstrFields2 = ""

  blnFirst = True
  blnFirst2 = True
  For i = 1 To Session("FieldNumber")
    If Not blnFirst And Not Session(i & "Operation") > "" Then
      sstrFields = sstrFields & ", "

    ElseIf Not Session(i & "Operation") > "" Then
      blnFirst = False

    End If

    If Not blnFirst2 And Not Session(i & "IsCalcField") > "" Then
      sstrFields2 = sstrFields2 & ", "

    ElseIf Not Session(i & "IsCalcField") > "" Then
      blnFirst2 = False

    End If

    If Session(i & "IsSum") > "" Then
      sstrFields = sstrFields & "SUM(" & Session(i & "Field") & ") as s" & Session(i & "Field")
      If Not Session(i & "IsCalcField") > "" Then
        sstrFields = sstrFields & ", " & Session(i & "Field")
        sstrFields2 = sstrFields2 & Session(i & "Field")

      End If

    ElseIf Session(i & "IsCount") > "" Then
      sstrFields = sstrFields & "COUNT(*) as s" & Session(i & "Field")
      If Not Session(i & "IsCalcField") > "" Then
        sstrFields = sstrFields & ", " & Session(i & "Field")

      End If

    ElseIf Not Session(i & "IsCalcField") > "" Then
      sstrFields = sstrFields & Session(i & "Field")
      sstrFields2 = sstrFields2 & Session(i & "Field")
      If Session(CurrentTable & Session(i & "ShowField") & "EditOption") = efFldTypeLookup Then
        sstrFields = sstrFields & ", " & Session(CurrentTable & Session(i & "ShowField") & "Field")
        sstrFields2 = sstrFields2 & ", " & Session(CurrentTable & Session(i & "ShowField") & "Field")
        blnFirst = False
        blnFirst2 = True

      End If

    End If

  Next

End Sub

Private Sub ShowSummary()
' (Monta um relatório)
  Dim sql As String, strSelect As String, strFrom As String, strJoin As String, strWhere As String
  Dim strGroup As String, strOrder As String, strSummary As String, strFields As String, strTip2 As String
  Dim intNumCols, i As Integer

' (Localiza o relatório)
  If FindList(sparEditList) Then
    SummaryPars Session(CurrentList & "List")
    intNumCols = Session("FieldNumber")

' (Monta o joins)
    strJoin = ""
    For i = 1 To Session("FieldNumber")
      If Not Session(i & "IsSum") > "" And Not Session(i & "IsCount") > "" And Not Session(i & "IsCalcField") > "" Then
        Session("CurrentField") = Session(i & "ShowField")
        If Session(CurrentField & "EditOption") = efFldTypeLookup Then
           strJoin = strJoin & MakeJoin(CurrentTable, OwnerTable(CurrentTable, Session("CurrentField")), strJoin)
           If InStr(strJoin, " " & Session(CurrentField & "LookupTable") & " ") = 0 Then
             strJoin = strJoin & " LEFT JOIN " & Session(CurrentField & "LookupTable") & _
             " on " & Session(CurrentField & "LookupTable") & "." & Session(CurrentField & "LookupField") & " = " & _
             OwnerTable(CurrentTable, Session("CurrentField")) & "." & Session(CurrentField & "Field")

           End If
        ElseIf OwnerTable(CurrentTable, Session("CurrentField")) <> CurrentTable Then
           strJoin = strJoin & MakeJoin(CurrentTable, OwnerTable(CurrentTable, Session("CurrentField")), strJoin)

        End If
     End If
    Next

' (Joga o filtro no WHERE)
    strWhere = Filter

    strGroup = " GROUP BY " & sstrFields2

' (Monta o ORDER BY do relatório)
    If Not Session("OrderField") > "" Then
      BeginReadOperation sstrFields

      Do While IsThereOperation
        If FindField(CurrentTable, NextOperation) Then
          sstrText = ""
          Session("OrderField") = Session("CurrentField")
          strOrder = " ORDER BY " & Session("CurrentField")

        End If

      Loop

    End If

    Dim blnFirst As Boolean
    If IsNumeric(Session("OrderField")) Then
      BeginReadOperation Session("OrderField")
      strOrder = " ORDER BY "

      blnFirst = True

      Do While IsThereOperation
        Select Case NextOperation
           Case ","

           Case Else
             Session("CurrentField") = LastOperation
             If Not blnFirst Then
               strOrder = strOrder & ", "

             Else
               strOrder = strOrder & " "

             End If

             Select Case Session(CurrentField & "EditOption")
               Case efFldTypeLookup
                 strOrder = strOrder & Session(CurrentField & "LookupTable") & "." & Session(CurrentField & "LookupFieldText")

               Case Else
                 strOrder = strOrder & Session(CurrentField & "Field")

            End Select

            blnFirst = False
            If Session("OrderDesc") = "1" Then
              strOrder = strOrder & " DESC "

            End If

        End Select

      Loop

    Else
      If Session("OrderField") > "" Then
        strOrder = " ORDER BY " & Session("OrderField")
        If Session("OrderDesc") = "1" Then
          strOrder = strOrder & " DESC "

        End If
      End If
    End If

    Dim intWidth As Integer

    intWidth = 0

    For i = 1 To Session("FieldNumber")
      Session("CurrentField") = Session(i & "ShowField")
      intWidth = intWidth + EditSize + 5

    Next

    If intWidth < 70 Then
      intWidth = 70

    ElseIf intWidth > 95 Then
      intWidth = 95

    End If

    Browse.Actual = CurrentTable & "Grid"

    FontStyles

    strSelect = "SELECT "

    strFrom = " FROM " & CurrentTable

' (Calcula o total da soma das colunas)
    blnFirst = True
    If sblnHaveSum Or sblnHaveGraph Then
      For i = 1 To Session("FieldNumber")
        If Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Then
          If Not blnFirst Then
            sql = sql & ", "

          End If
          blnFirst = False

          If Session(i & "IsSum") > "" Then
            sql = sql & " SUM(" & Session(i & "Field") & ") as s" & Session(i & "Field")

          Else
            sql = sql & " COUNT(*) as s" & Session(i & "Field")

          End If

        End If
        Session(i & "MaxValue") = 0
        Session(i & "TotalValue") = 0

      Next

      sql = strSelect & sql & strFrom & strWhere

      Rem sobjCMD2.CommandText = sql
      Rem sobjCMD2.CommandType = adCmdText

Rem      sobjShow.MessageError sql

      Select Case ConnectionType
      Case efConnDBDriverOracle, efConnDBProviderOracle
        sql = FixJoinForOracle(sql)
      End Select

      'Alterado/Criado (Ridai)
      Set CurrConnection = sobjConn

      If ShowSQLQuery Then ShowCurrentSQL sql, "ShowSummary", adOpenKeyset, adLockReadOnly

      On Error Resume Next

      sobjRS2.Open sql, sobjConn, adOpenKeyset, adLockReadOnly

      If Err.number <> 0 Then
        If DebugMode Then
          ShowSQLErrors sql, "ShowSummary", adOpenKeyset, adLockReadOnly
        Else
          ShowDefaultError
        End If
        Err.Clear
        Exit Sub

      End If
      On Error GoTo 0
      '----

      For i = 1 To Session("FieldNumber")
        If Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Then
          Session(i & "TotalValue") = sobjRS2("s" & Session(i & "Field"))

        End If

      Next
      sobjRS2.Close

    End If

    sql = strSelect & sstrFields & strFrom & strJoin & strWhere & strGroup & strOrder

Rem      sobjShow.MessageError sql
    Rem sobjCMD2.CommandText = sql
    Rem sobjCMD2.CommandType = adCmdText

    Select Case ConnectionType
    Case efConnDBDriverOracle, efConnDBProviderOracle
      sql = FixJoinForOracle(sql)
      sobjRS2.CursorLocation = adUseClient
    End Select

    'Alterado/Criado (Ridai)
    Set CurrConnection = sobjConn2

    If ShowSQLQuery Then ShowCurrentSQL sql, "ShowSummary", adOpenKeyset, adLockReadOnly

    On Error Resume Next

    sobjRS2.Open sql, sobjConn2, adOpenKeyset, adLockReadOnly

    If Err.number <> 0 Then
      If DebugMode Then
        ShowSQLErrors sql, "ShowSummary", adOpenKeyset, adLockReadOnly
      Else
        ShowDefaultError
      End If
      Err.Clear
      Exit Sub

    End If
    On Error GoTo 0
    '----

' (Calcula o valor máximo das colunas)
    If Not sobjRS2.EOF And Not sobjRS2.BOF Then
    Dim strMessage As String
      If sblnHaveGraph Then
        Do While Not sobjRS2.EOF
          For i = 1 To Session("FieldNumber")
            If Session(i & "IsGraph") > "" Then
              If Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Then
                If Session(i & "MaxValue") < sobjRS2("s" & Session(i & "Field")) Then
                  Session(i & "MaxValue") = sobjRS2("s" & Session(i & "Field"))

                End If
              Else
                If Session(i & "MaxValue") < sobjRS2(Session(i & "Field")) Then
                  Session(i & "MaxValue") = sobjRS2(Session(i & "Field"))

                End If
              End If
            End If
          Next
          sobjRS2.MoveNext

        Loop
        sobjRS2.MoveFirst

      End If

      strMessage = Session(CurrentList & "ListCaption")

      If ReportTitle > "" Then strMessage = ReportTitle

      sobjShow.Nbsp
      sobjShow.Center

      Browse.BeginBrowse sobjRS2

      Table.BeginTable intWidth & "%", strMessage, intNumCols, True

' (Monta os tips com a listagem e todos os relatórios)
      strTip2 = strTip2 & "&nbsp;&nbsp;&nbsp;<A CLASS=SS HRef=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & _
      efQSActionList & WhatList & GetFind & """>" & LanguageLocalizer(65) & "</A><BR>"
      For i = 1 To Session(CurrentTable & "ListNumber")
        If Session(CurrentTable & i & "ListName") <> sparEditList Then
          strTip2 = strTip2 & "&nbsp;&nbsp;&nbsp;<A CLASS=SS HRef=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & _
          efQSActionSummary & WhatList & GetFind & "&" & efQueryStrList & "=" & Session(CurrentTable & i & "ListName") & """>" & Session(CurrentTable & i & "ListCaption") & "</A><BR>"

        End If

      Next

      Table.ColumnColSpan = intNumCols

      Table.BeginRow 1
      Table.BeginColumn
      Table.BeginTable "100%"
      Table.ColumnAlign = "Right"

      Table.BeginRow 1
      Table.BeginColumn
      Response.Write "<FONT SIZE=2><DIV ID=F2 onMouseOver=""makeCool(F2);"" onMouseOut=""makeNormal(F2);"" onClick=""OpenDesc(T2);"" STYLE=""cursor:hand""><B>" & LanguageLocalizer(66) & "</B></DIV>"
      Response.Write "<table><tr><td align=justify>"
      Response.Write "<DIV ID=T2 class=dbody style=""display:'none';""><B>"
      Response.Write strTip2
      Response.Write "</B></DIV></td></tr></table>"
      Table.EndColumn
      Table.EndRow

      Table.EndTable
      Table.EndColumn
      Table.EndRow
      Table.ColumnAlign = ""
      Table.ColumnColSpan = ""

' (Monta a linha de cabecalho)
      Table.BeginRow 1, True
      For i = 1 To Session("FieldNumber")
        If Session(i & "IsGraph") > "" Then
          TableSummaryLegend Session(i & "MaxValue"), Session(i & "TotalValue"), Session(i & "CaptionField"), Session(i & "Type"), Session(i & "OrderSQL")

        ElseIf Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Then
          EditTableHeader Session(i & "CaptionField"), Session(i & "OrderSQL"), Session(CurrentField & "ListSize")

        Else
          Session("CurrentField") = Session(i & "ShowField")
          EditTableHeader Session(i & "CaptionField"), Session(i & "OrderSQL"), Session(CurrentField & "ListSize")

        End If

        If sstrGroupField > "" Then
          If Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Or Session(i & "IsTotal") > "" Then
            Session(i & "SubTotal") = 0
            Session(i & "Total") = 0

          End If
        End If

      Next
      Table.EndRow

' (Apresenta as linhas de dados do relatório)
      sstrGroupFieldLastValue = "kapushinefrebow"
      blnFirst = True
      Do While Not sobjRS2.EOF And Browse.IsBrowsePage
' (Apresenta a linha de sub-total)
        If sstrGroupField > "" Then
          If sstrGroupFieldLastValue <> "kapushinefrebow" Then
            If sstrGroupFieldLastValue <> sobjRS2(Session(sstrGroupField & "Field")) Then
              ShowSummarySubTotal intNumCols

            End If
          End If

          sstrGroupFieldLastValue = sobjRS2(Session(sstrGroupField & "Field"))

        End If

        Table.BeginRow 1

' (Busca os valores dos campos que não são de operação)
        For i = 1 To Session("FieldNumber")
          If Not Session(i & "Operation") > "" Or IsNull(Session(i & "Operation")) Then
            If Session(i & "IsCount") > "" Or Session(i & "IsSum") > "" Then
              Session(i & "Value") = sobjRS2("s" & Session(i & "Field"))

            Else
              Session(i & "Value") = sobjRS2(Session(i & "Field"))

            End If
          End If
        Next

        For i = 1 To Session("FieldNumber")
' (Calcula os campos de operação)
          If Session(i & "Operation") > "" Then
            If Not blnFirst And Session(i & "Have1stValue") > "" Then
              Session(i & "Value") = SolveOperation(i)

            ElseIf Not Session(i & "Have1stValue") > "" Then
              Session(i & "Value") = SolveOperation(i)

            End If
            blnFirst = False

          End If

' (Calcula os sub-totais)
          If sstrGroupField > "" Then
            If Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Or Session(i & "IsTotal") > "" Then
              Session(i & "SubTotal") = Session(i & "SubTotal") + Session(i & "Value")

            End If

          End If

' (Se o campo for de gráfico)
          If Session(i & "IsGraph") > "" Then
            Table.ColumnGraph Session(i & "Value"), "<a href=" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionList & WhatList & "&" & efQueryStrFilter & "=" & MakeFilter(sobjRS2, sstrFields2) & ">" & FormatVariable(Session(i & "Value"), Session(i & "Type")) & "</a>"

' (Se o campo for calculado, somatória ou quantidade)
          Else
            If Session(i & "IsCalcField") > "" Or Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Then
              Table.ColumnAlign = "Right"

              Table.BeginColumn
              Response.Write FormatVariable(Session(i & "Value"), Session(i & "Type"))
              Table.EndColumn
              Table.ColumnAlign = ""

' (Se o campo for comum)
            Else
              Session("CurrentField") = Session(i & "ShowField")
              Select Case TypeName(EditFieldValue(sobjRS2))
                Case "Long", "Integer", "Currency", "Double", "Single"
                  Table.ColumnAlign = "Right"

              End Select

              Table.BeginColumn
              Select Case Session(CurrentField & "EditOption")
                Case efFldTypeLookup, efFldTypeRadio, efFldTypeCheck, efFldTypeSelect
                  Response.Write sobjString.BreakString(EditFieldValue(sobjRS2), Session(CurrentField & "ListChars"))

                Case Else
                  Response.Write sobjString.BreakString(sobjString.Mask(FormatVariable(EditFieldValue(sobjRS2), Session(CurrentField & "Type")), Session(CurrentField & "Mask"), Session(CurrentField & "MaskCompleter")), Session(CurrentField & "ListChars"))

              End Select
              Table.ColumnAlign = ""
              Table.EndColumn

            End If
          End If

        Next
        Table.EndRow

        sobjRS2.MoveNext
        Browse.IncLine

      Loop

' (Apresenta o total)
      If sstrGroupField > "" Then
      Dim strColor As String
        ShowSummarySubTotal intNumCols

        strColor = Table.RowColor
        Table.RowColor = Table.HeadRowColor
        Table.BeginRow 1
        For i = 1 To Session("FieldNumber")
          If Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Or Session(i & "IsTotal") > "" Then
            If Session(i & "IsGraph") > "" Then
              Table.BeginColumn
              Table.BeginTable
              Table.BeginRow 1
              Table.ColumnWidth = "70"
              Table.ColumnAlign = "Right"
              Table.Column FormatVariable(Session(i & "Total"), Session(i & "Type")) & " "
              Table.EndRow
              Table.EndTable
              Table.ColumnWidth = ""
              Table.ColumnAlign = ""

              Table.EndColumn

            Else
              Table.ColumnAlign = "Right"
              Table.Column FormatVariable(Session(i & "Total"), Session(i & "Type"))
              Table.ColumnAlign = ""

            End If

          Else
            Table.Column "&nbsp;"

          End If
        Next
        Table.EndRow
        Table.RowColor = strColor

      End If

' (Limpa os Sessions utilizados)
      For i = 1 To Session("FieldNumber")
        ClearSession i & "Total"
        ClearSession i & "SubTotal"
        ClearSession i & "IsCalcField"
        ClearSession i & "IsSum"
        ClearSession i & "IsTotal"
        ClearSession i & "IsCount"
        ClearSession i & "Field"
        ClearSession i & "CaptionField"
        ClearSession i & "ShowField"
        ClearSession i & "OrderSQL"
        ClearSession i & "MaxValue"
        ClearSession i & "TotalValue"
        ClearSession i & "Operation"
        ClearSession i & "Type"
        ClearSession i & "Have1stValue"
        ClearSession i & "Value"

      Next
      Table.EndTable

      Form.EndForm

      Browse.Status = True
      Browse.EndBrowse
      Browse.Bar

      sobjShow.CenterEnd

' (Se não houver registros no SELECT)
    Else
      sobjShow.MessageError "<center>" & sobjString.FormatText(LanguageLocalizer(35), Session(CurrentList & "Caption")) & "</center><P>"

    End If

    sobjShow.BR
    sobjRS2.Close

    FPLPLink

' (Se o relatório não for encontrado)
  Else
    sobjShow.MessageError sobjString.FormatText(LanguageLocalizer(67), sparEditList)

  End If

End Sub

Private Function EditNextMathData() As String
' (Retorna as operações e colunas das operações do relatório)
Dim strSubText, strChar As String
Dim blnEnd As Boolean

  strChar = Left(sstrText, 1)
  blnEnd = False

  Do
    Select Case strChar
      Case "-", "+", "*", "/", ""
        If strSubText = "" Then
          strSubText = strChar
          sstrText = Mid(sstrText, 2)

        End If
        blnEnd = True

      Case Else
        strSubText = strSubText & strChar
        sstrText = Mid(sstrText, 2)
        strChar = Left(sstrText, 1)

    End Select

   Loop While Not blnEnd

   EditNextMathData = strSubText
   sstrSubText = strSubText

End Function

Private Function SolveOperation(ByVal intField As Integer) As Double
' (Calcula as operações do relatório)
Dim dblTotal As Double, strOperator As String

  dblTotal = 0
  BeginReadOperation Session(intField & "Operation")
  strOperator = "+"
  Do While IsThereOperation
    Select Case EditNextMathData
' (Guarda o operador)
       Case "+", "-", "/", "*"
         strOperator = LastOperation

       Case Else
' (Calcula a partir de um operando numérico)
         If IsNumeric(LastOperation) Then
           Select Case strOperator
             Case "+"
               dblTotal = dblTotal + CDbl(LastOperation)

             Case "-"
               dblTotal = dblTotal - CDbl(LastOperation)

             Case "*"
               dblTotal = dblTotal * CDbl(LastOperation)

             Case "/"
               dblTotal = dblTotal / CDbl(LastOperation)

           End Select

' (Calcula a partir de um operando coluna)
         ElseIf Len(LastOperation) = 1 Then
           Select Case strOperator
             Case "+"
               dblTotal = dblTotal + Session((Asc(UCase(LastOperation)) - 64) & "Value")

             Case "-"
               dblTotal = dblTotal - Session((Asc(UCase(LastOperation)) - 64) & "Value")

             Case "*"
               dblTotal = dblTotal * Session((Asc(UCase(LastOperation)) - 64) & "Value")

             Case "/"
               dblTotal = dblTotal / Session((Asc(UCase(LastOperation)) - 64) & "Value")

           End Select

        End If

    End Select

  Loop
  SolveOperation = dblTotal

End Function

Private Sub ShowSummarySubTotal(ByVal intNumCols As Integer)
' (Apresenta os sub-totais)
Dim strColor As String, i As Integer
  strColor = Table.RowColor
  Table.RowColor = "WHITE"
  Table.BeginRow 1
  For i = 1 To Session("FieldNumber")
    If Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Or Session(i & "IsTotal") > "" Then
      If Session(i & "IsGraph") > "" Then
        Table.BeginColumn
        Table.BeginTable
        Table.BeginRow 1
        Table.ColumnWidth = "70"
        Table.ColumnAlign = "Right"
        Table.Column FormatVariable(Session(i & "SubTotal"), Session(i & "Type")) & "&nbsp;"
        If Session(i & "TotalValue") > 0 Then
          Table.ColumnWidth = "50"
          Table.Column "&nbsp;" & sobjString.ZTINumber(100 * Session(i & "SubTotal") / Session(i & "TotalValue")) & "%" & "&nbsp;"

        End If
        Table.ColumnWidth = ""
        Table.EndRow
        Table.EndTable
        Table.ColumnAlign = ""

        Table.EndColumn
        Session(i & "Total") = Session(i & "Total") + Session(i & "SubTotal")
        Session(i & "SubTotal") = 0

      Else
        Table.ColumnAlign = "Right"
        Table.Column FormatVariable(Session(i & "SubTotal"), Session(i & "Type"))
        Session(i & "Total") = Session(i & "Total") + Session(i & "SubTotal")
        Session(i & "SubTotal") = 0
        Table.ColumnAlign = ""

      End If
    Else
      Table.Column "&nbsp;"

    End If

  Next
  Table.EndRow

  Table.RowColor = strColor
  Table.ColumnColSpan = intNumCols
  Table.BeginRow 1
  Table.Column "&nbsp;"
  Table.EndRow
  Table.ColumnColSpan = ""

End Sub

Private Sub TableSummaryLegend(ByVal intMax As Double, ByVal intTotal As Double, ByVal strText As String, ByVal strType As String, ByVal strOrder As String)
' (Apresenta a legenda de um gráfico de relatório)
Dim strCor, strCorFont As String
Dim i As Integer

  Table.LegendInfo intMax, intTotal

  Table.Padding = 0
  Table.Spacing = 0
  strCor = Table.RowColor
  Table.RowColor = Table.HeadRowColor
  strCorFont = Table.Style.InternalFont.Color
  Table.Style.InternalFont.Color = Table.Style.HeaderFont.Color

  Table.ColumnAlign = "Left"
  Table.BeginColumn
  Table.BeginTable

  If Table.GraphTotal > 0 Then
    Table.CellGraphTotalWidth = 120
    Table.ColumnWidth = 120

  Else
    Table.CellGraphTotalWidth = 70
    Table.ColumnWidth = 70

  End If

  Table.BeginRow 1
  Table.BeginColumn

' (Apresenta o link de ordenação da coluna de gráfico)
  If LCase(sobjString.ToString(Session("OrderField"))) = LCase(sobjString.ToString(strOrder)) Then
    If Session("OrderDesc") = "1" Then
      sobjShow.Image Initializer.HiperToolsWorkPath & "/ArrowDown.gif", "", sobjString.FormatText(LanguageLocalizer(37), EditCaption(False)), Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionSummary & WhatList & GetFind & "&" & efQueryStrOrderField & "=" & Server.URLEncode(strOrder) & EditListStr & GetFilter, "", 16, 9
      sobjShow.Nbsp
      Response.Write "<A CLASS=""HL"" HRef=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionSummary & WhatList & GetFind & "&" & efQueryStrOrderField & "=" & Server.URLEncode(strOrder) & EditListStr & GetFilter & """>" & strText & "</A>"

    Else
      sobjShow.Image Initializer.HiperToolsWorkPath & "/ArrowUp.gif", "", sobjString.FormatText(LanguageLocalizer(37), EditCaption(False)), Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionSummary & WhatList & GetFind & "&" & efQueryStrOrderField & "=" & Server.URLEncode(strOrder) & EditListStr & efQueryStrOrderDescStr & GetFilter, "", 16, 9
      sobjShow.Nbsp
      Response.Write "<A CLASS=""HL"" HREF=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionSummary & WhatList & GetFind & "&" & efQueryStrOrderField & "=" & Server.URLEncode(strOrder) & EditListStr & efQueryStrOrderDescStr & GetFilter & """>" & strText & "</A>"

    End If

  Else
    Response.Write "<A CLASS=""HL"" HREF=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionSummary & WhatList & GetFind & "&" & efQueryStrOrderField & "=" & Server.URLEncode(strOrder) & EditListStr & GetFilter & """>" & strText & "</A>"

  End If

  Table.ColumnAlign = "Left"
' (Se o valor máximo do gráfico for maior que 2)
  If intMax > 2 Then
    Table.CellGraphMeterWidth = 180
    Table.ColumnWidth = 60
    Table.ColumnAlign = "Left"
    Table.Column "0"
    Table.ColumnAlign = "Center"
    If intMax < 10 Then
      Table.Column FormatVariable(Round(intMax / 2, 1), strType)

    Else
      Table.Column FormatVariable(intMax \ 2, strType)

    End If

    Table.ColumnAlign = "Right"
    Table.Column FormatVariable(intMax, strType)
    Table.EndRow

    Table.BeginRow 1

    Table.ColumnWidth = 60
    If intTotal > 0 Then
       Table.Column "&nbsp;<Font Size=-1><B>" & FormatVariable(Table.GraphTotal, strType) & "<B></Font>&nbsp;"

    Else
       Table.Column "&nbsp;"

    End If

    Table.ColumnWidth = 60
    Table.ColumnAlign = "Left"
    Table.Column "|"
    Table.ColumnAlign = "Center"
    Table.Column "|"
    Table.ColumnAlign = "Right"
    Table.Column "|"

' (Se o valor máximo do gráfico não for maior que 2)
  Else
    Table.ColumnWidth = 180 \ (intMax + 1)
    Table.ColumnAlign = "Left"
    For i = 0 To intMax - 1
      Table.Column FormatVariable(i, strType)
      Table.ColumnAlign = "Center"

    Next

    Table.ColumnAlign = "Right"
    Table.Column FormatVariable(intMax, strType)
    Table.EndRow
    Table.ColumnAlign = "Center"

    Table.BeginRow 1

    Table.ColumnWidth = 120
    If intTotal > 0 Then
      Table.Column "&nbsp;<B>" & FormatVariable(Table.GraphTotal, strType) & "<B>&nbsp;"

    Else
       Table.Column "&nbsp;"

    End If

    Table.ColumnAlign = "Left"
    Table.ColumnWidth = 180 \ (intMax + 1)
    For i = 0 To intMax - 1
      Table.Column "|"
      Table.ColumnAlign = "Center"

    Next

    Table.ColumnAlign = "Right"
    Table.Column "|"

  End If
  Table.EndRow

  Table.EndTable

  Table.EndColumn

  Table.ColumnAlign = ""
  Table.ColumnWidth = ""
  Table.RowColor = strCor
  Table.Style.InternalFont.Color = strCorFont

End Sub

Public Function IsGridAction() As Boolean
' (Se a ação é de montagem de grid)
  IsGridAction = (sparEditAction = efQSActionGrid)

End Function

Private Function IsExternalField() As Boolean
' (Se o campo é atual é de uma tabela diferente da atual)
  IsExternalField = Session(CurrentField & "ExternalAddress")

End Function

Private Function HaveTabs() As Boolean
' (Se o formulário de edição possui tabs)
  HaveTabs = (InStr(Session(CurrentTable & "FieldUnit"), "|") > 0)

End Function

'Criado (Ridai Govinda)
Public Sub CreateTable(ByVal strTable As String)
  TableCreate strTable
End Sub
'---------------------

Public Sub TableCreate(ByVal strTable As String)
' (Cria uma tabela no BD)
Dim sql
Dim i, strType

  sql = ""

  strTable = Trim(strTable)

' (Se a tabela existir então exclui)
  If IsTable(strTable) Then
    TableDrop strTable

  End If

  'Alterado (Ridai)
  If Session(strTable & "FieldNumber") > "" Then
  ' (Cada campo da tabela)
    For i = 1 To Session(strTable & "FieldNumber")
      If i > 1 Then
        sql = sql & ", "

      End If

      sql = sql & Session(strTable & i & "Field") & " "

      strType = LCase(Session(strTable & i & "Type"))

      Select Case strType
  ' (Campos que não tem o tamanho definido pelo CREATE TABLE)
        Case "int", "money", "numeric", "datetime", "float", "real", "tinyint", "smallint", "bit", "text"
          'Alterado (Ridai)
          sql = sql & DataTypeByConnection(strType)
          '---------------

        'Criado (Ridai Govinda)
        'Caso específico se o Banco de Dados for Access
        ' não pode-se definir tamanho, então ignora-se o tamanho:
        Case "text"
          Select Case ConnectionType
          Case efConnDBDriverAccess, efConnDBProviderJet35, efConnDBProviderJet40
            sql = sql & DataTypeByConnection(strType)

          Case Else
            sql = sql & DataTypeByConnection(strType) & "(" & Session(strTable & i & "Size") & ")"

          End Select
        '------------
  ' (Campos que tem o tamanho definido pelo CREATE TABLE)
        Case Else
          'Alterado (Ridai)
          sql = sql & DataTypeByConnection(strType) & "(" & Session(strTable & i & "Size") & ")"
          '---------------

      End Select

  ' (Se o campo é requerido)
      If Not Session(strTable & i & "Null") Or strType = "bit" Then
        sql = sql & " NOT"

      End If

      sql = sql & " NULL"

    Next

  ' (Criação)
    sql = "CREATE TABLE " & strTable & " (" & sql & ")"
    Database.ConnExecute sobjConn, sql
    sobjShow.Message LanguageLocalizer(69) & strTable
    IndexCreate strTable
  Else
    'Alterado (Ridai Govinda)
    ShowError efErrorTableNoStruct, "EditForm", "TableCreate", LanguageLocalizer(10)
    '-----------------------
  End If

End Sub

Public Sub DropTable(ByVal strTable As String)
  TableDrop strTable
End Sub
'---------------------

Public Sub TableDrop(ByVal strTable As String)
' (Exclui uma tabela no BD)

  On Error Resume Next
  Database.ConnExecute sobjConn, "DROP TABLE " & strTable
  On Error GoTo 0
  sobjShow.Message LanguageLocalizer(70) & strTable

End Sub

'Alterado (Ridai)
Public Function HasTable(ByVal strTable As String) As Boolean
  HasTable = IsTable(strTable)
End Function

'Função nova utilizando ADOX
'Removida a função antiga de verificação de existência de Tabela no Banco...

Public Function IsTable(ByVal strTable As String) As Boolean

  Dim objCatalog As ADOX.Catalog, objTable As ADOX.Table

  Set objCatalog = Server.CreateObject("ADOX.Catalog")

  objCatalog.ActiveConnection = ConnectionString

  IsTable = False

  For Each objTable In objCatalog.Tables
    If Trim(LCase(strTable)) = Trim(LCase(objTable.Name)) Then IsTable = True
  Next

  Set objCatalog = Nothing

End Function

Public Sub TableUpdateBegin(ByVal strTable As String)
' (Inicia uma operação de UPDATE no BD)
  CurrentTable = strTable
  Session("FieldNumberUpdate") = 0
  Session("IndexNumberUpdate") = 0

End Sub

Public Sub FieldUpdateValue(ByVal strField As String, ByVal strValue As String)
' (Altera o valor de um campo)
  Session("FieldNumberUpdate") = Session("FieldNumberUpdate") + 1
  Session("FieldUpdate" & Session("FieldNumberUpdate")) = strField
  Session("FieldValue" & Session("FieldNumberUpdate")) = strValue

End Sub

Public Sub FieldUpdateIndex(ByVal strField As String, ByVal strOperator As String, ByVal strValue As String)
' (Altera um índice)
  Session("IndexNumberUpdate") = Session("IndexNumberUpdate") + 1
  Session("FieldIndex" & Session("IndexNumberUpdate")) = strField
  Session("IndexOperator" & Session("IndexNumberUpdate")) = strOperator
  Session("IndexValue" & Session("FieldNumberUpdate")) = strValue

End Sub

Public Sub TableUpdateEnd()
' (Executa o UPDATE montado)
Dim objRS As Object
Dim sql
Dim i

  sql = "UPDATE " & CurrentTable & " SET "

  For i = 1 To Session("FieldNumberUpdate")
    If i > 1 Then
      sql = sql & " AND "
    End If
    sql = sql & Session("FieldUpdate" & i) & " = " & Session("FieldValue" & i)
  Next

' (Altera os índices)
  If Session("IndexNumberUpdate") > 0 Then

    Database.AndState = False

    For i = 1 To Session("IndexNumberUpdate")
      sql = sql & Database.PutAnd & Session("FieldIndex" & i) & " " & Session("IndexOperator" & i) & " " & Session("IndexValue" & i)
    Next

  End If

' (Executa o UPDATE)
  Rem sobjCMD.CommandText = sql
  Rem sobjCMD.CommandType = adCmdText

  Set objRS = Server.CreateObject("ADODB.RecordSet")

  'Alterado/Criado (Ridai)
  Set CurrConnection = sobjConn

  If ShowSQLQuery Then ShowCurrentSQL sql, "TableUpdateEnd", adOpenDynamic, adLockPessimistic

  On Error Resume Next

  objRS.Open sql, sobjConn, adOpenDynamic, adLockPessimistic

  If Err.number <> 0 Then
    If DebugMode Then
      ShowSQLErrors sql, "TableUpdateEnd", adOpenDynamic, adLockPessimistic
    Else
      ShowDefaultError
    End If
    Err.Clear
    Exit Sub

  End If
  On Error GoTo 0
  '----

End Sub

'Criado (Ridai Govinda)
Public Sub CreateIndex(ByVal strTable As String)
  IndexCreate strTable
End Sub
'------------------

Public Sub IndexCreate(ByVal strTable As String)
' (Cria os índices de uma tabela)
Dim i
Dim strCreateIndex As String

  ' Alterado (Ridai)
  '(Não existe em Access/ não sei se existe nos outros tipos de banco)
  ' (Cria o índice CLUSTERED)
  If Session(strTable & "IndexClusteredName") > "" And ConnectionType = efConnDBDriverSQL Then
  '--------------
    On Error Resume Next
Rem    Database.ConnExecute sobjConn, "DROP INDEX " & strTable & "." & strTable & "_" & Session(strTable & "IndexClusteredName")

    On Error GoTo 0

    Database.ConnExecute sobjConn, "CREATE CLUSTERED INDEX " & strTable & "_" & Session(strTable & "IndexClusteredName") & " ON " & _
              strTable & "(" & Session(strTable & "IndexClusteredContent") & ")"

    sobjShow.Message LanguageLocalizer(71) & strTable & "_" & Session(strTable & "IndexClusteredName")

  End If

' (Cria os índices comuns)
  If Session(strTable & "IndexNumber") > 0 Then
    For i = 1 To Session(strTable & "IndexNumber")
      On Error Resume Next
Rem      Database.ConnExecute sobjConn, "DROP INDEX " & strTable & "." & strTable & "_" & Session(strTable & "IndexName")

      On Error GoTo 0

' Alterado por Ueslei/Kelson (Criação de UNIQUE INDEX)
      strCreateIndex = "CREATE "
      If Session(strTable & i & "IndexUnique") Then
        strCreateIndex = strCreateIndex & "UNIQUE "

      End If
      Database.ConnExecute sobjConn, strCreateIndex & "INDEX " & strTable & "_" & Session(strTable & i & "IndexName") & " ON " & _
                strTable & "(" & Session(strTable & i & "IndexContent") & ")"

      On Error Resume Next

      sobjShow.Message LanguageLocalizer(71) & strTable & "_" & Session(strTable & i & "IndexName")

      'Criado (Ridai Govinda)
      On Error GoTo 0
      '-------------------

    Next
  End If

End Sub

' (Estas últimas três rotinas não estão funcionando por fazerem parte do Grid. Por não estarem
' funcionando e pelo formato "atual" do grid ser efetivo, provavelmente serão substituídas)
' (A idéia para o novo formato esta explicado no arquivo "Relacionamentos de tabelas do EditForm
' .doc" no diretório info)
Private Sub ShowGrid()
Dim sql, i, strPrincipalTable, intAction, strURL
Rem  DebugMode

  FontStyles
  strPrincipalTable = CurrentTable
  CurrentTable = Session(CurrentTable & CurrentGrid & "GridTable")
  Session(strPrincipalTable & "GridWidth") = ""

  sobjShow.HTMLCR "<STYLE> .EF{font: 8pt Arial}</STYLE>"
  Form.Class = "EF"
  If Request.Form("btmEdit") > "" Then
    intAction = 1

  ElseIf Request.Form("btmSave") > "" Then
    intAction = 0
    If sparEditFind > "" Then
      GetForm Session(strPrincipalTable & CurrentGrid & "GridFields")
      BeginReadOperation Session(strPrincipalTable & CurrentGrid & "GridFields")
      Do While IsThereOperation
        Select Case NextOperation
          Case ",", ";", "-"
          Case "="
            IsThereOperation

          Case Else
            If FindField(CurrentTable, LastOperation) Then
              Session(CurrentField & "OldValue") = Session(CurrentField & "Value")

            End If

         End Select
      Loop

      sql = "SELECT * FROM " & CurrentTable & " WHERE " & sparEditFind

      Rem sobjCMD2.CommandText = sql
      Rem sobjCMD2.CommandType = adCmdText

      'Alterado/Criado (Ridai)
      Set CurrConnection = sobjConn2

      If ShowSQLQuery Then ShowCurrentSQL sql, "ShowGrid", adOpenDynamic, adLockOptimistic

      On Error Resume Next

      sobjRS2.Open sql, sobjConn2, adOpenDynamic, adLockOptimistic

      If Err.number <> 0 Then
        If DebugMode Then
          ShowSQLErrors sql, "ShowGrid", adOpenDynamic, adLockOptimistic
        Else
          ShowDefaultError
        End If
        Err.Clear
        Exit Sub

      End If
      On Error GoTo 0
      '----

      If sobjRS2.EOF Or sobjRS2.BOF Then
        sobjRS2.AddNew

      End If

      For i = 1 To Session(CurrentTable & "FieldNumber")
         sobjRS2(Session(CurrentTable & i & "Field")) = FormatFieldtoSave(Session(CurrentTable & _
                                                                  i & "OldValue"), Session(CurrentTable & i & "Type"), _
                                                                  Session(CurrentTable & i & "Size"), _
                                                                  Session(CurrentTable & i & "CharCase"))

      Next

      If ShowSQLQuery Then ShowCurrentSQL sql & " [Update / UPDATE]", "ShowGrid", adOpenDynamic, adLockOptimistic

      On Error Resume Next

      sobjRS2.Update

      If Err.number <> 0 Then
        If DebugMode Then
          ShowSQLErrors sql & " [Update / UPDATE]", "ShowGrid", adOpenDynamic, adLockOptimistic
        Else
          ShowDefaultError
        End If
        Err.Clear
        Exit Sub

      End If
      On Error GoTo 0
      '----

      sobjRS2.Close
      sparEditFind = ""
      Application.Lock
      Application(CurrentTable & MakeTarget) = ""
      Application.Unlock

    End If

  ElseIf Request.Form("btmDelete") > "" Then
    intAction = 0
    If sparEditFind > "" Then
      sql = "DELETE FROM " & CurrentTable & " WHERE " & sparEditFind
      Rem sobjCMD2.CommandType = adCmdText

      'Alterado/Criado (Ridai)
      If ShowSQLQuery Then ShowCurrentSQL sql, "ShowGrid", adOpenKeyset, adLockReadOnly

      On Error Resume Next

      sobjRS2.Open sql, sobjConn2, adOpenKeyset, adLockReadOnly

      If Err.number <> 0 Then
        If DebugMode Then
          ShowSQLErrors sql, "ShowGrid", adOpenKeyset, adLockReadOnly
        Else
          ShowDefaultError
        End If
        Err.Clear
        Exit Sub

      End If
      On Error GoTo 0
      '----

    End If
    sparEditFind = ""

  ElseIf Request.Form("btmInsert") > "" Or sparEditGrid = "1" Then
    intAction = 2
    sparEditFind = ""

  ElseIf Request.Form("btmCancel") > "" Then
    intAction = 0
    Application.Lock
    Application(CurrentTable & MakeTarget) = ""
    Application.Unlock

  Else
    intAction = 0

  End If

  If FindRelation(strPrincipalTable, CurrentTable) Or strPrincipalTable = CurrentTable Then
    sql = "SELECT * FROM " & CurrentTable

    BeginReadOperation Session(strPrincipalTable & CurrentGrid & "GridFields")

    Dim Width, strList
    Width = 0
    Do While IsThereOperation
      Select Case NextOperation
         Case ",", ";", "-"
         Case "="
           IsThereOperation

         Case Else
           If FindField(CurrentTable, LastOperation) Then
             If Not Session(strPrincipalTable & "GridWidth") > "" Then
               If Session(CurrentField & "Type") <> "varchar" Then
                 Width = Width + EditSize * 5

               Else
                 Width = Width + EditSize * 6

               End If

             End If
             If strList > "" Then
               strList = strList & ","

             End If
             strList = strList & Session("CurrentField")
             If Session(CurrentTable & Session("CurrentField") & "EditOption") = efFldTypeLookup Then
                sql = sql & " LEFT JOIN " & Session(CurrentTable & Session("CurrentField") & "LookupTable") & _
                " on " & Session(CurrentTable & Session("CurrentField") & "LookupTable") & "." & Session(CurrentTable & Session("CurrentField") & "LookupField") & " = " & _
                OwnerTable(CurrentTable, i) & "." & Session(CurrentTable & Session("CurrentField") & "Field")

             End If

           End If

       End Select

    Loop

    If Not Session(strPrincipalTable & "GridWidth") > "" Then
      If Width < 200 Then
        Width = 200

      ElseIf Width > 600 Then
        Width = 600

      End If
      Session(strPrincipalTable & "GridWidth") = Width

    End If

    Database.AndState = False

    For i = 1 To Session(strPrincipalTable & CurrentRelation & "FieldtoFieldNumber")
      If FindField(strPrincipalTable, Session(strPrincipalTable & CurrentRelation & "," & i & "RelationField1")) Then
          sql = sql & Database.PutAnd & Session(strPrincipalTable & CurrentRelation & "," & i & "RelationField2") & "=" & FormatForSQL(Session(strPrincipalTable & Session("CurrentField") & "Value"), Session(strPrincipalTable & Session("CurrentField") & "Type"))

      End If
    Next

    sobjConn.Close
    sobjConn.Provider = "SQLOLEDB.1"
    sobjConn.Open

    Rem sobjCMD2.ActiveConnection = sobjConn
    Rem sobjCMD2.CommandText = sql
    Rem sobjCMD2.CommandType = adCmdText

    sobjRS2.CursorLocation = adUseClient

    'Alterado/Criado (Ridai)
    Set CurrConnection = sobjConn

    If ShowSQLQuery Then ShowCurrentSQL sql, "ShowGrid", adOpenKeyset, adLockReadOnly

    On Error Resume Next

    sobjRS2.Open sql, sobjConn, adOpenKeyset, adLockReadOnly

    If Err.number <> 0 Then
      If DebugMode Then
        ShowSQLErrors sql, "ShowGrid", adOpenKeyset, adLockReadOnly
      Else
        ShowDefaultError
      End If
      Err.Clear
      Exit Sub

    End If
    On Error GoTo 0
    '----

    If Not sobjRS2.EOF And Not sobjRS2.BOF Or sparEditGrid = "1" Then
      Browse.BeginBrowse sobjRS2

      TotalOperation strList, "GridFieldNumber", False, ""

      Table.Padding = "1"
      Table.Spacing = "1"
      Table.BeginTable Session(strPrincipalTable & "GridWidth") + 20

      Table.BeginRow 2
      BeginReadOperation strList
      Do While IsThereOperation
        Select Case NextOperation
          Case ",", ";", "-", "="

          Case Else
              Session("CurrentField") = LastOperation
            Table.ColumnWidth = Session(CurrentField & "ListSize")
            Table.ColumnColor = Table.HeadRowColor
            Table.Column "<CENTER><B><SPAN CLASS=HL>" & EditCaption(False) & "</SPAN>"
            Table.ColumnColor = ""

         End Select

      Loop
      If intAction = 0 Then
        strURL = Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionGrid & GetFilter

        Form.BeginForm strURL, "frmEditGrid", "POST", "", ""
        Table.ColumnWidth = "40"
        Table.BeginColumn
        Form.Submit "btmInsert", " I "
        Table.EndColumn
        Form.EndForm

      Else
        Table.Column "&nbsp;"

      End If
      Table.EndRow

    Dim intPosition As Integer
    Dim blnFirst As Boolean
      If sparEditFind > "" And intAction <> 0 Then
        sobjRS2.Find sparEditFind
        intPosition = sobjRS2.AbsolutePosition
        blnFirst = True

      ElseIf intAction = 2 Then
        blnFirst = True

      Else
        blnFirst = False

      End If

      Dim strField As String

      Do
        If intPosition <> sobjRS2.AbsolutePosition Or blnFirst Then
            BeginReadOperation strList
          If (intAction <> 2 Or Not blnFirst) And Not sparEditGrid > "" Then
            For i = 1 To Session(CurrentTable & "FieldNumber")
              Session(CurrentTable & i & "Value") = sobjRS2(Session(CurrentTable & i & "Field"))

            Next

          Else
            EditClear efQSActionAdd
            For i = 1 To Session(strPrincipalTable & CurrentRelation & "FieldtoFieldNumber")
              If FindField(CurrentTable, Session(strPrincipalTable & CurrentRelation & "," & i & "RelationField2")) Then
                Session(CurrentField & "OldValue") = FieldValue(strPrincipalTable, Session(strPrincipalTable & CurrentRelation & "," & i & "RelationField1"))

              End If
            Next

          End If
          strURL = Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionGrid & GridTarget & GetFilter

          Form.BeginForm strURL, "frmEditGrid", "POST", "", ""

            Table.BeginRow 1

            Do While IsThereOperation
              Select Case NextOperation
                Case ",", ";", "-", "="
                Case Else
                  Session("CurrentField") = LastOperation
                  strField = Session("CurrentField")
                  If blnFirst Then
                    Table.BeginColumn
                    ShowField False, True, False, False
                    Table.EndColumn

                  Else
                  Select Case TypeName(EditFieldValue(sobjRS2))
                    Case "Long", "Integer", "Currency", "Double", "Single"
                      Table.ColumnAlign = "Right"

                  End Select

                  Table.ColumnWidth = Session(CurrentField & "ListSize")
                  Table.BeginColumn
                  If intAction = 0 Then
                    MakeGridLink strPrincipalTable

                  End If
                  Session("CurrentField") = strField
                  Select Case Session(CurrentField & "EditOption")
                    Case efFldTypeLookup, efFldTypeRadio, efFldTypeCheck, efFldTypeSelect
                      Response.Write sobjString.BreakString(EditFieldValue(sobjRS2), Session(CurrentField & "ListChars"))

                    Case Else
                      Response.Write sobjString.BreakString(sobjString.Mask(FormatVariable(EditFieldValue(sobjRS2), Session(CurrentField & "Type")), Session(CurrentField & "Mask"), Session(CurrentField & "MaskCompleter")), Session(CurrentField & "ListChars"))

                  End Select
                  EndMakeLink
                  Table.ColumnAlign = ""
                  Table.EndColumn

                End If
              End Select
            Loop
          Table.ColumnWidth = "40"
          Table.BeginColumn
          If intAction = 0 Then
              Form.Submit "btmEdit", "A"
              Form.Submit "btmDelete", "E"

          ElseIf blnFirst Then
              Form.Submit "btmSave", "S"
              Form.Submit "btmCancel", "C"

          Else
            Response.Write "&nbsp;"

            End If
          Table.EndColumn
            Table.EndRow
          Form.EndForm

        End If

        If Not sparEditGrid > "" Then
          If blnFirst Then
            blnFirst = False
            sobjRS2.MoveFirst

          Else
              sobjRS2.MoveNext

          End If
        End If

        Browse.IncLine

      Loop While Not sobjRS2.EOF And Browse.IsBrowsePage

        Table.EndTable

      Browse.Status = True
      Browse.EndBrowse
      Browse.Bar

    Else
      Table.BeginTable "450"
      Table.ColumnAlign = "Center"
      Table.ColumnVAlign = "Middle"
      Table.ColumnHeight = "96"
      Table.BeginRow 2
      Table.Column "<B><A CLASS=HL HREF=""" & Initializer.ScriptURL & "?" & efQueryStrAction & "=" & efQSActionGrid & GetFilter & efQueryStrGridStr & """>Incluir " & Session(CurrentTable & "Caption") & "</A>"
      Table.EndRow
      Table.EndTable

    End If

    sobjRS2.Close

  Else
    sobjShow.MessageError LanguageLocalizer(72)

  End If
  Session("LastCurrentTable") = strPrincipalTable

End Sub

Private Sub MakeGridLink(ByVal strTable As String)
  If Session(strTable & CurrentGrid & "GridOption") = 0 Then
    Response.Write "<A HREF=" & Initializer.ScriptName & "?" & efQueryStrAction & "=" & efQSActionEditor & "&" & _
                   efQueryStrWhat & "=" & Session(CurrentTable & "WhatList") & efQueryStrEditableStr & GridTarget & " Target=_blank>"

  ElseIf Session(strTable & CurrentGrid & "GridOption") = 1 Then
    Response.Write "<A HREF=" & Initializer.ScriptName & "?" & efQueryStrAction & "=" & efQSActionEditor & "&" & _
                   efQueryStrWhat & "=" & Session(strTable & "WhatList") & efQueryStrEditableStr & GridTarget & " Target=_Top>"

  End If

End Sub

Private Function GridTarget() As String
Dim strAux As String
Dim i As Integer

  strAux = ""
  For i = 1 To Session(CurrentTable & "FieldNumber")
    If Session(CurrentTable & i & "IsKey") Then
      If strAux > "" Then
        strAux = strAux & " And "

      End If
      strAux = strAux & Session(CurrentTable & i & "Field") & "=" & FormatForSQL(Session(CurrentTable & i & "Value"), Session(CurrentTable & i & "Type"))

    End If

  Next

  If strAux > "" Then
    GridTarget = "&" & efQueryStrFind & "=" & Server.URLEncode(strAux)

  End If

End Function

'Criado (Ridai)
Public Property Let LinkTarget(ByVal strParam As String)
  sstrLinkTarget = strParam
End Property

Public Property Get LinkTarget() As String
  LinkTarget = sstrLinkTarget
End Property
Rem =========================================================================
Rem Propriedade ConnectionType
Rem -------------------------------------------------------------------------
Public Property Get ConnectionType() As Integer
  ConnectionType = sintConnectionType
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let ConnectionType(ByRef intValue As Integer)
  sintConnectionType = intValue
End Property

Public Sub EstimateConnType(Optional ByVal strConnString As String = "")
  Dim strDriverProvider As String, arrAux As Variant

  strDriverProvider = ""
  sintConnectionType = efConnDBDriverSQL

  If strConnString > "" Then sstrConnectionString = strConnString

  arrAux = Split(sstrConnectionString, ";")
  If InStr(LCase(sstrConnectionString), "driver") > 0 _
     Or InStr(LCase(sstrConnectionString), "provider") > 0 Then
    If UBound(arrAux) > 0 Then
      arrAux = Split(arrAux(0), "=")
      If UBound(arrAux) > 0 Then strDriverProvider = LCase(arrAux(1))
      arrAux = ""
    End If
  End If

  If strDriverProvider > "" Then
    Select Case strDriverProvider
    'Drivers:
    Case efCnnStrDriverSQL
      sintConnectionType = efConnDBDriverSQL

    Case efCnnStrDriverAccess
      sintConnectionType = efConnDBDriverDBase

    Case efCnnStrDriverdBase
      sintConnectionType = efConnDBDriverAccess

    Case efCnnStrDriverFoxPro
      sintConnectionType = efConnDBDriverExcel

    Case efCnnStrDriverParadox
      sintConnectionType = efConnDBDriverFoxPro

    Case efCnnStrDriverText
      sintConnectionType = efConnDBDriverText

    Case efCnnStrDriverExcel
      sintConnectionType = efConnDBDriverParadox

    Case efCnnStrDriverOracle
      sintConnectionType = efConnDBDriverOracle

    'Providers:
    Case efCnnStrProviderJet35
      sintConnectionType = efConnDBProviderJet35

    Case efCnnStrProviderJet40
      sintConnectionType = efConnDBProviderJet40

    Case efCnnStrProviderOracle
      sintConnectionType = efConnDBProviderOracle

    Case efCnnStrProviderSQL
      sintConnectionType = efConnDBProviderSQL

    Case efCnnStrProviderODBC
      sintConnectionType = efConnDBProviderODBC

    Case efCnnStrProviderDTS
      sintConnectionType = efConnDBProviderDTS

    Case efCnnStrProviderSQLDTS
      sintConnectionType = efConnDBProviderSQLDTS

    Case efCnnStrProviderSimple
      sintConnectionType = efConnDBProviderSimple

    Case efCnnStrProviderRemote
      sintConnectionType = efConnDBProviderRemote

    End Select
  End If
End Sub

Public Function DataTypeByConnection(ByVal strType As String) As String
  DataTypeByConnection = strType

  Select Case ConnectionType
  Case efConnDBDriverAccess
    DataTypeByConnection = DataType_DriverAccess(strType)
  Case efConnDBDriverDBase
    DataTypeByConnection = DataType_DriverDBase(strType)
  Case efConnDBDriverExcel
    DataTypeByConnection = DataType_DriverExcel(strType)
  Case efConnDBDriverFoxPro
    DataTypeByConnection = DataType_DriverFoxPro(strType)
  Case efConnDBDriverParadox
    DataTypeByConnection = DataType_DriverParadox(strType)
  Case efConnDBDriverSQL
    DataTypeByConnection = DataType_DriverSQL(strType)
  Case efConnDBDriverOracle
    DataTypeByConnection = DataType_DriverOracle(strType)
  Case efConnDBDriverText
    DataTypeByConnection = DataType_DriverText(strType)
  Case efConnDBProviderDTS
    DataTypeByConnection = DataType_ProviderDTS(strType)
  Case efConnDBProviderJet35
    DataTypeByConnection = DataType_ProviderJet35(strType)
  Case efConnDBProviderJet40
    DataTypeByConnection = DataType_ProviderJet40(strType)
  Case efConnDBProviderODBC
    DataTypeByConnection = DataType_ProviderODBC(strType)
  Case efConnDBProviderOracle
    DataTypeByConnection = DataType_ProviderOracle(strType)
  Case efConnDBProviderRemote
    DataTypeByConnection = DataType_ProviderRemote(strType)
  Case efConnDBProviderSimple
    DataTypeByConnection = DataType_ProviderSimple(strType)
  Case efConnDBProviderSQL
    DataTypeByConnection = DataType_ProviderSQL(strType)
  Case efConnDBProviderSQLDTS
    DataTypeByConnection = DataType_ProviderSQLDTS(strType)
  End Select
End Function

Private Function DataType_DriverSQL(ByVal strType As String) As String
  DataType_DriverSQL = strType
End Function

Private Function DataType_DriverDBase(ByVal strType As String) As String
  DataType_DriverDBase = strType
End Function

Private Function DataType_DriverAccess(ByVal strType As String) As String
  Rem Tabela comparativa SQL Server = Access
  Rem Float = Double
  Rem Real = Single
  Rem VarChar = Text
  Rem Char = Text
  Rem Text = LongText (Memo)
  Rem Int = Long
  Rem DateTime = DateTime *
  Rem Money = Currency
  Rem TinyInt = Byte
  Rem SmallInt = Short
  Rem Bit = Bit *

  Const lconAccDataTypeFloat = "double"
  Const lconAccDataTypeReal = "single"
  Const lconAccDataTypeVarChar = "text"
  Const lconAccDataTypeChar = "text"
  Const lconAccDataTypeText = "longtext" '(só que sem definição de tamanho - tirar...)
  Const lconAccDataTypeInt = "long"
  Rem Const lconAccDataTypeDateTime = "datetime"
  Const lconAccDataTypeMoney = "currency"
  Const lconAccDataTypeTinyInt = "byte"
  Const lconAccDataTypeSmallInt = "short"
  Rem Const lconAccDataTypeBit = "bit"

  DataType_DriverAccess = strType

  Select Case Trim(strType)

  Rem Case efDataTypeDateTime
  Rem   DataType_DriverAccess = lconAccDataTypeDateTime
  Rem Case efDataTypeBit
  Rem   DataType_DriverAccess = lconAccDataTypeBit

  Case efDataTypeFloat
    DataType_DriverAccess = lconAccDataTypeFloat
  Case efDataTypeReal
    DataType_DriverAccess = lconAccDataTypeReal
  Case efDataTypeVarChar
    DataType_DriverAccess = lconAccDataTypeVarChar
  Case efDataTypeChar
    DataType_DriverAccess = lconAccDataTypeChar
  Case efDataTypeText
    DataType_DriverAccess = lconAccDataTypeText
  Case efDataTypeInt
    DataType_DriverAccess = lconAccDataTypeInt
  Case efDataTypeMoney
    DataType_DriverAccess = lconAccDataTypeMoney
  Case efDataTypeTinyInt
    DataType_DriverAccess = lconAccDataTypeTinyInt
  Case efDataTypeSmallInt
    DataType_DriverAccess = lconAccDataTypeSmallInt
  End Select
End Function

Private Function DataType_DriverExcel(ByVal strType As String) As String
  DataType_DriverExcel = strType
End Function

Private Function DataType_DriverFoxPro(ByVal strType As String) As String
  DataType_DriverFoxPro = strType
End Function

Private Function DataType_DriverText(ByVal strType As String) As String
  DataType_DriverText = strType

  'Não suportado ainda:
End Function

Private Function DataType_DriverParadox(ByVal strType As String) As String
  DataType_DriverParadox = strType
End Function

Private Function DataType_ProviderJet35(ByVal strType As String) As String
  'Testar para ver se está compatível e funcional:
  DataType_ProviderJet35 = DataType_DriverAccess(strType)
End Function

Private Function DataType_ProviderJet40(ByVal strType As String) As String
  'Testar para ver se está compatível e funcional:
  DataType_ProviderJet40 = DataType_DriverAccess(strType)
End Function

Private Function DataType_ProviderOracle(ByVal strType As String) As String
   DataType_ProviderOracle = DataType_DriverOracle(strType)
End Function

Private Function DataType_DriverOracle(ByVal strType As String) As String
  'Verificar em http://technet.oracle.com/ (conta: ibehold; senha: beholder)
  Rem Tabela comparativa SQL Server = Oracle
  Rem Float = Number()
  Rem Real = Number()
  Rem VarChar = VarChar2
  Rem Char = Char *
  Rem Text = Long VarChar2 (Memo)
  Rem Int = Integer
  Rem DateTime = Date
  Rem Money = Number()
  Rem TinyInt = SmallInt 'Não existe
  Rem SmallInt = SmallInt *
  Rem Bit = SmallInt 'Não existe

  Const lconOrcDataTypeFloat = "number"
  Const lconOrcDataTypeReal = "number"
  Const lconOrcDataTypeVarChar = "varchar2"
  Rem Const lconOrcDataTypeChar = "char"
  Const lconOrcDataTypeText = "long varchar" '(só que sem definição de tamanho - tirar...)
  Const lconOrcDataTypeInt = "integer"
  Const lconOrcDataTypeDateTime = "date"
  Const lconOrcDataTypeMoney = "number"
  Const lconOrcDataTypeTinyInt = "smallint"
  Const lconOrcDataTypeSmallInt = "smallint"
  Const lconOrcDataTypeBit = "smallint"

  DataType_DriverOracle = strType

  Select Case Trim(strType)
  Case efDataTypeDateTime
    DataType_DriverOracle = lconOrcDataTypeDateTime
  Case efDataTypeBit
    DataType_DriverOracle = lconOrcDataTypeBit
  Case efDataTypeFloat
    DataType_DriverOracle = lconOrcDataTypeFloat
  Case efDataTypeReal
    DataType_DriverOracle = lconOrcDataTypeReal
  Case efDataTypeVarChar
    DataType_DriverOracle = lconOrcDataTypeVarChar
  Rem Case efDataTypeChar
  Rem   DataType_DriverOracle = lconOrcDataTypeChar
  Case efDataTypeText
    DataType_DriverOracle = lconOrcDataTypeText
  Case efDataTypeInt
    DataType_DriverOracle = lconOrcDataTypeInt
  Case efDataTypeMoney
    DataType_DriverOracle = lconOrcDataTypeMoney
  Case efDataTypeTinyInt
    DataType_DriverOracle = lconOrcDataTypeTinyInt
  Case efDataTypeSmallInt
    DataType_DriverOracle = lconOrcDataTypeSmallInt
  End Select

End Function

Private Function DataType_ProviderSQL(ByVal strType As String) As String
  'Testar problemas com este provider:
  DataType_ProviderSQL = strType
End Function

Private Function DataType_ProviderODBC(ByVal strType As String) As String
  'Desevolver tabela de equalização para SQL-92 (Transact-SQL)
  DataType_ProviderODBC = strType
End Function

Private Function DataType_ProviderDTS(ByVal strType As String) As String
  DataType_ProviderDTS = strType
End Function

Private Function DataType_ProviderSQLDTS(ByVal strType As String) As String
  DataType_ProviderSQLDTS = strType
End Function

Private Function DataType_ProviderSimple(ByVal strType As String) As String
  DataType_ProviderSimple = strType
End Function

Private Function DataType_ProviderRemote(ByVal strType As String) As String
  DataType_ProviderRemote = strType
End Function

Rem =========================================================================
Rem Recupera o próximo Identity desta Tabela do Application
Rem -------------------------------------------------------------------------
Private Function GetIdentity(ByVal strTable As String) As Long
  Dim strSQL As String
  Dim lngIdentity As Long
  Dim blnAllRight As Boolean

  blnAllRight = False
  strSQL = ""


  If NextIdentityRequestTable > "" Then

    Application.Lock

    If strTable = NextIdentityRequestTable Then blnAllRight = True

    Rem Application("htEf" & NextIdentityRequestTable & "." & NextIdentityRequestField & "." & NextIdentityTarget)
    If IsEmpty(Application("htEf" & NextIdentityRequestTable & "." & NextIdentityRequestField)) Then

      strSQL = "SELECT MAX(" & NextIdentityRequestField
      strSQL = strSQL & ") AS Codigo FROM " & NextIdentityRequestTable

      Rem sobjCMD2.CommandText = strSQL
      Rem sobjCMD2.CommandType = adCmdText

      'Alterado/Criado (Ridai)
      Set CurrConnection = sobjConn2

      If ShowSQLQuery Then ShowCurrentSQL strSQL, "GetIdentity", adOpenDynamic, adLockPessimistic

      On Error Resume Next

      sobjRS2.Open strSQL, sobjConn2, adOpenDynamic, adLockPessimistic

      If Err.number <> 0 Then
        If DebugMode Then
          ShowSQLErrors strSQL, "GetIdentity", adOpenDynamic, adLockPessimistic
        Else
          ShowDefaultError
        End If
        Err.Clear
        Exit Function

      End If
      On Error GoTo 0
      '----

      If Not IsNull(sobjRS2("Codigo")) Then
        'Session(CurrentTable & i & "OldValue") = CLng(sobjRS2("Codigo")) + 1
        'Session(CurrentTable & i & "Value") = Session(CurrentTable & i & "OldValue")

        lngIdentity = CLng(sobjRS2("Codigo")) + 1

      Else
      ' (Não há registro)
        'Session(CurrentTable & i & "Value") = 1
        'Session(CurrentTable & i & "OldValue") = 1

        lngIdentity = 1

      End If
      sobjRS2.Close

    Else
      lngIdentity = Application("htEf" & NextIdentityRequestTable & "." & NextIdentityRequestField)
      lngIdentity = lngIdentity + 1

    End If

    Application("htEf" & NextIdentityRequestTable & "." & NextIdentityRequestField) = lngIdentity
    Application.Unlock

  Else
    lngIdentity = 0
  End If

  GetIdentity = lngIdentity

End Function

Private Function IdentityCheck(Optional ByVal blnCareForEmpty As Boolean = True) As Boolean

  IdentityCheck = False

  If Left(Session(CurrentField & "DefaultValue"), 8) = constrProximo Then
     If LCase(Session(CurrentField & "Type")) = efDataTypeInt Then IdentityCheck = True

  Else
    If blnCareForEmpty Then IdentityCheck = True

  End If

End Function
Rem ====================================================================
Rem Retorna em a data que se compara o agora para se liberar um registro
Rem travado.
Rem --------------------------------------------------------------------
Private Property Get GetLockTimeOut() As Date
  Dim dtmTimeOut As Date
  Dim strLock As String

  dtmTimeOut = 0
  strLock = Application(CurrentTable & MakeTarget)

  If strLock > "" Then
    If IsDate(strLock) Then
      dtmTimeOut = DateAdd("n", CLng(LockTimeOut), CDate(strLock))
    End If
  End If

  Rem por discrepância da MS esta propriedade é Read-Only on Runtime...
  Rem App.LogMode = vbLogToNT
  Rem App.LogEvent "  |  GetLockTimeOut para tabela " & CurrentTable & " retornou: " & Format(dtmTimeOut, "dd-mm-yyyy hh:mm:ss"), vbLogEventTypeInformation

  GetLockTimeOut = dtmTimeOut

End Property

Public Property Get LockTimeOutRemaining() As Long
  Dim lngRemain As Long
  Dim dtmTimeOut As Date

  lngRemain = 0
  dtmTimeOut = GetLockTimeOut

  lngRemain = DateDiff("s", Now(), dtmTimeOut)
  sbytLockRemainType = 0

  If lngRemain > 60 Then
    lngRemain = DateDiff("n", Now(), dtmTimeOut)
    sbytLockRemainType = 1
  End If

  lngRemain = Abs(lngRemain)

  LockTimeOutRemaining = lngRemain
End Property

Public Property Get LockTimeOut() As Integer
  LockTimeOut = sintLockTimeOut
End Property

Public Property Let LockTimeOut(ByVal intTimeOut As Integer)
  sintLockTimeOut = intTimeOut
End Property

Public Property Get LockClientIP() As String
  LockClientIP = Application(CurrentTable & MakeTarget & "_ClientIP")
End Property

Public Property Get LockSessionID() As Long
  Dim strSession As String
  strSession = Application(CurrentTable & MakeTarget & "_SessionID")

  LockSessionID = 0

  If strSession > "" Then
    If IsNumeric(strSession) Then LockSessionID = CLng(strSession)
  End If
End Property

Public Function RecordLocked(Optional ByVal bytLockOption As Byte = efLockGetState) As Boolean
  Dim strAppWhere As String
  Dim dtmTimeOut As Date

  RecordLocked = False
  strAppWhere = CurrentTable & MakeTarget

  Select Case bytLockOption
  Case efLockDoLock
    Application.Lock
    Application(strAppWhere) = Now()
    Application(strAppWhere & "_SessionID") = Session.SessionID
    Application(strAppWhere & "_ClientIP") = Request.ServerVariables("REMOTE_ADDR")
    Application.Unlock
    RecordLocked = True

  Case efLockGetState
    If Application(strAppWhere) > "" Then
        If Session.SessionID <> Application(strAppWhere & "_SessionID") Then
            RecordLocked = True
        End If
    End If

  Case efLockClearLock

    dtmTimeOut = GetLockTimeOut
    If dtmTimeOut > 0 Then
      If Now >= dtmTimeOut Then
        Application.Lock
        Application(strAppWhere) = ""
        Application(strAppWhere & "_SessionID") = ""
        Application(strAppWhere & "_ClientIP") = ""
        Application.Unlock
        RecordLocked = True
      End If
    End If

  Case efLockForceClearLock
    Application.Lock
    Application(strAppWhere) = ""
    Application(strAppWhere & "_SessionID") = ""
    Application.Unlock
    RecordLocked = True

  End Select

End Function

Friend Property Let NextIdentityTarget(ByVal strTarget As String)
  Session("htNextIdentityTarget") = strTarget
End Property

Friend Property Get NextIdentityTarget() As String
  NextIdentityTarget = Session("htNextIdentityTarget")
End Property

Friend Property Let NextIdentityRequestTable(ByVal strTable As String)
  Session("htNextIdentityRequestTable") = strTable
End Property

Friend Property Get NextIdentityRequestTable() As String
  NextIdentityRequestTable = Session("htNextIdentityRequestTable")
End Property

Friend Property Let NextIdentityRequestField(ByVal strField As String)
  Session("htNextIdentityRequestField") = strField
End Property

Friend Property Get NextIdentityRequestField() As String
  NextIdentityRequestField = Session("htNextIdentityRequestField")
End Property

Private Function CheckValueAfterGet() As Boolean
  'Primeiramente para verificar se o valor do campo primary key é único...
  Dim strSQL As String

  CheckValueAfterGet = False

  If Session(CurrentField & "IsKey") And Session(CurrentField & "Value") > "" Then

    strSQL = "SELECT " & Session(CurrentField & "Field") & " AS Codigo FROM " & _
             CurrentTable & " WHERE " & Session(CurrentField & "Field") & " = " & _
             Session(CurrentField & "Value")

    Rem sobjCMD.CommandText = strSQL
    Rem sobjCMD.CommandType = adCmdText

    Set CurrConnection = sobjConn

    If ShowSQLQuery Then ShowCurrentSQL strSQL, "CheckValueAfterGet", adOpenForwardOnly, adLockReadOnly

    On Error GoTo SQLErrorHandler

    sobjRS.Open strSQL, sobjConn, adOpenForwardOnly, adLockReadOnly

    On Error GoTo 0

    If sobjRS.EOF Then CheckValueAfterGet = True

    sobjRS.Close

  Else
    CheckValueAfterGet = True

  End If

  Exit Function

SQLErrorHandler:
  If Err.number <> 0 Then
    If DebugMode Then
      ShowSQLErrors strSQL, "CheckValueAfterGet", adOpenForwardOnly, adLockReadOnly
    Else
      ShowDefaultError
    End If
    Err.Clear

  End If
End Function

'Mais um "gato" para resolver as merdas que o Eduardo fez...
Private Function GetFieldsForList() As String
  Dim strResult As String, strResult2 As String, strBuff As String, strField As String, strField2 As String
  Dim i As Integer, intCurrentField As Integer, j As Integer

  strBuff = ""
  strField = ""
  strResult = ""
  i = 0

  BeginReadOperation Session(CurrentTable & "FieldList")

  Do While IsThereOperation
    If FindField(CurrentTable, NextOperation) Then

      If strResult > "" Then strResult = strResult & ", "

      strField2 = ""
      strResult2 = ""

      If Not IsExternalField Then
        If Session(CurrentField & "EditOption") = efFldTypeLookup Then
          strResult2 = Session(CurrentField & "LookupTable") & "." & Session(CurrentField & "LookupFieldText")
          strResult = strResult & CurrentTable & "." & Session(CurrentField & "Field")
          strField2 = Session(CurrentField & "LookupFieldText")
          strField = Session(CurrentField & "Field")

        Else
          strResult = strResult & CurrentTable & "." & Session(CurrentField & "Field")
          strField = Session(CurrentField & "Field")

        End If

      Else
        strResult = strResult & OwnerTable(CurrentTable, Session("CurrentField")) & "." & Session(CurrentField & "Field")
        strField = Session(CurrentField & "Field")

      End If

      If InStr(strBuff, strField) > 0 Then
        strField = strField & "_" & i
        strResult = strResult & " AS " & strField
      End If

      If strBuff > "" Then strBuff = strBuff & ","
      strBuff = strBuff & strField

      Session(CurrentField & "Field_BackUp") = Session(CurrentField & "Field")
      Session(CurrentField & "Field") = strField

      If strField2 > "" Then

        If strBuff > "" Then strBuff = strBuff & ","
        strBuff = strBuff & strField2
        Session(CurrentField & "LookupFieldText_BackUp") = Session(CurrentField & "LookupFieldText")
        Session(CurrentField & "LookupFieldText") = strField2

        If InStr(strBuff, strField2) > 0 Then
          strField = strField2 & "_" & i
          strResult2 = strResult2 & " AS " & strField2
        End If

        If strResult2 > "" Then
          If strResult > "" Then strResult = strResult & ", "
          strResult = strResult & strResult2

        End If


      End If

      If Session(CurrentField & "ExternalLink") > "" Then
        If InStr(strBuff, Session(CurrentField & "ExtLinkField")) = 0 Then
          If strBuff > "" Then strBuff = strBuff & ","
          strBuff = strBuff & Session(CurrentField & "ExtLinkField")
          If strResult > "" Then strResult = strResult & ", "
          strResult = strResult & Session(CurrentField & "ExtLinkField")

        End If
      End If

      i = i + 1

      For j = 1 To Session(CurrentField & "ILTargetNumber")

        intCurrentField = Session("CurrentField")

        If FindField(CurrentTable, Session(CurrentField & "," & j & "ILTarget")) Then
          If strResult > "" Then strResult = strResult & ", "

          strField = Session(CurrentField & "Field")
          strResult = strResult & CurrentTable & "." & strField

          If InStr(strBuff, strField) > 0 Then
            strField = strField & "_" & i
            strResult = strResult & " AS " & strField
          End If

          If strBuff > "" Then strBuff = strBuff & ","
          strBuff = strBuff & strField

          If Not Session(CurrentField & "Field_BackUp") > "" Then
            Session(CurrentField & "Field_BackUp") = Session(CurrentField & "Field")
            Session(CurrentField & "Field") = strField
          End If

          Session("CurrentField") = intCurrentField

        End If

      Next

    End If

  Loop

  GetFieldsForList = strResult

End Function

Rem Mais outro gato:

Private Function GetJoinForList() As String
  Dim strLookUpTable As String, strJoin As String, strBuff As String
  Dim intJCounter As Integer, i As Integer
  Dim blnSameTable As Boolean

  intJCounter = 0
  blnSameTable = False

  strJoin = CurrentTable

  For i = 1 To Session(CurrentTable & "FieldNumber")
    If Session(CurrentTable & i & "EditOption") = efFldTypeLookup Then

      If intJCounter > 0 Then strJoin = "( " & strJoin & " )"

      strLookUpTable = Session(CurrentTable & i & "LookupTable")

      If LCase(Trim(CurrentTable)) = LCase(Trim(strLookUpTable)) Then
        strLookUpTable = strLookUpTable & "_" & intJCounter
      End If

      strJoin = strJoin & " LEFT JOIN " & Session(CurrentTable & i & "LookupTable")
      If LCase(Trim(Session(CurrentTable & i & "LookupTable"))) <> LCase(Trim(strLookUpTable)) Then
        strJoin = strJoin & " AS " & strLookUpTable
        blnSameTable = True
      End If

      If InStr(strBuff, strLookUpTable) > 0 And Not blnSameTable Then
        strLookUpTable = strLookUpTable & "_" & intJCounter
        strJoin = strJoin & " AS " & strLookUpTable
        Session(CurrentTable & i & "LookupTableBackUp") = Session(CurrentTable & i & "LookupTable")
        Session(CurrentTable & i & "LookupTable") = strLookUpTable
      Else
        If strBuff > "" Then strBuff = strBuff & "."
        strBuff = strBuff & strLookUpTable
      End If
      blnSameTable = False

      strJoin = strJoin & " ON " & OwnerTable(CurrentTable, i) & "." & Session(CurrentTable & i & "Field") & " = " & _
      strLookUpTable & "." & Session(CurrentTable & i & "LookupField")

      intJCounter = intJCounter + 1

    End If
  Next

  'sql = sql & " LEFT JOIN " & Session(CurrentTable & i & "LookupTable") & _
  '" ON " & Session(CurrentTable & i & "LookupTable") & "." & Session(CurrentTable & i & "LookupField") & " = " & _
  'OwnerTable(CurrentTable, i) & "." & Session(CurrentTable & i & "Field")

  GetJoinForList = strJoin
End Function

Private Sub RestoreLookupTableNames()
  Dim i As Integer

  For i = 1 To Session(CurrentTable & "FieldNumber")
    If Session(CurrentTable & i & "LookupTableBackUp") > "" Then
      Session(CurrentTable & i & "LookupTable") = Session(CurrentTable & i & "LookupTableBackUp")
      Session(CurrentTable & i & "LookupTableBackUp") = ""
    End If
  Next
End Sub

Private Sub RecoverFieldsDef()
  BeginReadOperation Session(CurrentTable & "FieldList")

  Do While IsThereOperation
    If FindField(CurrentTable, NextOperation) Then
      If Session(CurrentField & "Field_BackUp") > "" Then
        Session(CurrentField & "Field") = Session(CurrentField & "Field_BackUp")
        Session(CurrentField & "Field_BackUp") = ""
      End If
      If Session(CurrentField & "LookupFieldText_BackUp") > "" Then
        Session(CurrentField & "LookupFieldText") = Session(CurrentField & "LookupFieldText_BackUp")
        Session(CurrentField & "LookupFieldText_BackUp") = ""
      End If
    End If
  Loop

End Sub

Private Function FixListOrder(ByVal strOrder As String, ByVal strOrder2 As String) As String
  Dim arrAux As Variant, i As Integer

  arrAux = Split(strOrder, ",")

  strOrder = ""

  If IsArray(arrAux) Then
    For i = 0 To UBound(arrAux)

      If InStr(strOrder2, Replace(arrAux(i), "DESC", "")) > 0 Then
        arrAux(i) = ""
      End If

      If arrAux(i) > "" Then
        If strOrder > "" Then strOrder = strOrder & ", "
        strOrder = strOrder & arrAux(i)
      End If

    Next

    arrAux = ""

    If strOrder > "" Then strOrder2 = strOrder2 & ", "
    strOrder = strOrder2 & strOrder

  End If

  FixListOrder = strOrder

End Function

Private Sub EnumerateFields(objRS As ADODB.RecordSet)
  Dim objField As ADODB.Field

  Table.BeginTable "50%", , 3, True
  For Each objField In objRS.Fields
    Table.BeginRow
    Table.Cell objField.Name
    Table.Cell objField.Type
    Table.Cell objField.NumericScale
    Table.EndRow
  Next
  Table.EndTable

End Sub

Private Sub ShowDefaultError()
  sobjShow.MessageError Initializer.ErrorMessage
  sobjShow.MessageError "Mensagem: { " & Err.Description & " }"
End Sub

Private Property Get AllowJavaScriptRefresh() As String
  AllowJavaScriptRefresh = Session("htEfJSRefresh")
End Property

Private Property Let AllowJavaScriptRefresh(ByVal strRefresh As String)
  Session("htEfJSRefresh") = strRefresh
End Property

Private Property Get Message() As String
  Message = Session("htEfMessage")
End Property

Private Property Let Message(ByVal strMessage As String)
  Session("htEfMessage") = strMessage
End Property

Private Sub ShowLockMessage()
  Dim blnShAutoWrite As Boolean
  blnShAutoWrite = sobjShow.AutoWrite

  sobjShow.AutoWrite = False
  Message = LockMessage & sobjShow.BR

  If AllowOnLockRefresh Then
    Message = Message & _
    sobjString.FormatText(LockRefreshText, LockTimeOutRemaining, IIf(sbytLockRemainType > 0, "minuto(s)", "segundo(s)")) & sobjShow.BR
    AllowJavaScriptRefresh = "1"
  End If

  If AllowOnLockShowIP Then
   Message = Message & sobjString.FormatText(LockShowIpText, LockClientIP) & sobjShow.BR
  End If

  If DebugMode Then
    Message = Message & sobjShow.Font(, sobjShow.MessageErrorColor) & "ID da Sessão que bloqueou este registro " & LockSessionID & sobjShow.FontEnd()
  End If

  sobjShow.AutoWrite = blnShAutoWrite
End Sub

Private Sub SaveFile(ByVal lngIdentity As Long, ByVal i As Integer)
  Dim strPre As String, strFileName As String, strFileName2 As String
  Dim objFS As New Scripting.FileSystemObject

  Select Case Session(CurrentTable & i & "EditOption")
  Case efFldTypeFile
    strPre = "File"
  Case efFldTypeImage
    strPre = "Image"
  End Select

  Select Case Session(CurrentTable & i & "EditOption")
  Case efFldTypeFile, efFldTypeImage
    If Session(CurrentTable & i & "Value") > "" Then
      strFileName = Session(CurrentTable & i & "Value")
      strFileName2 = Session(CurrentTable & i & strPre & "NamePrefix") & lngIdentity & "." & GetExtension(strFileName)
      Session(CurrentTable & i & "Value") = strFileName2
      strFileName = Session(CurrentTable & i & strPre & "Root") & "/" & strFileName
      strFileName2 = Session(CurrentTable & i & strPre & "Root") & "/" & strFileName2
      strFileName = Replace(Path.DelRightSlash(Initializer.SiteRootDir) & "\" & Path.InvertSlash(strFileName), "\\", "\")
      strFileName2 = Replace(Path.DelRightSlash(Initializer.SiteRootDir) & "\" & Path.InvertSlash(strFileName2), "\\", "\")

      On Error Resume Next
      objFS.MoveFile strFileName, strFileName2

      If Err.number <> 0 Then
        ShowError efErrorFileHandling, "EditForm", "EditSave", sobjString.FormatText(LanguageLocalizer(11), Replace(Session(CurrentTable & i & "Caption"), "&", ""))
      End If
      On Error GoTo 0

    End If

  End Select

  Set objFS = Nothing
End Sub

Private Function GetExtension(ByVal strFileName As String) As String
  Dim arrAux As Variant
  Dim i As Integer

  arrAux = Split(strFileName, ".")

  GetExtension = ""

  If IsArray(arrAux) And UBound(arrAux) > -1 Then
    GetExtension = arrAux(UBound(arrAux))
  End If

End Function

Private Sub ShowUpFileAction()
  Dim Upload As HiperTools30.Upload
  Dim UpFile As HiperTools30.UploadItem
  Dim strExt As String

  Select Case Request.QueryString(efQueryStrFileOpt)
  Case efFileOptChoose
    ShowScreenUpFile

  Case efFileOptSave
    If LCase(Request.ServerVariables("REQUEST_METHOD")) = "post" Then
      Set Upload = Server.CreateObject("HiperTools30.Upload")
      Upload.Path = Replace(Initializer.SiteRootDir & Path.InvertSlash(Request.QueryString(efQueryStrFileFolder)), "\\", "\")

      Set UpFile = Upload.GetFirstFile
      strExt = GetExtension(UpFile.UserFilename)

      If UpFile.TotalBytes <= 0 Then
        sobjShow.MessageError "O arquivo escolhido está vazio."
        ShowScreenUpFile
      Else
        FindField CurrentTable, Request.QueryString(efQueryStrFileField)

        If ValidFileExts(strExt) Then
          strExt = "." & strExt
          On Error Resume Next

          UpFile.SaveAs Request.QueryString(efQueryStrFileName) & strExt

          If Err.number <> 0 Then
            sobjShow.MessageError "Ocorreu um erro ao salvar o arquivo: " & Err.Description
            ShowScreenUpFile
          Else
            Session(CurrentField & "Value") = Request.QueryString(efQueryStrFileName) & strExt
            sobjShow.HTMLCR "<SCRIPT> parent.window.close(); </SCRIPT>"

          End If
          On Error GoTo 0
        Else
          sobjShow.MessageError "Arquivo de tipo inválido ou extensão inválida."
          ShowScreenUpFile
        End If

      End If

    Else
      sobjShow.MessageError "O método de envio do formulário deve ser ""POST""!"
    End If
  End Select
End Sub

Private Sub ShowScreenUpFile()
  Dim strOnSubmit As String, arrAux As Variant
  Dim i As Long, j As Long

  FindField CurrentTable, Request.QueryString(efQueryStrFileField)
  strOnSubmit = "return " & ValidFileExtsOnClient & "('" & Session(CurrentField & "ExtensionsList") & "', document.frmFileOnEditForm.filFound);"

  Form.EncType = "MULTIPART/FORM-DATA"
  Form.BeginForm Initializer.ScriptName & "?" & efQueryStrAction & "=" & efQSActionGetFile & "&" & _
                 efQueryStrFileOpt & "=" & efFileOptSave & "&" & efQueryStrFileName & "=" & _
                 Request.QueryString(efQueryStrFileName) & "&" & _
                 efQueryStrFileFolder & "=" & Request.QueryString(efQueryStrFileFolder) & "&" & _
                 efQueryStrFileField & "=" & Request.QueryString(efQueryStrFileField) & WhatList, _
                 "frmFileOnEditForm", , , strOnSubmit

  FindField CurrentTable, Request.QueryString(efQueryStrFileField)

  Table.BeginTable "370", "Envio de arquivo - " & Replace(Session(CurrentField & "Caption"), "&", ""), 2, True
  Table.BeginRow 1

  Table.BeginCell

  Form.File "filFound", 10, 150

  Table.BeginCell

  arrAux = Split(Session(CurrentField & "ExtensionsList"), ",")

  If IsArray(arrAux) And Trim(Session(CurrentField & "ExtensionsList")) <> "*" Then
    Table.BeginTable "95%", , 4
    Table.BeginRow 1, True
    Table.CellColSpan = 4

    Table.Cell "<U>Tipos de arquivos válidos</U>:"
    Table.CellColSpan = ""

    For i = 0 To UBound(arrAux)
      If i = 0 Or i Mod 4 = 0 Then Table.BeginRow 1
      Table.Cell "*." & arrAux(i)
    Next

    j = i Mod 4
    If j <> 0 Then
      For i = 0 To j
        Table.Cell "&nbsp;"
      Next
    End If

    Table.EndTable
  Else
    sobjShow.HTMLCR "Sem limitação de tipo de arquivos."
  End If

  Table.EndRow

  Table.CellColSpan = 2
  Table.CellAlign = "CENTER"
  Table.BeginRow 1
  Table.BeginCell

  Form.Submit "btmSave", String(12, " ") & "Enviar" & String(12, " ")

  Table.EndRow
  Table.EndTable

  Form.EndForm

End Sub

Private Function ValidFileExts(ByVal strExt As String) As Boolean
  Dim blnValid As Boolean
  Dim arrAux As Variant
  Dim i As Integer

  blnValid = False

  If Trim(Session(CurrentField & "ExtensionsList")) = "*" Then
    blnValid = True
  Else
    arrAux = Split(Session(CurrentField & "ExtensionsList"), ",")

    If IsArray(arrAux) Then

      For i = 0 To UBound(arrAux)
        If Trim(LCase(strExt)) = Trim(LCase(arrAux(i))) Then blnValid = True
      Next

    End If
  End If

  ValidFileExts = blnValid
End Function

Private Function ValidFileExtsOnClient() As String
  ValidFileExtsOnClient = "htEf_VldFileExt"

  Validate.BeginJS

  sobjShow.HTMLCR "function " & ValidFileExtsOnClient & "(strExts, txtFile){  "
  sobjShow.HTMLCR "  var blnValid = false;"
  sobjShow.HTMLCR "  var arrAux, i, strExt;"
  sobjShow.HTMLCR "  var regex = / /i; //Para limpar os espaços existentes"
  sobjShow.HTMLCR "  "
  sobjShow.HTMLCR "  arrAux  = txtFile.value.split('.');"
  sobjShow.HTMLCR "  strExt  = arrAux[arrAux.length - 1];"
  sobjShow.HTMLCR "  strExts = strExts.replace(regex,'');"
  sobjShow.HTMLCR "  strExts = strExts.toLowerCase();"
  sobjShow.HTMLCR "  strExt  = strExt.toLowerCase();"
  sobjShow.HTMLCR ""
  sobjShow.HTMLCR "  if( strExts == '*'){"
  sobjShow.HTMLCR "    blnValid = true;"
  sobjShow.HTMLCR "  }"
  sobjShow.HTMLCR "  else{"
  sobjShow.HTMLCR "    arrAux = strExts.split(',');"
  sobjShow.HTMLCR "    "
  sobjShow.HTMLCR "    for(i = 0; i <= arrAux.length; i++){ if(strExt == arrAux[i]) blnValid = true;}"
  sobjShow.HTMLCR "    "
  sobjShow.HTMLCR "  }"
  sobjShow.HTMLCR "  "
  sobjShow.HTMLCR "  if(!blnValid) {"
  sobjShow.HTMLCR "    alert('Arquivo de tipo inválido ou extensão inválida.');"
  sobjShow.HTMLCR "    txtFile.focus();"
  sobjShow.HTMLCR "  }"
  sobjShow.HTMLCR "  return blnValid;"
  sobjShow.HTMLCR "}"

  Validate.EndJS

End Function
'-------------------------------------------------------------------------------------------------
' Novas Propriedades e métodos incluidos pelo Ruben Zevallos Jr.
' apartir de 17/02/2001
'-------------------------------------------------------------------------------------------------
Private Sub SetFieldMethod(ByVal strField As String, ByVal varValue As Variant, ByVal strMethod As String, ByVal strSession As String)

  If IsFieldExist(strField, strMethod) Then
    Session(CurrentField & strSession) = varValue

  End If

End Sub

Private Function IsFieldExist(ByVal strField As String, ByVal strMethod As String)
  IsFieldExist = True

  If Not FindField(CurrentTable, strField) Then
    ShowError efErrorNoField, "EditForm", strMethod, sobjString.FormatText(LanguageLocalizer(12), strField, CurrentTable)

    IsFieldExist = False

  End If

End Function

Public Sub FieldHidden(ByVal strField As String, ByVal blnResult As Boolean)
'(Define se o campo será incluido no sistema como Hidden. Parametros: nome do campo, True/False)

  SetFieldMethod strField, blnResult, "Hidden", "FieldHidden"

End Sub

Public Sub FieldAddParameter(ByVal strField As String, ByVal strParameter As String, ByVal Value)
'(Adiciona parametros extras ao INPUT do Field.
'Parametros: nome do campo, parametro, valor

  If IsFieldExist(strField, "FieldAddParameter") Then
    Session(CurrentField & "ParameterNumber") = Session(CurrentField & "ParameterNumber") + 1

    Session(CurrentField & "," & Session(CurrentField & "Parameter") & "Parameter") = strParameter
    Session(CurrentField & "," & Session(CurrentField & "Parameter") & "Value") = Value

  End If

End Sub

Public Sub FieldCaption(ByVal strField As String, ByVal varValue)
'(Define um novo Caption para o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "Caption", "FieldCaption"

End Sub

Public Sub FieldEditType(ByVal strField As String, ByVal varValue)
'(Define um novo Type Option para o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "EditOption", "FieldEditType"

End Sub

Public Sub FieldValidation(ByVal strField As String, ByVal varValue)
'(Define um novo Validation Option para o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "ValidationOption", "FieldValidation"

End Sub

Public Sub FieldPresentation(ByVal strField As String, ByVal varValue)
'(Define um novo Presentation Option para o Field:
'Parametros: nome do campo, valor)
'efPresentationEdit ou efPresentationReadOnly;

  SetFieldMethod strField, varValue, "ShowOption", "FieldPresentation"

End Sub

Public Sub FieldDefaultValue(ByVal strField As String, ByVal varValue)
'(Define um novo Default Value Option para o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "DefaultValue", "FieldDefaultValue"

End Sub

Public Sub FieldRequireMessage(ByVal strField As String, ByVal varValue)
'(Define um novo Default Value Option para o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "RequireMessage", "FieldRequireMessage"

End Sub

Public Sub FieldRequired(ByVal strField As String, ByVal varValue)
'(Define se o campo é requerido ou não :
'Parametros: nome do campo, True/False)
' efRequired ou efNotRequired;

  SetFieldMethod strField, varValue, "Required", "FieldRequired"

End Sub

Public Sub FieldDisabled(ByVal strField As String, ByVal blnResult As Boolean)
'(Define se o campo é Desabilitado ou não :
'Parametros: nome do campo,  valor (True/False))
'efDisabled ou efNotDisabled

  SetFieldMethod strField, blnResult, "Disabled", "FieldDisabled"

End Sub

Public Sub FieldSize(ByVal strField As String, ByVal varValue)
'(Redefine o Size do campo:
'Parametros: nome do campo, novo valor)

  SetFieldMethod strField, varValue, "Size", "FieldSize"

End Sub

Public Sub FieldNull(ByVal strField As String, ByVal varValue)
'(Redefine o Null do campo:
'Parametros: nome do campo, novo valor)

  SetFieldMethod strField, varValue, "Null", "FieldNull"

End Sub

Public Sub FieldType(ByVal strField As String, ByVal varValue)
'(Redefine o Type do campo:
'Parametros: nome do campo, novo valor)

  SetFieldMethod strField, varValue, "Type", "FieldType"

End Sub

Public Sub FieldTabIndex(ByVal strField As String, ByVal varValue)
'(Redefine o Type do campo:
'Parametros: nome do campo, novo valor)

  SetFieldMethod strField, varValue, "TabIndex", "FieldTabIndex"

End Sub

Public Sub FieldTextAreaAlign(ByVal strField As String, ByVal varValue)
'(Redefine o TextArea Align do campo:
'Parametros: nome do campo, novo valor)
'ABSBOTTOM | ABSMIDDLE | BASELINE | BOTTOM | LEFT | MIDDLE | RIGHT | TEXTTOP | TOP

  SetFieldMethod strField, varValue, "TextAreaAlign", "FieldTextAreaAlign"

End Sub

Public Sub FieldTextAreaWrap(ByVal strField As String, ByVal varValue)
'(Redefine o TextArea Wrap do campo:
'Parametros: nome do campo, novo valor)
'"OFF | PHYSICAL | VIRTUAL"

  SetFieldMethod strField, varValue, "TextAreaWrap", "FieldTextAreaWrap"

End Sub

Public Sub FieldAutoComplete(ByVal strField As String, ByVal blnResult As Boolean)
'(Redefine o AutoComplete do campo:
'Parametros: nome do campo, novo valor (True/False))
'"OFF | ON"

  SetFieldMethod strField, blnResult, "AutoComplete", "FieldAutoComplete"

End Sub

Public Sub FieldReadOnly(ByVal strField As String, ByVal blnResult As Boolean)
'(Redefine o AutoComplete do campo:
'Parametros: nome do campo, novo valor (True/False))
'efReadOnly or efNotReadOnly"

  SetFieldMethod strField, blnResult, "ReadOnly", "FieldReadOnly"

End Sub


Public Sub FieldStyleSheet(ByVal strField As String, ByVal varValue)
'(Define um novo Style o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "StyleSheet", "FieldStyleSheet"

End Sub

Public Sub FieldClass(ByVal strField As String, ByVal varValue)
'(Define um novo Class o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "Class", "FieldClass"

End Sub


Public Sub FieldAlign(ByVal strField As String, ByVal varValue)
'(Define um novo Align o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "Align", "FieldAlign"

End Sub

Public Sub FieldVAlign(ByVal strField As String, ByVal varValue)
'(Define um novo VAlign o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "VAlign", "FieldVAlign"

End Sub

Public Sub FieldFontFace(ByVal strField As String, ByVal varValue)
'(Define um novo FontFace o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "FontFace", "FieldFontFace"

End Sub

Public Sub FieldFontSize(ByVal strField As String, ByVal varValue)
'(Define um novo FontSize o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "FontSize", "FieldFontSize"

End Sub

Public Sub FieldFontColor(ByVal strField As String, ByVal varValue)
'(Define um novo FontColor o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "FontColor", "FieldFontColor"

End Sub

Public Sub FieldFontBGColor(ByVal strField As String, ByVal varValue)
'(Define um novo FontBGColor o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "FontBGColor", "FieldFontBGColor"

End Sub

Public Sub FieldFontBold(ByVal strField As String, ByVal blnResult As Boolean)
'(Define um novo FontBold o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, blnResult, "FontBold", "FieldFontBold"

End Sub

Public Sub FieldFontItalic(ByVal strField As String, ByVal blnResult As Boolean)
'(Define um novo FontItalic o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, blnResult, "FontItalic", "FieldFontItalic"

End Sub

Public Sub FieldFontUnderline(ByVal strField As String, ByVal blnResult As Boolean)
'(Define um novo FontUnderline o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, blnResult, "FontUnderline", "FieldFontUnderline"

End Sub


Public Sub FieldCaptionStyleSheet(ByVal strField As String, ByVal varValue)
'(Define um novo Style do Caption:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "CaptionStyleSheet", "FieldCaptionStyleSheet"

End Sub

Public Sub FieldCaptionClass(ByVal strField As String, ByVal varValue)
'(Define um novo Class do Caption:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "CaptionClass", "FieldCaptionClass"

End Sub


Public Sub FieldCaptionAlign(ByVal strField As String, ByVal varValue)
'(Define um novo Align do Caption:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "CaptionAlign", "FieldCaptionAlign"

End Sub

Public Sub FieldCaptionVAlign(ByVal strField As String, ByVal varValue)
'(Define um novo VAlign do Caption:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "CaptionVAlign", "FieldCaptionVAlign"

End Sub

Public Sub FieldCaptionFontFace(ByVal strField As String, ByVal varValue)
'(Define um novo FontFace do Caption:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "CaptionFontFace", "FieldCaptionFontFace"

End Sub

Public Sub FieldCaptionFontSize(ByVal strField As String, ByVal varValue)
'(Define um novo FontSize do Caption:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "CaptionFontSize", "FieldCaptionFontSize"

End Sub

Public Sub FieldCaptionFontColor(ByVal strField As String, ByVal varValue)
'(Define um novo FontColor o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "CaptionFontColor", "FieldCaptionFontColor"

End Sub

Public Sub FieldCaptionFontBGColor(ByVal strField As String, ByVal varValue)
'(Define um novo FontBGColor o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, varValue, "CaptionFontBGColor", "FieldCaptionFontBGColor"

End Sub

Public Sub FieldCaptionFontBold(ByVal strField As String, ByVal blnResult As Boolean)
'(Define um novo FontBold o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, blnResult, "CaptionFontBold", "FieldCaptionFontBold"

End Sub

Public Sub FieldCaptionFontItalic(ByVal strField As String, ByVal blnResult As Boolean)
'(Define um novo FontItalic o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, blnResult, "CaptionFontItalic", "FieldCaptionFontItalic"

End Sub

Public Sub FieldCaptionFontUnderline(ByVal strField As String, ByVal blnResult As Boolean)
'(Define um novo FontUnderline o Field:
'Parametros: nome do campo, valor)

  SetFieldMethod strField, blnResult, "CaptionFontUnderline", "FieldCaptionFontUnderline"

End Sub

Public Property Let Language(ByVal intLanguage As Integer)

  sintLanguage = intLanguage

  Select Case intLanguage
    Case efLanEnglish
      Buttons.AddCaption = "Include"
      Buttons.AddTitle = "Include a record"
      Buttons.AddAccessKey = "I"

      Buttons.CancelCaption = "Cancel"
      Buttons.CancelTitle = "Cancel all changes"
      Buttons.CancelAccessKey = "C"

      Buttons.DeleteCaption = "Delete"
      Buttons.DeleteTitle = "Delete the current record"
      Buttons.DeleteAccessKey = "P"

      Buttons.EditCaption = "Edit"
      Buttons.EditTitle = "Edit the current record"
      Buttons.EditAccessKey = "E"

      Buttons.FindCaption = "Locate"
      Buttons.FindTitle = "Activate the locate mode"
      Buttons.FindAccessKey = ""

      Buttons.FirstCaption = constrFirst
      Buttons.FirstTitle = "First Record"
      Buttons.FirstAccessKey = ""

      Buttons.LastCaption = constrLast
      Buttons.LastTitle = "Last record"
      Buttons.LastAccessKey = ""

      Buttons.ListCaption = "Listing"
      Buttons.ListTitle = "Activate the listing mode"
      Buttons.ListAccessKey = ""

      Buttons.NextCaption = constrNext
      Buttons.NextTitle = "Next record"
      Buttons.NextAccessKey = ""

      Buttons.PreviousCaption = constrPrevious
      Buttons.PreviousTitle = "Previous Record"
      Buttons.PreviousAccessKey = ""

      Buttons.SaveAddCaption = "Save and Add"
      Buttons.SaveAddTitle = "Save all changes and enter in include mode"
      Buttons.SaveAddAccessKey = ""

      Buttons.SaveCaption = "Save"
      Buttons.SaveTitle = "Save all changes"
      Buttons.SaveAccessKey = "S"

      Buttons.SaveCopyCaption = "Save and Copy"
      Buttons.SaveCopyTitle = "Save all changes and enter in include mode coping the current record data"
      Buttons.SaveCopyAccessKey = ""

      Buttons.CopyCaption = "Copy"
      Buttons.CopyTitle = "Include a record coping data from the current record"
      Buttons.CopyAccessKey = ""

      LockRefreshText = "New attempt for changing will be executed automatically in $d $s."
      LockMessage = "This record is being changed in this right moment by other user."
      LockShowIpText = "The host IP that blocked this record is: $s."
      ListReportsCaption = "Other listing"
      ListSetOfOrdersCaption = "Order sorting methods"

    Case Else
  End Select
End Property

Public Property Get Language() As Integer
  Language = sintLanguage

End Property

Private Function LanguageLocalizer(ByVal intMessage) As String
  Dim strMsg

Rem sobjString.FormatText(LanguageLocalizer(3), strField, CurrentTable)

  Select Case sintLanguage
    Case efLanEnglish
      Select Case intMessage
        Case 1
          strMsg = "The property ConnectionString must be filled"

        Case 2
          strMsg = "Invalid data definition: Empty name or empty type or size is zero (0)"

        Case 3
          strMsg = "The field ""$s"", of table ""$s"", must be definied as (Long) Integer at database, because it is of type Identity (Counter, Autonumeration) as defined at constant efNext "
        Case 4
          strMsg = "The table ""$s"" must be defined before to be inserted in grid"

        Case 5
          strMsg = "Especification of two or more field in a lookup is only possible at lookup of type ""pop-up""."

        Case 6
          strMsg = "Duplicated Primary key"

        Case 7
          strMsg = "The field $s is required and it's value is null. " & vbCrLf & "Contact the system developer."

        Case 8
          strMsg = "The field ""$s"", of table ""$s"", is defined as lookup, but your definition of lookup was not done or is incomplete"

        Case 9
          strMsg = "The field ""$s"" designeted for calculation at report ""$s"" was not defined."

        Case 10
          strMsg = "The table do not have a structure defined"

        Case 11
          strMsg = "Impossible to rename the file. Field: $s"

        Case 12
          strMsg = "The field ""$s"", of table ""$s"", do not exist"

        Case 13
          strMsg = "Localization of records - $s"

        Case 14
          strMsg = "Select a field of table and type the value to be localized."

        Case 15
          strMsg = "at report "

        Case 16
          strMsg = "equal of"

        Case 17
          strMsg = "Or"

        Case 18
          strMsg = "Editing"

        Case 19
          strMsg = "Listing"

        Case 20
          strMsg = "Report"

        Case 21
          strMsg = "Clear"

        Case 22
          strMsg = " or "

        Case 23
          strMsg = " and "

        Case 24
          strMsg = "equal of"

        Case 25
          strMsg = "contenting the text" 'contendo o texto"

        Case 26
          strMsg = "starting with"

        Case 27
          strMsg = "ending with"

        Case 28
          strMsg = "greater than"

        Case 29
          strMsg = "less than"

        Case 30
          strMsg = "greater or equal than"

        Case 31
          strMsg = "less or equal than"

        Case 32
          strMsg = "between"

        Case 33
          strMsg = "Atempt to paging the recordset"

        Case 34
          strMsg = " First $s found"

        Case 35
          strMsg = "There is not records for table $s for choosed options!!!"

        Case 36
          strMsg = "Include a record"

        Case 37
          strMsg = "Ordered descended by the field $s"

        Case 38
          strMsg = "Ordered ascedented by the field $s"

        Case 39
          strMsg = "Record not located"

        Case 40
          strMsg = "This record was deleted, is not possible to be changed"

        Case 41
          strMsg = "//Click at certing button to change if it exist..."

        Case 42
          strMsg = "Are you sure that you want to cancel including of record?\nAll data will be lost."

        Case 43
          strMsg = "Are your sure that you want to delete this record?"

        Case 44
          strMsg = "Filtered table"

        Case 45
          strMsg = "Including"

        Case 46
          strMsg = "Updating"

        Case 47
          strMsg = "Records: "

        Case 48
          strMsg = "Record "

        Case 49
          strMsg = " of "

        Case 50
          strMsg = "Field required"

        Case 51
          strMsg = "File not found"

        Case 52
          strMsg = "The field $s must be filled"

        Case 53
          strMsg = "The field $s was filled in the wronge way"

        Case 54
          strMsg = "This $s already exist at database."

        Case 55
          strMsg = "Initial page"

        Case 56
          strMsg = "Last option"

        Case 57
          strMsg = "HiperTools30.EditForm. $s generated the folling errors (ADODB.Errors):"

        Case 58
          strMsg = "Description: {$s}"

        Case 59
          strMsg = "Source: {$s}"

        Case 60
          strMsg = "Number: {$s}"

        Case 61
          strMsg = "The SQL query generated the errors:"

        Case 62
          strMsg = "Type of cursor: "

        Case 63
          strMsg = "Type of lock: "

        Case 64
          strMsg = " generated the follow query SQL:"

        Case 65
          strMsg = "Complete listing"

        Case 66
          strMsg = "Other listing"

        Case 67
          strMsg = "Invalid option - Listing($s) !"

        Case 68
          strMsg = "List of "

        Case 69
          strMsg = "Table created "

        Case 70
          strMsg = "Table deleted "

        Case 71
          strMsg = "Index created "

        Case 72
          strMsg = "Não foi definido um relacionamento entre a tabela folha e a tabela do grid"

        Case 73
          strMsg = "Equal"

        Case 74
          strMsg = "Field"

        Case 75
          strMsg = "Operator"

        Case 76
          strMsg = "Value"

      Case Else
        strMsg = "No message defined! (" & intMessage & ")"
      End Select

    Case Else
      Select Case intMessage
        Case 1
          strMsg = "A propriedade ConnectionString deve ser preenchida"

        Case 2
          strMsg = "Definição de dado inválida, nome vazio, tipo vazio, ou tamanho zero (0)"

        Case 3
          strMsg = "O campo ""$s"", da tabela ""$s"", deve ser definido como (Long) Integer no banco de dados, pois ele é do tipo" & _
                   " Identity (Counter, Auto-numeração) como definido com a constante efNext "
        Case 4
          strMsg = "A tabela ""$s"" deve ser definida antes de ser colocada num grid"

        Case 5
          strMsg = "Especificação de dois, ou mais campos num lookup só é possível no lookup do tipo ""pop-up""."

        Case 6
          strMsg = "Chave primária com valor duplicado"

        Case 7
          strMsg = "O campo $s é requerido e o seu valor está nulo. " & vbCrLf & "Contate o desenvolvedor do sistema."

        Case 8
          strMsg = "O campo ""$s"", da tabela ""$s"", está definido como lookup, mas sua definição de lookup não foi feita ou não está completa"

        Case 9
          strMsg = "O campo ""$s"" designado para cálculo no relatório ""$s"" não foi definido."

        Case 10
          strMsg = "A tabela não possui estrutura definida"

        Case 11
          strMsg = "Impossível renomear o arquivo. Campo: $s"

        Case 12
          strMsg = "O campo ""$s"", da tabela ""$s"", não existe"

        Case 13
          strMsg = "Localização de Registros - $s"

        Case 14
          strMsg = "Selecione um campo da tabela e digite o valor a ser localizado."

        Case 15
          strMsg = "no relatório "

        Case 16
          strMsg = "Igual a"

        Case 17
          strMsg = "Ou"

        Case 18
          strMsg = "Edição"

        Case 19
          strMsg = "Listagem"

        Case 20
          strMsg = "Relatório"

        Case 21
          strMsg = "Limpar"

        Case 22
          strMsg = " ou "

        Case 23
          strMsg = " e "

        Case 24
          strMsg = "igual a"

        Case 25
          strMsg = "contendo o texto"

        Case 26
          strMsg = "iniciando com"

        Case 27
          strMsg = "terminando com"

        Case 28
          strMsg = "maior que"

        Case 29
          strMsg = "menor que"

        Case 30
          strMsg = "maior ou igual a"

        Case 31
          strMsg = "menor ou igual a"

        Case 32
          strMsg = "entre"

        Case 33
          strMsg = "Tentativa de paginação de recordset"

        Case 34
          strMsg = " Primeiro $s encontrado"

        Case 35
          strMsg = "Não há detalhamento na Tabela $s para as opções escolhidas !!!"

        Case 36
          strMsg = "Incluir um registro"

        Case 37
          strMsg = "Ordenado descendentemente pelo campo $s"

        Case 38
          strMsg = "Ordenado ascendentemente pelo campo $s"

        Case 39
          strMsg = "Registro não localizado"

        Case 40
          strMsg = "Este registro foi excluído, não podendo ser mais alterado"

        Case 41
          strMsg = "//Clica em um certo no botão alterar se ele existir..."

        Case 42
          strMsg = "Você tem certeza que deseja cancelar a inclusão do registro ? \nOs dados digitados serão perdidos."

        Case 43
          strMsg = "Você tem certeza que deseja excluir este registro ?"

        Case 44
          strMsg = "Tabela Filtrada"

        Case 45
          strMsg = "Inclusão"

        Case 46
          strMsg = "Alteração"

        Case 47
          strMsg = "Registros: "

        Case 48
          strMsg = "Registro "

        Case 49
          strMsg = " de "

        Case 50
          strMsg = "Campo obrigatório"

        Case 51
          strMsg = "Arquivo não encontrado"

        Case 52
          strMsg = "O campo $s deve ser preenchido"

        Case 53
          strMsg = "O campo $s foi preenchido de forma inválida"

        Case 54
          strMsg = "Este(a) $s já existe no cadastro."

        Case 55
          strMsg = "Página Inicial"

        Case 56
          strMsg = "Última Opção"

        Case 57
          strMsg = "HiperTools30.EditForm. $s gerou o(s) seguite(s) Erro(s) (ADODB.Errors):"

        Case 58
          strMsg = "Descrição: {$s}"

        Case 59
          strMsg = "Fonte: {$s}"

        Case 60
          strMsg = "Número: {$s}"

        Case 61
          strMsg = "SQL Query que gerou o(s) Erro(s):"

        Case 62
          strMsg = "Tipo de cursor: "

        Case 63
          strMsg = "Tipo de lock: "

        Case 64
          strMsg = " gerou a seguite query SQL:"

        Case 65
          strMsg = "Listagem Completa"

        Case 66
          strMsg = "Outras listagens"

        Case 67
          strMsg = "Opção não existente - Listagem ($s) !"

        Case 68
          strMsg = "Lista de "

        Case 69
          strMsg = "Criada a tabela "

        Case 70
          strMsg = "Excluida a tabela "

        Case 71
          strMsg = "Criado o Indice "

        Case 72
          strMsg = "Não foi definido um relacionamento entre a tabela folha e a tabela do grid"

        Case 73
          strMsg = "Igual"

        Case 74
          strMsg = "Campo"

        Case 75
          strMsg = "Operador"

        Case 76
          strMsg = "Valor"

        Case 77
          strMsg = ""

        'ShowLockMessage

      Case Else
        strMsg = "Nenhuma mensagem definida! (" & intMessage & ")"
      End Select
  End Select

  LanguageLocalizer = strMsg

End Function

