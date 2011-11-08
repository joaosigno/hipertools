VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =========================================================================
' Form.cls
' -------------------------------------------------------------------------
' Descricao   : Include de Edicao
' Criaçao     : 11:23h 7/2/1998
' Local       : Brasilia/DF
' Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------
' Alterações:
' Responsável: Ridai Govinda Pombo <ridai@zevallos.com.br>
' Data: 21/02/2000 3:56PM
' -------------------------------------------------------------------------
' Descrição: 1a. Fase (Adicionando críticas ao objeto, com o auxílio do Validate)
'   +  Método ShowField (implementado)
'   +  Método ShowScreen

Option Explicit
  
  Rem -----------------------------------------------------------------------
  Rem Globais das Funcoes dos formularios
  Rem -----------------------------------------------------------------------
  Private Const ffValidationPrefix = "htVld_"
  
  'DebugID
  Private slngDebugID As Long
  
  Rem -----------------------------------------------------------------------
  Rem Globais das Funcoes dos formularios
  Rem -----------------------------------------------------------------------
  Private sstrEncType As String
  Private sstronClick As String
  Private sstronChange As String
  Private sstrAccessKey As String
  Private sstrAlt As String
  Private sblnAutoComplete As Boolean
  Private sstrTabIndex As String
  Private sstrTitle As String
  Private sblnReadOnly As Boolean
  Private sblnEnable As Boolean
  Private sstrClass As String
  Private sintMaxLength As Integer
  Private sstrId As String
  Private sblnJavaScriptOpened As Boolean
  Private sblnPaletteDefined As Boolean
  Private sintLastPalette As Integer
  
  'Criado (Ridai)
  Private sstrCStyleSheet As String
  
  Private sstrFormName As String
  Private sblnFormEnabled As Boolean
  Private sblnHaveForm As Boolean
  Private sblnShowOnValidate As Boolean
  Private sintFormMode As Integer
  Private sstrLastVld As String
  
  Private Validate As HiperTools30.Validate
  Private sobjString As HiperTools30.String
  Private sobjShow As HiperTools30.Show
  Private sobjFields As HiperTools30.FormFields
  '---
  
  Private Table As HiperTools30.Table
  
  Private ScriptContext As ScriptingContext

  Private sstrComments As String, _
          sstrClassTitle As String, _
          sstrVersion As String, _
          sstrCompanyName As String, _
          sstrProductName As String, _
          sstrLegalCopyright As String

  Private Request As Request, _
          Response As Response, _
          Server As Server, _
          Session As Session, _
          Application As Application

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  TimeBombX xtbForm, 15656, 12934, 3080, 14624
  
  Set ScriptContext = ScriptContext
  
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application

  Set Table = Server.CreateObject("HiperTools30.Table")
  'Criado (Ridai)
  Set Validate = Server.CreateObject("HiperTools30.Validate")
  Set sobjString = Server.CreateObject("HiperTools30.String")
  Set sobjShow = Server.CreateObject("HiperTools30.Show")
  
  sstrFormName = ""
  sintFormMode = frModeShow
  sstrLastVld = ""
  sblnHaveForm = False
  sblnShowOnValidate = True
  
  sstrCStyleSheet = ""
  '----
  Enable = True
  sblnPaletteDefined = False
  sintLastPalette = 0

End Sub

Public Sub OnEndPage()
  
  Set ScriptContext = Nothing
  
  Set Request = Nothing
  Set Response = Nothing
  Set Server = Nothing
  Set Session = Nothing
  Set Application = Nothing

  Set Table = Nothing
  'Criado (Ridai)
  Set Validate = Nothing
  Set sobjString = Nothing
  Set sobjShow = Nothing
  
  If Not sobjFields Is Nothing Then Set sobjFields = Nothing
  '---

End Sub

Private Sub Class_Initialize()
  sstrComments = App.Comments
  sstrClassTitle = App.Title
  sstrVersion = App.Major & "." & _
                App.Minor & "." & _
                App.Revision
  sstrCompanyName = App.CompanyName
  sstrProductName = App.ProductName
  
  'Criado (Ridai)
  slngDebugID = GetDebugID

End Sub

Public Property Get Comments() As String
  Comments = sstrComments
End Property

Public Property Get ClassTitle() As String
  ClassTitle = sstrClassTitle
End Property

Public Property Get Version() As String
  Version = sstrVersion
End Property

Public Property Get CompanyName() As String
  CompanyName = sstrCompanyName
End Property

Public Property Get ProductName() As String
  ProductName = sstrProductName
End Property

Public Property Get LegalCopyright() As String
  LegalCopyright = sstrLegalCopyright
End Property

' Criado pelo Ruben Zevallos Jr.
Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get AutoComplete() As Boolean
  AutoComplete = sblnAutoComplete
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let AutoComplete(ByVal blnAutoComplete As Boolean)
  sblnAutoComplete = blnAutoComplete
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get TabIndex() As String
  TabIndex = sblnAutoComplete
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let TabIndex(ByVal strTabIndex As String)
  sstrTabIndex = strTabIndex
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Alt() As String
  Alt = sstrAlt
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Alt(ByVal strAlt As String)
  sstrAlt = strAlt
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get MaxLength() As Integer
  MaxLength = sintMaxLength
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let MaxLength(ByVal intMaxLength As Integer)
  sintMaxLength = intMaxLength
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Title() As String
  Title = sstrTitle
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Title(ByVal strTitle As String)
  sstrTitle = strTitle
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Enable() As Boolean
  Enable = sblnEnable
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Enable(ByVal blnEnable As Boolean)
  sblnEnable = blnEnable
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ReadOnly() As Boolean
  ReadOnly = sblnReadOnly
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let ReadOnly(ByVal blnReadOnly As Boolean)
  sblnReadOnly = blnReadOnly
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get AccessKey() As String
  AccessKey = sstrAccessKey
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let AccessKey(ByVal strText As String)
  sstrAccessKey = strText
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get Class() As String
  Class = sstrClass
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Class(ByVal strText As String)
  sstrClass = strText
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ID() As String
  ID = sstrId
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let ID(ByVal strText As String)
  sstrId = strText
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get EncType() As String
  EncType = sstrEncType
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let EncType(ByVal strEncType As String)
  sstrEncType = strEncType
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get onClick() As String
  onClick = sstronClick
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let onClick(ByVal strText As String)
  sstronClick = strText
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get onChange() As String
  onChange = sstronChange
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let onChange(ByVal strText As String)
  sstronChange = strText
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub CheckBox(ByVal strName As String, _
                    ByVal strValue As String, _
                    Optional ByVal strAlign As String = "", _
                    Optional ByVal intState As Integer = 0)
  FormInput "CHECKBOX", strName, 0, strValue, 0, strAlign, intState
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub Radio(ByVal strName As String, _
                 ByVal strValue As String, _
                 Optional ByVal strAlign As String = "", _
                 Optional ByVal intState As Integer = 0)
  FormInput "RADIO", strName, 0, strValue, 0, strAlign, intState
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub Reset(ByVal strValue As String)
  FormInput "RESET", , , strValue
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub Submit(ByVal strName As String, _
                  ByVal strValue As String)
  FormInput "SUBMIT", strName, , strValue
End Sub

'Criado (Ridai)
Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub Button(ByVal strName As String, _
                  ByVal strValue As String)
  FormInput "BUTTON", strName, , strValue
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub Password(ByVal strName As String, _
                    ByVal intSize As Integer, _
                    Optional ByVal intMaxLenght As Integer = 0)
  FormInput "PASSWORD", strName, intSize, , intMaxLenght
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub Text(ByVal strName As String, _
                ByVal intSize As Integer, _
                Optional ByVal strValue As Variant = "", _
                Optional ByVal intMaxLength As Integer = 0)
  FormInput "TEXT", strName, intSize, strValue, intMaxLength
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub File(ByVal strName As String, _
                ByVal intSize As Integer, _
                Optional ByVal intMaxLength As Integer = 0)
  FormInput "FILE", strName, intSize, , intMaxLength
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub Hidden(ByVal strName As String, _
                  Optional ByVal strValue As String = "")
  FormInput "HIDDEN", strName, 0, strValue
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub Image(ByVal strName As String, _
                 ByVal strImage As String, _
                 Optional ByVal strAlt As String = "", _
                 Optional ByVal strValue As String = "")
  Alt = strAlt
  'Alterado (Ridai Govinda)
  FormInput "IMAGE", strName, , strValue, , , , strImage
  '-----------------------
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function FunctionFormInput(ByVal strType As String, _
                                  ByVal strName As String, _
                                  Optional ByVal intSize As Integer = 0, _
                                  Optional ByVal strValue As String = "", _
                                  Optional ByVal intMaxLength As Integer = 0, _
                                  Optional ByVal strAlign As String = "", _
                                  Optional ByVal intState As Integer = 0, _
                                  Optional ByVal strSRC As String = "") As String
  Dim strHTML

  strHTML = "<INPUT TYPE=""" & strType & """"

  If strName > "" Then
    strHTML = strHTML & " NAME=""" & strName & """"

  End If

  If intSize > 0 Then
    strHTML = strHTML & " SIZE=""" & intSize & """"

  End If

  If strValue > "" Then
     strHTML = strHTML & " VALUE=""" & strValue & """"

  End If

  If intMaxLength > 0 Then
    strHTML = strHTML & " MAXLENGTH=""" & intMaxLength & """"

  ElseIf sintMaxLength > 0 Then
    strHTML = strHTML & " MAXLENGTH=""" & sintMaxLength & """"
  
  End If
  
  If strAlign > "" Then
    strHTML = strHTML & " ALIGN=""" & strAlign & """"

  End If

  If strSRC > "" Then
    strHTML = strHTML & " SRC=""" & strSRC & """"

  End If

  If sstronClick > "" Then
    strHTML = strHTML & " onClick=""" & sstronClick & """"

  End If
  
  If sstronChange > "" Then
    strHTML = strHTML & " onChange=""" & sstronChange & """"

  End If
  
  If sstrAccessKey > "" Then
    strHTML = strHTML & " ACCESSKEY=""" & sstrAccessKey & """"

  End If

  If sstrAlt > "" Then
    strHTML = strHTML & " ALT=""" & sstrAlt & """"

  End If
  
  'Criado pelo Ruben Zevallos Jr.
  If Not sblnAutoComplete Then
    strHTML = strHTML & " AUTOCOMPLETE=""OFF"""

  End If
  
  If sstrTabIndex > "" Then
    strHTML = strHTML & " TABINDEX=""" & sstrTabIndex & """"

  End If
  '------------------
  
  If sstrTitle > "" Then
    strHTML = strHTML & " TITLE=""" & sstrTitle & """"

  End If

  If sblnReadOnly Then
    strHTML = strHTML & " READONLY"

  End If

  If Not sblnEnable Then
    strHTML = strHTML & " DISABLED"

  End If

  If sstrClass > "" Then
    strHTML = strHTML & " CLASS=""" & sstrClass & """"

  End If
  
  'Criado (Ridai Govinda)
  If sstrCStyleSheet > "" Then
    strHTML = strHTML & " STYLE=""" & sstrCStyleSheet & """"
  End If
  '---------------------

  'Alterado (Ridai Govinda)
  If intState <> 0 Then
    strHTML = strHTML & " CHECKED"

  End If
  '----------------------

  FunctionFormInput = strHTML & ">"
  
  sstrAlt = ""
  sstrTitle = ""
  sstrAccessKey = ""

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function FunctionSubmit(ByVal strName As String, Optional ByVal strValue As String = "") As String
  FunctionSubmit = FunctionFormInput("submit", strName, , strValue)
End Function

'Criado
Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function FunctionButton(ByVal strName As String, Optional ByVal strValue As String = "") As String
  FunctionButton = FunctionFormInput("BUTTON", strName, , strValue)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub FormInput(ByVal strType As String, _
                     Optional ByVal strName As String = "", _
                     Optional ByVal intSize As Integer = 0, _
                     Optional ByVal strValue As String = "", _
                     Optional ByVal intMaxLength As Integer = 0, _
                     Optional ByVal strAlign As String = "", _
                     Optional ByVal intState As Integer = 0, _
                     Optional ByVal strSRC As String = "")
  Response.Write FunctionFormInput(strType, strName, intSize, strValue, intMaxLength, strAlign, intState, strSRC)
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub BeginForm(ByVal strAction As String, _
                     Optional ByVal strName As String = "", _
                     Optional ByVal strMethod As String = "POST", _
                     Optional ByVal strTarget As String = "", _
                     Optional ByVal strOnSubmit As String = "", _
                     Optional ByVal intFormMode As Integer = frModeShow, _
                     Optional ByVal blnShowOnValidate As Boolean = True, _
                     Optional ByVal strClass As String = "", _
                     Optional ByVal strStyle As String = "", _
                     Optional ByVal blnAutoComplete As Boolean = True)

  'Alterado/Criado (Ridai)
  Dim strHTML As String
  
  Mode = intFormMode
  FormName = strName
  sblnHaveForm = True
  ShowOnValidate = blnShowOnValidate
  '------

  strHTML = "<FORM ACTION=" & Chr(34) & strAction & Chr(34)

  If strMethod > "" Then
    strHTML = strHTML & " METHOD=" & Chr(34) & strMethod & Chr(34)

  End If

  If strName > "" Then
    strHTML = strHTML & " NAME=" & Chr(34) & strName & Chr(34)

  End If

  If strTarget > "" Then
    strHTML = strHTML & " TARGET=" & Chr(34) & strTarget & Chr(34)

  End If

  'Criado Ruben Zevallos Jr.
  If strClass > "" Then
    strHTML = strHTML & " CLASS=" & Chr(34) & strClass & Chr(34)

  End If

  If strStyle > "" Then
    strHTML = strHTML & " STYLE=" & Chr(34) & strStyle & Chr(34)

  End If

  If Not blnAutoComplete Then
    strHTML = strHTML & " AUTOCOMPLETE=""OFF"""

  End If
  '---------------
  
  'Alterado (Ridai)
  Select Case Mode
  Case frModeEdit, frModeValidate
    If Mode <> frModeValidate Or (Mode = frModeValidate And ShowOnValidate) Then
      strOnSubmit = ffValidationPrefix & "OnSubmit()"
    End If
  End Select
  '---------------
  
  If strOnSubmit > "" Then
    strHTML = strHTML & " onSubmit=" & Chr(34) & strOnSubmit & Chr(34)

  End If
  
  If sstrEncType > "" Then
    strHTML = strHTML & " ENCTYPE=""" & sstrEncType & """"
    
  End If

  'Alterado (Ridai)
  If Mode <> frModeValidate Or (Mode = frModeValidate And ShowOnValidate) Then
    Response.Write strHTML & ">"
  End If
  
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub EndForm()
  'Alterado (Ridai)
  'Colocar o Script que Valida todos os campos criados...
  Select Case Mode
  Case frModeEdit
  Case frModeValidate
  End Select
  '---------------
  
  'Alterado (Ridai)
  If Mode <> frModeValidate Or (Mode = frModeValidate And ShowOnValidate) Then
    Response.Write "</FORM>"
  End If
  '---------------
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub TextArea(ByVal strName As String, _
                    Optional ByVal intRows As Integer = 0, _
                    Optional ByVal intCols As Integer = 0, _
                    Optional ByVal strContent As String = "", _
                    Optional ByVal strAlign As String = "", _
                    Optional ByVal strWrap As String = "")

  Response.Write "<TEXTAREA NAME=" & strName

  If intRows > 0 Then
    Response.Write " ROWS=""" & intRows & """"
    
  End If

  If intCols > 0 Then
    Response.Write " COLS=""" & intCols & """"
    
  End If

  'Criado pelo Ruben Zevallos Jr.
  If strAlign > "" Then
    Response.Write " ALIGN=""" & strAlign & """"
    
  End If

  If strWrap > "" Then
    Response.Write " WRAP=""" & strWrap & """"
    
  End If
  '----------------

  If sstrAccessKey > "" Then
    Response.Write " ACCESSKEY=""" & sstrAccessKey & """"

  End If
  
  If sstronChange > "" Then
    Response.Write " onChange=""" & sstronChange & """"

  End If
  
  If sstrTitle > "" Then
    Response.Write " TITLE=""" & sstrTitle & """"

  End If

  If sblnReadOnly Then
    Response.Write " READONLY"

  End If

  If Not sblnEnable Then
    Response.Write " DISABLED"

  End If

  If sstrClass > "" Then
    Response.Write " CLASS=""" & sstrClass & """"

  End If

  If Not sblnAutoComplete Then
    Response.Write " AUTOCOMPLETE=""OFF"""

  End If

  Response.Write ">"

  If strContent > "" Then
    Response.Write strContent
    
  End If
  
  Response.Write "</TEXTAREA>"

  sstrTitle = ""
  sstrAccessKey = ""

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub SelectBegin(ByVal strName As String, _
                       Optional ByVal intSize As Integer = 0, _
                       Optional ByVal blnMultiple As Boolean = False)
    BeginSelect strName, intSize, blnMultiple
                       
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub BeginSelect(ByVal strName As String, _
                       Optional ByVal intSize As Integer = 0, _
                       Optional ByVal blnMultiple As Boolean = False)
  Response.Write "<SELECT name=""" & strName & """"
  
  If intSize > 0 Then
    Response.Write " SIZE=""" & intSize & """"
    
  End If

  If blnMultiple Then
    Response.Write " MULTIPLE"
    
  End If
  
  If sstrAccessKey > "" Then
    Response.Write " ACCESSKEY=""" & sstrAccessKey & """"

  End If
  
  If sstrTitle > "" Then
    Response.Write " TITLE=""" & sstrTitle & """"

  End If

  If sblnReadOnly Then
    Response.Write " READONLY"

  End If

  If Not sblnEnable Then
    Response.Write " DISABLED"

  End If

  If sstronChange > "" Then
    Response.Write " onChange=""" & sstronChange & """"

  End If

  If sstrClass > "" Then
    Response.Write " CLASS=""" & sstrClass & """"

  End If

  sstrTitle = ""
  sstrAccessKey = ""

  Response.Write ">"

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub SelectEnd()
  EndSelect
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub EndSelect()
  Response.Write "</SELECT>"
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub SelectOption(ByVal strValue As String, _
                        ByVal strText As String, _
                        Optional ByVal blnSelect As Boolean = False)
  Response.Write "<OPTION value=""" & strValue & """"
  
  If blnSelect Then
    Response.Write " SELECTED"
    
  End If
  
  If sstrId > "" Then Response.Write " ID=""" & sstrId & """"
  
  Response.Write ">" & strText

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub SelectChooseOption(Optional ByVal strOption As Variant = "", _
                              Optional ByVal strDescription As String = "", _
                              Optional ByVal strOptionFix As Variant = "")
  SelectOption CStr(strOption), strDescription, CStr(strOption) = CStr(strOptionFix)
End Sub
  
Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub SelectMonthName(ByVal strName As String, _
                           Optional ByVal intSize As Integer = 0, _
                           Optional ByVal blnMultiple As Boolean = False, _
                           Optional ByVal intMonth As Integer = 0)
  If intMonth = 0 Then
    intMonth = Month(Now())
  
  End If

  SelectBegin strName, intSize, blnMultiple
  SelectChooseOption 1, "Janeiro", intMonth
  SelectChooseOption 2, "Fevereiro", intMonth
  SelectChooseOption 3, "Março", intMonth
  SelectChooseOption 4, "Abril", intMonth
  SelectChooseOption 5, "Maio", intMonth
  SelectChooseOption 6, "Junho", intMonth
  SelectChooseOption 7, "Julho", intMonth
  SelectChooseOption 8, "Agosto", intMonth
  SelectChooseOption 9, "Setembro", intMonth
  SelectChooseOption 10, "Outubro", intMonth
  SelectChooseOption 11, "Novembro", intMonth
  SelectChooseOption 12, "Dezembro", intMonth
  SelectEnd
                                               
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub SelectWeekDay(ByVal strName As String, _
                         Optional ByVal intSize As Integer = 0, _
                         Optional ByVal blnMultiple As Boolean = False, _
                         Optional ByVal intWeekDay As Integer = 0)
  If intWeekDay = 0 Then
    intWeekDay = Weekday(Now())
  
  End If

  SelectBegin strName, intSize, blnMultiple
  SelectChooseOption 1, "Domingo", intWeekDay
  SelectChooseOption 2, "Segunda-Feira", intWeekDay
  SelectChooseOption 3, "Terça-Feira", intWeekDay
  SelectChooseOption 4, "Quarta-Feira", intWeekDay
  SelectChooseOption 5, "Quinta-Feira", intWeekDay
  SelectChooseOption 6, "Sexta-Feira", intWeekDay
  SelectChooseOption 7, "Sábado", intWeekDay
  SelectEnd
                                               
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub SelectState(ByVal strName As String, _
                       Optional ByVal intSize As Integer = 0, _
                       Optional ByVal blnMultiple As Boolean = False, _
                       Optional ByVal strUF As String = "")

  SelectBegin strName, intSize, blnMultiple
  If Not strUF > "" Then
    SelectChooseOption "", "", ""
  
  End If
  SelectChooseOption "AC", "Acre", strUF
  SelectChooseOption "AP", "Amapá", strUF
  SelectChooseOption "AL", "Alagoas", strUF
  SelectChooseOption "AM", "Amazonas", strUF
  SelectChooseOption "BA", "Bahia", strUF
  SelectChooseOption "CE", "Ceará", strUF
  SelectChooseOption "DF", "Distrito Federal", strUF
  SelectChooseOption "ES", "Espirito Santo", strUF
  SelectChooseOption "GO", "Goiás", strUF
  SelectChooseOption "MA", "Maranhão", strUF
  SelectChooseOption "MT", "Mato Grosso", strUF
  SelectChooseOption "MS", "Mato Grosso do Sul", strUF
  SelectChooseOption "MG", "Minas Gerais", strUF
  SelectChooseOption "PA", "Pará", strUF
  SelectChooseOption "PB", "Paraiba", strUF
  SelectChooseOption "PR", "Paraná", strUF
  SelectChooseOption "PE", "Pernambuco", strUF
  SelectChooseOption "PI", "Piauí", strUF
  SelectChooseOption "RJ", "Rio de Janeiro", strUF
  SelectChooseOption "RN", "Rio Grande do Norte", strUF
  SelectChooseOption "RS", "Rio Grande do Sul", strUF
  SelectChooseOption "RO", "Rondônia", strUF
  SelectChooseOption "RR", "Roraima", strUF
  SelectChooseOption "SC", "Santa Catarina", strUF
  SelectChooseOption "SP", "São Paulo", strUF
  SelectChooseOption "SE", "Sergipe", strUF
  SelectChooseOption "TO", "Tocantins", strUF
  SelectEnd
                                               
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub SelectCountry(ByVal strName As String, _
                         Optional ByVal intSize As Integer = 0, _
                         Optional ByVal blnMultiple As Boolean = False, _
                         Optional ByVal strPais As String = "")
  SelectBegin strName, intSize, blnMultiple
  If strPais = "" Then
    SelectChooseOption "", "", ""
  
  End If
  SelectChooseOption "AF", "Afeganistão", strPais
  SelectChooseOption "AL", "Albânia", strPais
  SelectChooseOption "DZ", "Argélia", strPais
  SelectChooseOption "AS", "Samoa Americano", strPais
  SelectChooseOption "AD", "Andorra", strPais
  SelectChooseOption "AO", "Angola", strPais
  SelectChooseOption "AI", "Anguilla", strPais
  SelectChooseOption "AQ", "Antártica", strPais
  SelectChooseOption "AG", "Antígua", strPais
  SelectChooseOption "AR", "Argentina", strPais
  SelectChooseOption "AM", "Armênia", strPais
  SelectChooseOption "AW", "Aruba", strPais
  SelectChooseOption "AU", "Austrália", strPais
  SelectChooseOption "AT", "Áustria", strPais
  SelectChooseOption "AZ", "Azerbeidjã", strPais
  SelectChooseOption "BS", "Bahamas", strPais
  SelectChooseOption "BH", "Bahrain", strPais
  SelectChooseOption "BD", "Bangladesh", strPais
  SelectChooseOption "BB", "Barbados", strPais
  SelectChooseOption "BE", "Bélgica", strPais
  SelectChooseOption "BZ", "Belize", strPais
  SelectChooseOption "BJ", "Benin", strPais
  SelectChooseOption "BM", "Ilhas Bermudas", strPais
  SelectChooseOption "BT", "Butão", strPais
  SelectChooseOption "BO", "Bolívia", strPais
  SelectChooseOption "BA", "Bosnia e Herzegovina", strPais
  SelectChooseOption "BW", "Botsuana", strPais
  SelectChooseOption "BV", "Ilha de Bouvet", strPais
  SelectChooseOption "BR", "Brasil", strPais
  SelectChooseOption "IO", "Território Britânico no Oceano Índico", strPais
  SelectChooseOption "VG", "Ilhas de Virgem britânicas", strPais
  SelectChooseOption "BN", "Brunei", strPais
  SelectChooseOption "BG", "Bulgária", strPais
  SelectChooseOption "BF", "Burkina Faso", strPais
  SelectChooseOption "BU", "Birmânia", strPais
  SelectChooseOption "BI", "Burundi", strPais
  SelectChooseOption "BY", "Byelorussian SSR", strPais
  SelectChooseOption "CA", "Canadá", strPais
  SelectChooseOption "CT", "Cantão e Ilhas de Enderbury", strPais
  SelectChooseOption "CV", "Cabo Verde", strPais
  SelectChooseOption "KY", "Ilhas de caimão", strPais
  SelectChooseOption "CF", "República africana central", strPais
  SelectChooseOption "TD", "Chade", strPais
  SelectChooseOption "CL", "Chile", strPais
  SelectChooseOption "CN", "China", strPais
  SelectChooseOption "CX", "Ilha de Natal", strPais
  SelectChooseOption "CC", "Ilhas de Cocos", strPais
  SelectChooseOption "CO", "Colômbia", strPais
  SelectChooseOption "KM", "Comores", strPais
  SelectChooseOption "CG", "Congo", strPais
  SelectChooseOption "CK", "Cozinhe Ilhas", strPais
  SelectChooseOption "CR", "Costa Rica", strPais
  SelectChooseOption "HR", "Croácia (Hrvatska)", strPais
  SelectChooseOption "CU", "Cuba", strPais
  SelectChooseOption "CY", "Chipre", strPais
  SelectChooseOption "CS", "Tchecoslováquia", strPais
  SelectChooseOption "KH", "Campuchea democrático", strPais
  SelectChooseOption "YD", "Iêmen democrático", strPais
  SelectChooseOption "DK", "Dinamarca", strPais
  SelectChooseOption "DJ", "Djibouti", strPais
  SelectChooseOption "DM", "Dominica", strPais
  SelectChooseOption "DO", "República dominicana", strPais
  SelectChooseOption "NQ", "Dronning Maud Terra", strPais
  SelectChooseOption "TP", "Timor oriental", strPais
  SelectChooseOption "EC", "Equador", strPais
  SelectChooseOption "EG", "Egito", strPais
  SelectChooseOption "SV", "El Salvador", strPais
  SelectChooseOption "GQ", "Guiné equatorial", strPais
  SelectChooseOption "ER", "Eritrea", strPais
  SelectChooseOption "EE", "Estônia", strPais
  SelectChooseOption "ET", "Etiópia", strPais
  SelectChooseOption "FO", "Ilhas de Faeroe", strPais
  SelectChooseOption "FK", "Ilhas de Falkland", strPais
  SelectChooseOption "FJ", "Fiji", strPais
  SelectChooseOption "FI", "Finlândia", strPais
  SelectChooseOption "FR", "França", strPais
  SelectChooseOption "FX", "França, Metropolita,", strPais
  SelectChooseOption "GF", "Guiana Francesa", strPais
  SelectChooseOption "PF", "Polinésia Francesa", strPais
  SelectChooseOption "TF", "Territórios Meridionais franceses", strPais
  SelectChooseOption "GA", "Gabão", strPais
  SelectChooseOption "GM", "Gâmbia", strPais
  SelectChooseOption "GE", "Geórgia", strPais
  SelectChooseOption "DE", "Alemanha", strPais
  SelectChooseOption "GH", "Gana", strPais
  SelectChooseOption "GI", "Gibraltar", strPais
  SelectChooseOption "GR", "Grécia", strPais
  SelectChooseOption "GL", "Greenland", strPais
  SelectChooseOption "GD", "Granada", strPais
  SelectChooseOption "GP", "Guadalupe", strPais
  SelectChooseOption "GU", "Guam", strPais
  SelectChooseOption "GT", "Guatemala", strPais
  SelectChooseOption "GN", "Guiné", strPais
  SelectChooseOption "GW", "Guiné-bisseu", strPais
  SelectChooseOption "GY", "Guiana", strPais
  SelectChooseOption "HT", "Haiti", strPais
  SelectChooseOption "HM", "Ilhas Heard e Mc Donald", strPais
  SelectChooseOption "HN", "Honduras", strPais
  SelectChooseOption "HK", "Hong Kong", strPais
  SelectChooseOption "HU", "Hungria", strPais
  SelectChooseOption "IS", "Islândia", strPais
  SelectChooseOption "IN", "Índia", strPais
  SelectChooseOption "ID", "Indonésia", strPais
  SelectChooseOption "IR", "Irã", strPais
  SelectChooseOption "IQ", "Iraque", strPais
  SelectChooseOption "IE", "Irlanda", strPais
  SelectChooseOption "IL", "Israel", strPais
  SelectChooseOption "IT", "Itália", strPais
  SelectChooseOption "CI", "Costa do marfim", strPais
  SelectChooseOption "JM", "Jamaica", strPais
  SelectChooseOption "JP", "Japão", strPais
  SelectChooseOption "JT", "Ilha de Johnston", strPais
  SelectChooseOption "JO", "Jordan", strPais
  SelectChooseOption "KZ", "Kazakstan", strPais
  SelectChooseOption "KE", "Quênia", strPais
  SelectChooseOption "KI", "Quiribati", strPais
  SelectChooseOption "KW", "Kuwait", strPais
  SelectChooseOption "KG", "Kyrgyzstan", strPais
  SelectChooseOption "LA", "República Democrática de Lao", strPais
  SelectChooseOption "LV", "Latvia", strPais
  SelectChooseOption "LB", "Líbano", strPais
  SelectChooseOption "LS", "Lesoto", strPais
  SelectChooseOption "LR", "Libéria", strPais
  SelectChooseOption "LY", "Líbia Jamahiria árabe", strPais
  SelectChooseOption "LI", "Liechtenstein", strPais
  SelectChooseOption "LT", "Lituânia", strPais
  SelectChooseOption "LU", "Luxemburgo", strPais
  SelectChooseOption "MO", "Macau", strPais
  SelectChooseOption "MK", "Macedônia", strPais
  SelectChooseOption "MG", "Madagáscar", strPais
  SelectChooseOption "MY", "Malásia", strPais
  SelectChooseOption "MW", "Malauí", strPais
  SelectChooseOption "MV", "Maldivas", strPais
  SelectChooseOption "ML", "Mali", strPais
  SelectChooseOption "MT", "Malta", strPais
  SelectChooseOption "MQ", "Martinica", strPais
  SelectChooseOption "MR", "Mauritânia", strPais
  SelectChooseOption "MU", "Maurício", strPais
  SelectChooseOption "MX", "México", strPais
  SelectChooseOption "MI", "Ilhas Midway", strPais
  SelectChooseOption "MC", "Mônaco", strPais
  SelectChooseOption "MN", "Mongólia", strPais
  SelectChooseOption "MS", "Montserrat", strPais
  SelectChooseOption "MA", "Marrocos", strPais
  SelectChooseOption "MZ", "Moçambique", strPais
  SelectChooseOption "NA", "Namíbia", strPais
  SelectChooseOption "NP", "Nepal", strPais
  SelectChooseOption "NR", "Nauru", strPais
  SelectChooseOption "NL", "Países Baixos", strPais
  SelectChooseOption "AN", "Países Baixos Antilhas", strPais
  SelectChooseOption "NT", "Zona Neutra", strPais
  SelectChooseOption "NC", "Calidônia novo", strPais
  SelectChooseOption "NZ", "Nova Zelândia", strPais
  SelectChooseOption "NI", "Nicarágua", strPais
  SelectChooseOption "NE", "Níger", strPais
  SelectChooseOption "NG", "Nigéria", strPais
  SelectChooseOption "NU", "Niue", strPais
  SelectChooseOption "NF", "Ilha de Norfolk", strPais
  SelectChooseOption "MP", "Ilhas Mariana do norte", strPais
  SelectChooseOption "NO", "Noruega", strPais
  SelectChooseOption "OM", "Omã", strPais
  SelectChooseOption "PC", "Ilhas Pacíficas", strPais
  SelectChooseOption "PK", "Paquistão", strPais
  SelectChooseOption "PA", "Panamá", strPais
  SelectChooseOption "PG", "Papua-Nova Guiné", strPais
  SelectChooseOption "PY", "Paraguai", strPais
  SelectChooseOption "KP", "República da Coréia", strPais
  SelectChooseOption "PE", "Peru", strPais
  SelectChooseOption "PH", "Filipinas", strPais
  SelectChooseOption "PN", "Ilha de Pitcairn", strPais
  SelectChooseOption "PL", "Polônia", strPais
  SelectChooseOption "PT", "Portugal", strPais
  SelectChooseOption "TW", "Província da China Taiwan", strPais
  SelectChooseOption "PR", "Porto Rico", strPais
  SelectChooseOption "QA", "Qatar", strPais
  SelectChooseOption "KR", "República da Coréia", strPais
  SelectChooseOption "RE", "Reunião", strPais
  SelectChooseOption "RO", "Romênia", strPais
  SelectChooseOption "RU", "Federação russa", strPais
  SelectChooseOption "RW", "Ruanda", strPais
  SelectChooseOption "GS", "Ilhas S. Geórgia e S. Intercale", strPais
  SelectChooseOption "LC", "São Lúcia", strPais
  SelectChooseOption "VC", "São Vincent/Grenadines", strPais
  SelectChooseOption "WS", "Samoa", strPais
  SelectChooseOption "SM", "San Marino", strPais
  SelectChooseOption "ST", "São Tomé e Principe", strPais
  SelectChooseOption "SA", "Arábia Saudita", strPais
  SelectChooseOption "SN", "Senegal", strPais
  SelectChooseOption "SC", "Seicheles", strPais
  SelectChooseOption "SL", "Sierra Leone", strPais
  SelectChooseOption "SG", "Cingapura", strPais
  SelectChooseOption "SB", "Ilhas de Salomão", strPais
  SelectChooseOption "SO", "Somália", strPais
  SelectChooseOption "ZA", "África do Sul", strPais
  SelectChooseOption "ES", "Espanha", strPais
  SelectChooseOption "LK", "Sri Lanka", strPais
  SelectChooseOption "SH", "St. Helena", strPais
  SelectChooseOption "KN", "St. Kitts Nevis Anguilla", strPais
  SelectChooseOption "PM", "St. Pierre e Miquelon", strPais
  SelectChooseOption "SD", "Sudão", strPais
  SelectChooseOption "SR", "Suriname", strPais
  SelectChooseOption "SJ", "Ilhas Svalbard e Jan Mayen", strPais
  SelectChooseOption "SZ", "Suazilândia", strPais
  SelectChooseOption "SE", "Suécia", strPais
  SelectChooseOption "CH", "Suíça", strPais
  SelectChooseOption "SY", "Syran República árabe", strPais
  SelectChooseOption "TH", "Tailândia", strPais
  SelectChooseOption "TG", "Togo", strPais
  SelectChooseOption "TK", "Tokelau", strPais
  SelectChooseOption "TO", "Tonga", strPais
  SelectChooseOption "TT", "Trinidad e Tobago", strPais
  SelectChooseOption "TN", "Tunísia", strPais
  SelectChooseOption "TR", "Turquia", strPais
  SelectChooseOption "TM", "Turkmenistan", strPais
  SelectChooseOption "TC", "Turcos e Ilhas de Caicos", strPais
  SelectChooseOption "TV", "Tuvalu", strPais
  SelectChooseOption "UM", "EUA Ilhas Periféricas Secundárias", strPais
  SelectChooseOption "UG", "Uganda", strPais
  SelectChooseOption "UA", "SSR ucraniano", strPais
  SelectChooseOption "AE", "Emirados árabes unidos", strPais
  SelectChooseOption "GB", "Reino Unido", strPais
  SelectChooseOption "CM", "República unida de Cameroon", strPais
  SelectChooseOption "TZ", "República unida de Tanzânia", strPais
  SelectChooseOption "US", "Estados Unidos", strPais
  SelectChooseOption "PU", "Estados Unidos Ilhas do Pacífico", strPais
  SelectChooseOption "VI", "Estados Unidos Ilhas Virgem", strPais
  SelectChooseOption "HV", "Volta superior", strPais
  SelectChooseOption "UY", "Uruguai", strPais
  SelectChooseOption "UZ", "Uzbekistan", strPais
  SelectChooseOption "VU", "Vanuatu", strPais
  SelectChooseOption "VA", "Vaticano", strPais
  SelectChooseOption "VE", "Venezuela", strPais
  SelectChooseOption "VN", "Vietnã", strPais
  SelectChooseOption "WK", "Ilha Wake", strPais
  SelectChooseOption "WF", "Wallis e Ilhas de Futuma", strPais
  SelectChooseOption "EH", "Saara Ocidental", strPais
  SelectChooseOption "YE", "Iêmen", strPais
  SelectChooseOption "YU", "Iugoslávia", strPais
  SelectChooseOption "ZR", "Zaire", strPais
  SelectChooseOption "ZM", "Zâmbia", strPais
  SelectChooseOption "ZW", "Zimbabwee", strPais
  SelectEnd

End Sub

'Criado (Ridai Govinda)
Public Property Let StyleSheet(ByVal strParam As String)
  sstrCStyleSheet = strParam
End Property

Public Property Get StyleSheet() As String
  StyleSheet = sstrCStyleSheet
End Property
'---------------------

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub Palette(ByVal strName As String, Optional ByVal strColor As String = "")
  DefinePalette
  sintLastPalette = sintLastPalette + 1
  If Not strColor > "" Or Left(strColor, 1) <> "#" Or Len(strColor) <> 7 Then
    strColor = "#FFFFFF"

  End If
  If sblnEnable Then
    Hidden strName, strColor
    Response.Write "<SPAN ID='srgbimagespan" & sintLastPalette & "' style='background:" & strColor & ";width:31;height:25'><IMG ID='srgbimage" & sintLastPalette & _
    "' SRC='" & Session("htDefaultHiperToolsWorkPath") & "\Palette.gif' onclick='showImage(""" & strName & """," & sintLastPalette & ")'></SPAN>"
    
  Else
    Response.Write "<SPAN style='background:" & strColor & ";width:31;height:25'></SPAN>"

  End If

End Sub

Private Sub DefinePalette()
  
  If Not sblnPaletteDefined Then
    sblnPaletteDefined = True
    sobjShow.HTMLCR "<SPAN ID='Teste'>"

    sobjShow.HTMLCR "<SCRIPT LANGUAGE='JavaScript'>"

    sobjShow.HTMLCR "function Hex(intValue){"
    sobjShow.HTMLCR " if((intValue>-1)&&(intValue<10)){"
    sobjShow.HTMLCR "   return(intValue.toString());"
    sobjShow.HTMLCR " }"
    sobjShow.HTMLCR " else{"
    sobjShow.HTMLCR "   switch(intValue){"
    sobjShow.HTMLCR "   case 10:return('A');break;"
    sobjShow.HTMLCR "   case 11:return('B');break;"
    sobjShow.HTMLCR "   case 12:return('C');break;"
    sobjShow.HTMLCR "   case 13:return('D');break;"
    sobjShow.HTMLCR "   case 14:return('E');break;"
    sobjShow.HTMLCR "   case 15:return('F');break;"
    sobjShow.HTMLCR "   }"
    sobjShow.HTMLCR " }"
    sobjShow.HTMLCR "}"
    sobjShow.HTMLCR "  "

    sobjShow.HTMLCR "function Hex2(intValue){"
    sobjShow.HTMLCR " if(intValue > 255) intValue = 255;"
    sobjShow.HTMLCR " return( Hex( Math.floor(intValue / 16) ) + Hex(intValue % 16));"
    sobjShow.HTMLCR "}"
    sobjShow.HTMLCR "  "

  Dim intStepNum
    intStepNum = 32
    
    sobjShow.HTMLCR "var i, j , k, strHex;"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR "document.write(""<STYLE>.SP{width:" & (128 \ intStepNum) & ";height:25}.SP2{width:8;height:20;border:1 inset white}</STYLE>"");"
    sobjShow.HTMLCR "document.write(""<SPAN ID='srgbPaleta' style='display:none;background:silver;width:" & ((intStepNum + 2) * 128 / intStepNum + 100) & ";border:2 outset white;position:absolute;padding:2'>"");"
    sobjShow.HTMLCR "document.write(""<TABLE CELLPADDING=0 CELLSPACING=0 BORDER=0 WIDTH='100%'>"");"
    sobjShow.HTMLCR "document.write(""<TR><TD>"");"
    sobjShow.HTMLCR "for(i = 0; i < " & (intStepNum + 1) & "; i++){"
    sobjShow.HTMLCR "  strHex=Hex2(i * " & (256 \ intStepNum) & ");"
    sobjShow.HTMLCR "  document.write(""<SPAN CLASS='SP' ID='srgbSPRed"" + strHex + ""' onmousedown='SpanRedClick(document.all.srgbSPRed"" + strHex + "");SpanClick()' style='background:#"" + strHex + ""0000'></SPAN>"");"
    sobjShow.HTMLCR "}"
    sobjShow.HTMLCR "  "

    sobjShow.HTMLCR "document.write(""</TD>"");"
    sobjShow.HTMLCR "document.write(""<TD ROWSPAN=3>"");"
    sobjShow.HTMLCR "document.write(""<SPAN ID='srgbSep' style='width:6;height:75'></SPAN>"");"
    sobjShow.HTMLCR "document.write(""</TD>"");"
    sobjShow.HTMLCR "document.write(""<TD ROWSPAN=3>"");"
    sobjShow.HTMLCR "document.write(""<SPAN ID='srgbcColor' style='font:bold 8pt Verdana' style='Width:65'>Teste</SPAN><BR><SPAN ID='srgbColor' ondblclick='document.all.item('srgbPaleta').style.display='none'; style='height:62;width:80;background:black'></SPAN>"");"
    sobjShow.HTMLCR "document.write(""</TD></TR>"");"
    sobjShow.HTMLCR "document.write(""<TR><TD>"");"
    sobjShow.HTMLCR "for(i=0;i<" & (intStepNum + 1) & ";i++){"
    sobjShow.HTMLCR "  strHex=Hex2(i*" & (256 \ intStepNum) & ");"
    sobjShow.HTMLCR "  document.write(""<SPAN CLASS='SP' ID='srgbSPGre"" + strHex + ""' onmousedown='SpanGreenClick(document.all.srgbSPGre"" + strHex + "");SpanClick()' style='background:#00"" + strHex + ""00'></SPAN>"");"
    sobjShow.HTMLCR "}"
    sobjShow.HTMLCR "document.write(""</TD></TR>"");"
    sobjShow.HTMLCR "document.write(""<TR><TD>"");"
    sobjShow.HTMLCR "for(i=0;i<" & (intStepNum + 1) & ";i++){"
    sobjShow.HTMLCR "  strHex=Hex2(i*" & (256 \ intStepNum) & ");"
    sobjShow.HTMLCR "  document.write(""<SPAN CLASS='SP' ID='srgbSPBlu"" + strHex + ""' onmousedown='SpanBlueClick(document.all.srgbSPBlu"" + strHex + "");SpanClick()' style='background:#0000"" + strHex + ""'></SPAN>"");"
    sobjShow.HTMLCR "}"
    sobjShow.HTMLCR "document.write(""</TD></TR>"");"
    sobjShow.HTMLCR "document.write(""<TR><TD COLSPAN=3 HEIGHT=20 VALIGN='CENTER' ALIGN='CENTER'>"");"
    sobjShow.HTMLCR "document.write(""<SPAN ID='srgbSep' style='padding:4;height:30'>"");"
    sobjShow.HTMLCR "for(i=0;i<3;i++){"
    sobjShow.HTMLCR "  for(j=0;j<3;j++){"
    sobjShow.HTMLCR "    for(k=0;k<3;k++){"
    sobjShow.HTMLCR "      strHex=Hex2(i*128)+Hex2(j*128)+Hex2(k*128);"
    sobjShow.HTMLCR "      document.write(""<SPAN ID='srgbExample"" + strHex + ""' onclick='exampleClick(document.all.srgbExample"" + strHex + "")' CLASS='SP2' style='background:#"" + strHex + ""'></SPAN>"");"
    sobjShow.HTMLCR "    }"
    sobjShow.HTMLCR "  }"
    sobjShow.HTMLCR "}"
    sobjShow.HTMLCR "document.write(""</SPAN>"");"
    sobjShow.HTMLCR "document.write(""</TD></TR>"");"
    sobjShow.HTMLCR "document.write(""</TABLE>"");"
    sobjShow.HTMLCR "document.write(""</SPAN>"");"
    sobjShow.HTMLCR "</Script>"
    sobjShow.HTMLCR "  "
  
    sobjShow.HTMLCR "</SPAN>"
  
    sobjShow.HTMLCR "<Script Language='JavaScript'>"
    sobjShow.HTMLCR "var ondrag=false;"
    sobjShow.HTMLCR "var lastRed, lastGreen, lastBlue;"
    sobjShow.HTMLCR "var actualNo, actualHid;"

    sobjShow.HTMLCR "function exampleClick(strColor){"

    sobjShow.HTMLCR "SpanRedClick(document.all.item('srgbSPRed' + strColor.id.substr(11,2)));"
    sobjShow.HTMLCR "SpanGreenClick(document.all.item('srgbSPGre' + strColor.id.substr(13,2)));"
    sobjShow.HTMLCR "SpanBlueClick(document.all.item('srgbSPBlu' + strColor.id.substr(15,2)));"

    sobjShow.HTMLCR "SpanClick();"

    sobjShow.HTMLCR "}"
    sobjShow.HTMLCR "  "

    sobjShow.HTMLCR "function showImage(strName,i){"
    sobjShow.HTMLCR "document.all.srgbPaleta.style.display="""";"
    sobjShow.HTMLCR "actualNo=i;"
    sobjShow.HTMLCR "actualHid=strName;"
  
    sobjShow.HTMLCR "SpanRedClick(document.all.item('srgbSPRed' + document.all(strName).value.substr(1,2)));"
    sobjShow.HTMLCR "SpanGreenClick(document.all.item('srgbSPGre' + document.all(strName).value.substr(3,2)));"
    sobjShow.HTMLCR "SpanBlueClick(document.all.item('srgbSPBlu' + document.all(strName).value.substr(5,2)));"

    sobjShow.HTMLCR "SpanClick();"

    sobjShow.HTMLCR "document.all.item('srgbPaleta').style.left=event.x;"
    sobjShow.HTMLCR "document.all.item('srgbPaleta').style.top=event.y;"
    sobjShow.HTMLCR "}"
    sobjShow.HTMLCR "  "
   
    sobjShow.HTMLCR "function mouseUp(){"
    sobjShow.HTMLCR "ondrag=false;"
    sobjShow.HTMLCR "}"
  
    sobjShow.HTMLCR "function mouseDown(){"
    sobjShow.HTMLCR "if((event.srcElement.id.substr(0,4)!='srgb')&&(!ondrag)){"
    sobjShow.HTMLCR "document.all.item('srgbPaleta').style.display='none';"
    sobjShow.HTMLCR "}else{"
    sobjShow.HTMLCR "ondrag=false;"
    sobjShow.HTMLCR "}}"
    sobjShow.HTMLCR "  "
  
    sobjShow.HTMLCR "function mouseMove(){"
    sobjShow.HTMLCR "if((event.button==1)&&(event.srcElement.id.substr(0,6)=='srgbSP')){"
    sobjShow.HTMLCR "ondrag=true;"
    sobjShow.HTMLCR "switch (event.srcElement.id.substr(0, 9)){"
    sobjShow.HTMLCR "case 'srgbSPRed':SpanRedClick(event.srcElement);break;"
    sobjShow.HTMLCR "case 'srgbSPGre':SpanGreenClick(event.srcElement);break;"
    sobjShow.HTMLCR "case 'srgbSPBlu':SpanBlueClick(event.srcElement);break;"
    sobjShow.HTMLCR "}}}"
    sobjShow.HTMLCR "  "
  
    sobjShow.HTMLCR "function SpanRedClick(item){"
    sobjShow.HTMLCR "if(lastRed!=null){"
    sobjShow.HTMLCR "lastRed.style.border='';"
    sobjShow.HTMLCR "lastRed.style.width='" & (128 \ intStepNum) & "';"
    sobjShow.HTMLCR "}"
    sobjShow.HTMLCR "item.style.border='2 outset silver';"
    sobjShow.HTMLCR "item.style.width='" & (256 \ intStepNum) & "';"
    sobjShow.HTMLCR "lastRed=item;"
    sobjShow.HTMLCR "}"
    sobjShow.HTMLCR "  "
  
    sobjShow.HTMLCR "function SpanGreenClick(item){"
    sobjShow.HTMLCR "if(lastGreen!=null){"
    sobjShow.HTMLCR "lastGreen.style.border='';"
    sobjShow.HTMLCR "lastGreen.style.width='" & (128 \ intStepNum) & "';"
    sobjShow.HTMLCR "}"
    sobjShow.HTMLCR "item.style.border='2 outset silver';"
    sobjShow.HTMLCR "item.style.width='" & (256 \ intStepNum) & "';"
    sobjShow.HTMLCR "lastGreen=item;"
    sobjShow.HTMLCR "}"
  
    sobjShow.HTMLCR "function SpanBlueClick(item){"
    sobjShow.HTMLCR "if(lastBlue!=null){"
    sobjShow.HTMLCR "lastBlue.style.border='';"
    sobjShow.HTMLCR "lastBlue.style.width='" & (128 \ intStepNum) & "';"
    sobjShow.HTMLCR "}"
    sobjShow.HTMLCR "item.style.border='2 outset silver';"
    sobjShow.HTMLCR "item.style.width='" & (256 \ intStepNum) & "';"
    sobjShow.HTMLCR "lastBlue=item;"
    sobjShow.HTMLCR "}"
    sobjShow.HTMLCR "  "
  
    sobjShow.HTMLCR "function SpanClick(){"
    sobjShow.HTMLCR "document.all.item('srgbcColor').innerText='#'+lastRed.id.substr(9,2)+lastGreen.id.substr(9,2)+lastBlue.id.substr(9,2);"
    sobjShow.HTMLCR "document.all.item('srgbColor').style.background=document.all.item('srgbcColor').innerText;"
    sobjShow.HTMLCR "document.all(actualHid).value=document.all.item('srgbcColor').innerText;"
    sobjShow.HTMLCR "document.all.item('srgbimagespan' + actualNo).style.background=document.all.item('srgbcColor').innerText;"
    sobjShow.HTMLCR "}"
    sobjShow.HTMLCR "  "

    sobjShow.HTMLCR "document.onmousemove=mouseMove;"
    sobjShow.HTMLCR "document.onclick=mouseDown;"
    sobjShow.HTMLCR "</SCRIPT>"
    sobjShow.HTMLCR "  "
  
  End If

End Sub

Public Function AddField(ByVal strField As String, Optional ByVal strCaption As String = "", _
                      Optional ByVal intEditOption As Integer = efFldTypeText, _
                      Optional ByVal intValidationOption As Integer = 0, _
                      Optional ByVal intShowOption As Integer = 0, _
                      Optional ByVal vntDefaultValue As Variant = "", _
                      Optional ByVal strRequireMessage As String = "", _
                      Optional ByVal bytCaptionPos As Byte = 0, _
                      Optional ByVal strTitle As String = "", _
                      Optional ByVal intSize As Integer = -1, _
                      Optional ByVal intShowSize As Integer = -1, _
                      Optional ByVal strDataType As String = efDataTypeVarChar, _
                      Optional ByVal intDataSize As Integer = -1, _
                      Optional ByVal blnFocus As Boolean, _
                      Optional ByVal blnEnable As Boolean, _
                      Optional ByVal blnEmpty As Boolean, _
                      Optional ByVal strVldImage As String = "") As HiperTools30.FormField
  Dim objFField As HiperTools30.FormField
  Set objFField = Fields.Add(strField)
  Set objFField.Parent = Me
  
  objFField.Name = strField
  objFField.Caption = strCaption
  objFField.CaptionPosition = bytCaptionPos
  objFField.EditOption = intEditOption
  objFField.Validation = intValidationOption
  objFField.Title = strTitle
  objFField.DefaultValue = vntDefaultValue
  
  objFField.Size = intSize
  objFField.SizeToShow = intShowSize
  objFField.DataType = strDataType
  objFField.DataSize = intDataSize
  objFField.HasFocus = blnFocus
  objFField.Enabled = blnEnable
  objFField.EmptyField = blnEmpty
  objFField.RequireMessage = strRequireMessage
  
  objFField.ValidateImage = strVldImage
  
End Function

'Criado (Ridai)
Public Sub ShowField(ByVal strField As String, Optional ByVal strCaption As String = "", _
                      Optional ByVal intEditOption As Integer = efFldTypeText, _
                      Optional ByVal intValidationOption As Integer = 0, _
                      Optional ByVal intShowOption As Integer = 0, _
                      Optional ByVal vntDefaultValue As Variant = "", _
                      Optional ByVal strRequireMessage As String = "", _
                      Optional ByVal bytCaptionPos As Byte = 0, _
                      Optional ByVal strTitle As String = "", _
                      Optional ByVal intSize As Integer = -1, _
                      Optional ByVal intShowSize As Integer = -1, _
                      Optional ByVal strDataType As String = efDataTypeVarChar, _
                      Optional ByVal intDataSize As Integer = -1, _
                      Optional ByVal blnFocus As Boolean, _
                      Optional ByVal blnEnable As Boolean, _
                      Optional ByVal blnEmpty As Boolean, _
                      Optional ByVal blnShowVldAction As Boolean = False, _
                      Optional ByVal strVldImage As String = "")
    
    Dim objFField As HiperTools30.FormField
    
    Set objFField = Fields.Add(strField)
    Set objFField.Parent = Me
    
    objFField.Name = strField
    objFField.Caption = strCaption
    objFField.CaptionPosition = bytCaptionPos
    objFField.EditOption = intEditOption
    objFField.Validation = intValidationOption
    objFField.Title = strTitle
    objFField.DefaultValue = vntDefaultValue
    
    objFField.Size = intSize
    objFField.SizeToShow = intShowSize
    objFField.DataType = strDataType
    objFField.DataSize = intDataSize
    objFField.HasFocus = blnFocus
    objFField.Enabled = blnEnable
    objFField.EmptyField = blnEmpty
    objFField.RequireMessage = strRequireMessage
    
    objFField.ShowAll blnShowVldAction, strVldImage
    
End Sub

Public Sub ShowScreen()
  Rem Montagem do Formulário utilizando especificação de Unit (string de configuração)
  
  'If Validation <> efValOptNone Then
  '  If LastValidation = "" Then
  '    LastValidation = LastValidation & "Vld_" & Name & "()"
  '
  '  Else
  '    LastValidation = LastValidation & " && V" & Name & "()"
  '
  '  End If
  'End If
End Sub

Public Function Valid() As Boolean
  Dim objFField As HiperTools30.FormField
  
  Valid = True
  
  For Each objFField In Fields
    objFField.Read
    
    If objFField.ValidationError > "" Then
      Valid = False
    End If
  
  Next
End Function

Public Function ValidField(ByVal strField As String) As Boolean
  ValidField = Fields(strField).Valid
End Function

'Propriedades:
Public Property Get Fields() As HiperTools30.FormFields
  If sobjFields Is Nothing Then Set sobjFields = Server.CreateObject("HiperTools30.FormFields")
  
  Set Fields = sobjFields
End Property

Public Property Let ShowOnValidate(ByVal blnParam As Boolean)
  sblnShowOnValidate = blnParam
End Property

Public Property Get ShowOnValidate() As Boolean
  ShowOnValidate = sblnShowOnValidate
End Property

Public Property Let FormName(ByVal strParam As String)
  sstrFormName = strParam
End Property

Public Property Get FormName() As String
  FormName = sstrFormName
End Property

Public Property Get HaveForm() As Boolean
  HaveForm = sblnHaveForm
End Property

Public Property Let FormEnabled(ByVal blnParam As Boolean)
  sblnFormEnabled = blnParam
End Property

Public Property Get FormEnabled() As Boolean
  FormEnabled = sblnFormEnabled
End Property

Public Property Let Mode(ByVal intParam As Integer)
  sintFormMode = intParam
End Property

Public Property Get Mode() As Integer
  Mode = sintFormMode
End Property

Public Property Let LastValidation(ByVal strParam As String)
  If Not sstrLastVld > "" Then
    sstrLastVld = sstrLastVld & strParam & "()"
  
  Else
    sstrLastVld = sstrLastVld & " && " & strParam & "()"
  
  End If
End Property

Public Property Get LastValidation() As String
  LastValidation = sstrLastVld
End Property

Public Property Get DebugID() As Long
   DebugID = slngDebugID
End Property
