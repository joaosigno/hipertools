VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "FormField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =========================================================================
' FormField.cls
' -------------------------------------------------------------------------
' Descricao   : Classe de Fields p/ o Form.ShowField... etc
' Criaçao     : 02/04/2000 11:08AM
' Local       : Brasilia/DF
' Elaborado   : Ridai Govinda Pombo <ridai@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 97-2000 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------

  Option Explicit
  
  'Constantes:
  Private Const ffValidationPrefix = "htVld_"
  Private Const ffImagemCampoRequerido = "/CR.gif"
  Private Const ffImagemCampoErrado = "/CE.gif"
  
  'Propriedades com validações (mais robustas)
  Private sintFieldEditOption As Integer
  Private sintFieldValidate As Integer
  Private sintFieldSize As Integer
  Private sintFieldSizeGet As Integer
  Private sintFieldSizeShow As Integer
  Private sintFieldDataSize As Integer
  Private sstrFieldDataType As String
  Private sstrFieldRequireMessage As String
  Private sstrFieldFileRoot As String
  Private sstrFieldFileNamePref As String
  
  'Friends:
  Private sblnFormEnabled As Boolean
  Private sintFormMode As Integer
  Private sstrFormName As String
  
  'Valor do Field:
  Private svntFieldValue As Variant
  Private svntDefaultFieldValue As Variant
  
  'Propriedades mais simples:
  Public Enabled As Boolean
  Public HasFocus As Boolean
  Public EmptyField As Boolean
  Public Name As String
  Public Caption As String
  Public CaptionPosition As Byte
  Public CaptionChecked As String
  Public CaptionNotChecked As String
  Public Title As String
  Public Required As Boolean
  Public ValidationError As String
  Public ValueChecked As String
  Public ValueNotChecked As String
  Public ShowOption As Integer
  Public ID As Long
  'Especializadas:
  Public Mask As String
  Public MaskCompleter As String
  Public DoubleValidate As String
  Public DoubleValidateMsg As String
  Public BeginYear As Integer
  Public EndYear As Integer
  Public RadioColumns As Integer
  Public ImageWidth As Integer
  Public ImageHeight As Integer
  Public TextAreaHeight As Integer
  'Imagens
  Public RequiredImage As String
  Public ErrorImage As String
  Public ValidateImage As String
  'Behavior (comportamento)
  Public BeforeField As String
  Public AfterField As String
  
  'Collection de FormFieldItem (structure)
  Private scolFieldItems As Collection
  
  'Tipo (para Field do Tipo Select/Radio)
  Public Type FormFieldItem
    Value As String
    Caption As String
  End Type
  
  'HiperTools30 objects:
  'o form terá que ser uma referência para o Form que chamou o FormField
  Private objPForm As HiperTools30.Form
  
  Private Initializer As HiperTools30.Initializer
  Private Validate As HiperTools30.Validate
  Private Table As HiperTools30.Table
  Private sobjShow As HiperTools30.Show
  Private sobjString As HiperTools30.String
  '---
  
  Private ScriptContext As ScriptingContext

  Private Request As Request, _
          Response As Response, _
          Server As Server, _
          Session As Session, _
          Application As Application

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  Set ScriptContext = ScriptContext
  
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application
  
  Set Initializer = Server.CreateObject("HiperTools30.Initializer")
  Set Table = Server.CreateObject("HiperTools30.Table")
  Set Validate = Server.CreateObject("HiperTools30.Validate")
  Set sobjString = Server.CreateObject("HiperTools30.String")
  Set sobjShow = Server.CreateObject("HiperTools30.Show")
  
  'Propriedades complexas:
  sintFieldEditOption = efFldTypeText
  sintFieldValidate = efValOptNone
  sintFieldSize = -1
  sintFieldSizeShow = -1
  sintFieldSizeGet = -1
  sintFieldDataSize = -1
  sstrFieldDataType = ""
  sstrFieldRequireMessage = ""
  sstrFieldFileRoot = ""
  sstrFieldFileNamePref = ""
  
  'Propriedades Simples:
  Enabled = True
  HasFocus = True
  EmptyField = False
  Name = ""
  Caption = ""
  CaptionPosition = 0
  Title = 0
  ID = -1
  'Especializados:
  Mask = ""
  MaskCompleter = ""
  DoubleValidate = ""
  DoubleValidateMsg = ""
  BeginYear = 0
  EndYear = 0
  RadioColumns = 0
  ImageWidth = -1
  ImageHeight = -1
  Required = False
  ValidationError = ""
  ValidateImage = ""
  'Imagens
  RequiredImage = ffImagemCampoRequerido
  ErrorImage = ffImagemCampoErrado
  'Behavior (comportamento)
  BeforeField = ""
  AfterField = ""

End Sub

Public Sub OnEndPage()
  
  DestroyForm
  
  Set Initializer = Nothing
  Set Validate = Nothing
  Set Table = Nothing
  Set sobjShow = Nothing
  Set sobjString = Nothing
  
  Set ScriptContext = Nothing
  
  Set Request = Nothing
  Set Response = Nothing
  Set Server = Nothing
  Set Session = Nothing
  Set Application = Nothing

End Sub

'Propriedades:
Public Property Get FieldItems() As Collection
  If scolFieldItems Is Nothing Then
    Set scolFieldItems = New Collection
  End If
  
  Set FieldItems = scolFieldItems
End Property

Public Property Let Value(ByVal vntParam As Variant)
  'Implementar validações aqui mesmo:
  If Not IsObject(vntParam) Then svntFieldValue = vntParam
End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
  Value = svntFieldValue
End Property

Public Property Let DefaultValue(ByVal vntParam As Variant)
  'Implementar outras validaçoes aki mermo:
  If TypeName(vntParam) = TypeName(Value) And Not IsObject(vntParam) Then _
    svntDefaultFieldValue = vntParam
  
End Property

Public Property Get DefaultValue() As Variant
  DefaultValue = svntDefaultFieldValue
End Property

Public Property Let DataSize(ByVal intSize As Integer)
  Const lconDataSizeMax = 500
  
  sintFieldDataSize = IIf(intSize <= lconDataSizeMax, intSize, lconDataSizeMax)
End Property

Public Property Get DataSize() As Integer
  DataSize = sintFieldDataSize
End Property

Public Property Let EditOption(ByVal intParam As Integer)
  
  Select Case intParam
  Case efFldTypeText, efFldTypeHTTP, efFldTypeEMail, efFldTypeAtualization, _
       efFldTypeColor, efFldTypeFile, efFldTypeLookup, efFldTypeCheck, _
       efFldTypeTextArea, efFldTypeRadio, efFldTypeSelect, efFldTypePassword, _
       efFldTypeSeparateDate, efFldTypeImage, efFldTypeUF
    sintFieldEditOption = intParam
  
  End Select
End Property

Public Property Get EditOption() As Integer
  EditOption = sintFieldEditOption
End Property

Public Property Let Validation(ByVal intParam As Integer)
  Select Case intParam
  Case efValOptNone, efValOptCGC, efValOptCPF, efValOptDate, _
       efValOptSepDate, efValOptDateMToday, efValOptSepDateMToday, _
       efValOptTime, efValOptEmail, efValOptCompareDates, efValOptCEP
    sintFieldValidate = intParam
    
  End Select
End Property

Public Property Get Validation() As Integer
  Validation = sintFieldValidate
End Property

Public Property Get DataSizeToShow() As Integer
' (propriedade que retorna o tamanho de um campo para visualização)
  
' (Por tipo)
  Select Case DataType
    Case efDataTypeInt, efDataTypeTinyInt, efDataTypeSmallInt
      DataSizeToShow = Len(sobjString.ToString(256 ^ Size)) - 1

    Case efDataTypeFloat
      DataSizeToShow = Len(sobjString.ToString(256 ^ Size)) + 1

    Case efDataTypeMoney
      DataSizeToShow = Len(sobjString.ToString(256 ^ Size)) + 5

    Case efDataTypeDateTime
      Select Case Validation
        Case efValOptTime
          DataSizeToShow = 5

        Case Else
          DataSizeToShow = 10

      End Select

    Case Else
      DataSizeToShow = Size

  End Select

End Property

Public Property Let RequireMessage(ByVal strParam As String)
  sstrFieldRequireMessage = strParam
End Property

Public Property Get RequireMessage() As String
  RequireMessage = sstrFieldRequireMessage
End Property

Public Property Let Size(ByVal intParam As Integer)
  sintFieldSize = intParam
End Property

Public Property Get Size() As Integer
  Size = sintFieldSize
End Property

Public Property Get SizeToGet() As Integer
' (Função que retorna o nº máximo caracteres de um campo)
  If sintFieldSizeGet > -1 Then
    SizeToGet = sintFieldSizeGet

  Else
    SizeToGet = DataSizeToShow

  End If

End Property

Public Property Let SizeToGet(ByVal intParam As Integer)
  sintFieldSizeGet = intParam
End Property

Public Property Get SizeToShow() As Integer
' (Função que retorna o tamanho de um campo para visualização)
  Const lconMaxShowSize = 80
  Dim intSize As Integer
  
  If sintFieldSizeShow > -1 Then
    SizeToShow = sintFieldSizeShow

  Else
    intSize = DataSizeToShow
    
    If intSize <= lconMaxShowSize Then
      SizeToShow = intSize

    Else
      SizeToShow = lconMaxShowSize

    End If

  End If

End Property

Public Property Let SizeToShow(ByVal intParam As Integer)
  sintFieldSizeShow = intParam
End Property

Public Property Get DataType() As String
  DataType = sstrFieldDataType
End Property

Public Property Let DataType(ByVal strParam As String)
  Select Case strParam
  Case efDataTypeFloat, efDataTypeReal, efDataTypeVarChar, efDataTypeChar, _
       efDataTypeText, efDataTypeInt, efDataTypeDateTime, efDataTypeMoney, efDataTypeTinyInt, _
       efDataTypeSmallInt, efDataTypeBit
    sstrFieldDataType = strParam
  End Select
End Property

Public Property Get FileNamePrefix() As String
  FileNamePrefix = sstrFieldFileNamePref
End Property

Public Property Let FileNamePrefix(ByVal strParam As String)
  sstrFieldFileNamePref = strParam
End Property

Public Property Get FileRoot() As String
  FileRoot = sstrFieldFileRoot
End Property

Public Property Let FileRoot(ByVal strParam As String)
  sstrFieldFileRoot = strParam
End Property

Friend Property Set Parent(objParam As HiperTools30.Form)
  Set objPForm = objParam
End Property

Public Property Get Parent() As HiperTools30.Form
  If objPForm Is Nothing Then Set objPForm = Server.CreateObject("HiperTools30.Form")
  Set Parent = objPForm
End Property

'propriedades Friend para evitar referência cíclica.
'São exatamente as propriedades do Form:

' Friend Property Let FormName(ByVal strParam As String)
'   sstrFormName = strParam
' End Property
'
' Friend Property Get FormName() As String
'   FormName = sstrFormName
' End Property
'
' Friend Property Let FormEnabled(ByVal blnParam As Boolean)
'   sblnFormEnabled = blnParam
' End Property
'
' Friend Property Get FormEnabled() As Boolean
'   FormEnabled = sblnFormEnabled
' End Property
'
' Friend Property Let Mode(ByVal intParam As Integer)
'   sintFormMode = intParam
' End Property
'
' Friend Property Get Mode() As Integer
'   Mode = sintFormMode
' End Property
'---

'Métodos:

Public Sub FieldMask(ByVal strMask As String, Optional ByVal strCompleter As String = "")
' (Define a máscara a ser colocada em um campo pela função Mask do objeto String, parametros:nome
' do campo, máscara, completador da máscara)
  Mask = strMask
  MaskCompleter = strCompleter

End Sub

Public Sub AddValue(ByVal vntValue As Variant, ByVal strCaption As String)
  Dim typFormField As FormFieldItem
  
  If EditOption <> efFldTypeCheck Then
    typFormField.Caption = strCaption
    typFormField.Value = vntValue
    
    scolFieldItems.Add typFormField, strCaption
  Else
    If ValueChecked > "" Then
      ValueNotChecked = vntValue
      CaptionNotChecked = strCaption
    Else
      ValueChecked = vntValue
      CaptionChecked = strCaption
    End If
  End If
End Sub

Public Sub ShowAll(Optional ByVal blnShowVldAction As Boolean = False, _
                   Optional ByVal strVldImage As String = "", _
                   Optional ByVal bytShowVldType As Byte = 0)
  
  sobjShow.HTML BeforeField
  
  If objPForm.Mode <> frModeValidate Or (objPForm.Mode = frModeValidate And objPForm.ShowOnValidate) Then
    
    If strVldImage > "" Then
      ValidateImage = strVldImage
      bytShowVldType = ffVldActionImage
    End If
    
    If objPForm.Mode = frModeShow Then
      ShowAsText
    ElseIf objPForm.Mode = frModeEdit Then
      Show
      ShowValidation
    End If
    
    If blnShowVldAction Then
      ShowValidateAction bytShowVldType
    End If
  
  End If

  sobjShow.HTML AfterField
  
End Sub

Public Sub Show(Optional blnFocus As Boolean, Optional blnEnabled As Boolean, Optional blnEmpty As Boolean)
  InstanceForm
  
  If Not IsMissing(blnFocus) Then HasFocus = blnFocus
  If Not IsMissing(blnEnabled) Then Enabled = blnEnabled
  If Not IsMissing(blnEmpty) Then EmptyField = blnEmpty
  
  objPForm.Enable = Enabled
  If objPForm.FormEnabled Then objPForm.Enable = objPForm.FormEnabled
  
  If Caption > "" And EditOption <> efFldTypeFile _
                  And EditOption <> efFldTypeCheck Then
    ShowCaption
  End If
   
  If ShowOption > -1 Then
    objPForm.Title = Title
    If EmptyField Then Value = ""
  
  Else
    If Not Value > "" Then
      Value = ""
  
    Else
      If EditOption = efFldTypeText Then Value = FormatVariable(Value, DataType)
    
    End If
  End If
  
  Select Case EditOption
  Case efFldTypeText, efFldTypeHTTP, efFldTypeEMail, efFldTypeFTP, efFldTypeGopher
    ShowTextField
  
  Case efFldTypeAtualization
    ShowImages
    ShowBreak
    Response.Write Value
    
  Case efFldTypeColor
    objPForm.Palette Name, sobjString.ToString(Value)
  
  Case efFldTypeFile, efFldTypeImage
    ShowFileField
  
  Case efFldTypePassword
    ShowImages
    ShowBreak
    objPForm.Password Name, SizeToShow
  
  Case efFldTypeUF, efFldTypeCountry
    ShowEspecialSelectField
  
  Case efFldTypeTextArea
    ShowImages
    ShowBreak
    objPForm.TextArea Name, CInt(TextAreaHeight), CInt(SizeToShow), sobjString.ToString(Value)
  
  Case efFldTypeCheck
    objPForm.CheckBox Name, ValueChecked, _
    "LEFT", BitBln(sobjString.ToString(Value) = sobjString.ToString(ValueNotChecked))
    ShowCaption
      
  Case efFldTypeRadio
    ShowRadioField
  
  Case efFldTypeSelect
    ShowSelectField
  
  Case efFldTypeSeparateDate
    ShowSepDateField
  Case efFldTypeLookup
    Err.Raise 300, "HiperTools30.FormField.Show", "Não implementado ainda"
  
  End Select
  
  If HasFocus Then
    Select Case EditOption
      Case efFldTypeImage
  
      Case efFldTypeRadio
        sobjShow.CRHTMLCR "<SCRIPT LANGUAGE=""JavaScript"">"
        sobjShow.HTMLCR "document.forms." & objPForm.FormName & "." & Name & "(0).focus();"
        sobjShow.HTMLCR "</SCRIPT>"
  
      Case Else
        sobjShow.CRHTMLCR "<SCRIPT LANGUAGE=""JavaScript"">"
        sobjShow.HTMLCR "document.forms." & objPForm.FormName & "." & Name & ".focus();"
        sobjShow.HTMLCR "</SCRIPT>"
  
    End Select
  
  End If
  
  Rem DestroyForm

End Sub

Public Sub ShowValidation()
  If objPForm.FormName > "" And objPForm.HaveForm Then
  
    Select Case Validation
    Case efValOptCGC
      Validate.FormJavaValidCGC Name, Parent.FormName, ffValidationPrefix & Name
  
    Case efValOptCPF
      Validate.FormJavaValidCPF Name, Parent.FormName, ffValidationPrefix & Name
  
    Case efValOptCEP
      Validate.FormJavaValidCEP Name, Parent.FormName, ffValidationPrefix & Name
  
    Case efValOptDate
      Validate.FormJavaValidDate Name, Parent.FormName, ffValidationPrefix & Name, False, ""
  
    Case efValOptCompareDates
      Validate.FormJavaCompareDate Name, DoubleValidate, Parent.FormName, ffValidationPrefix & Name, DoubleValidateMsg
  
    Case efValOptTime
      Rem FormJavaValidTime name, ffValidationPrefix & name, False, ""
  
    Case efValOptSepDate
      Validate.FormJavaValidSepDate Name & "DayValue", Name & "MonthValue", Name & "YearValue", Parent.FormName, ffValidationPrefix & Name, False, ""
  
    Case efValOptDateMToday
      Validate.FormJavaValidDate Name, Parent.FormName, ffValidationPrefix & Name, True, GetCaption(, , , False) & " deve ser maior que hoje"
  
    Case efValOptSepDateMToday
      Validate.FormJavaValidSepDate Name & "DayValue", Name & "MonthValue", Name & "YearValue", Parent.FormName, ffValidationPrefix & Name, True, GetCaption(, , , False) & " deve ser maior que hoje"
  
    Case efValOptEmail
      Validate.FormJavaValidEMail Name, Parent.FormName, ffValidationPrefix & Name
  
    End Select
    
    If Validation <> efValOptNone Then
      objPForm.LastValidation = ffValidationPrefix & Name
    End If
    
    Validate.JavaEmpty
    Validate.JavaBegin ffValidationPrefix & "All" & Parent.FormName, Parent.FormName
  
    If EditOption <> efFldTypePassword Or Not sobjString.ToString(Value) > "" And _
       EditOption <> efFldTypeCheck Then
      If RequireMessage > "" Then
        Validate.JavaRequired Name, RequireMessage
  
      ElseIf Required And EditOption <> efFldTypeRadio Then
        Validate.JavaRequired Name, "O campo " & GetCaption(, , , False) & " deve ser preenchido"
  
      End If
    End If
    Response.Write "} "
    Validate.JSEnd

  End If
  
End Sub

Public Sub ShowValidateAction(Optional ByVal bytShowVldType As Byte = 0)
  If objPForm.FormName > "" And objPForm.HaveForm Then
    Select Case bytShowVldType
    Case ffVldActionButton 'Butão
      objPForm.onClick = "return " & ffValidationPrefix & Name & "();"
      objPForm.Button "btn" & ffValidationPrefix & Name, "Validar"
    
    Case ffVldActionLink 'Link
      Response.Write "<A HREF=""#"" onClick=""javascript:" & ffValidationPrefix & Name & "();""> Validar </A>"
    
    Case ffVldActionImage 'Imagem
      If ValidateImage > "" Then
        Response.Write "<A HREF=""#"" onClick=""javascript:" & ffValidationPrefix & Name & "();"">"
        sobjShow.Image ValidateImage
        Response.Write "</A>"
      End If
    
    End Select
  End If
End Sub

Public Sub ShowAsText()
    Dim typFieldItem As Variant
    
    If EditOption <> efFldTypeImage Then ShowCaption , True, , 1
    
    Select Case EditOption
    ' (Campos tipos Texto, UF, Area de texto e Data de Atualização)
    Case efFldTypeText, efFldTypeUF, efFldTypeTextArea, efFldTypeAtualization
      
      'Implementar:
      
      'If (Session(CurrentField & "ExternalLink") > "") Then
      '  BeginMakeLink True
      '
      'End If

      Response.Write sobjString.Mask(FormatVariable(Value, DataType), Mask, MaskCompleter)
      
      'EndMakeLink

    ' (Campo tipo Cor)
    Case efFldTypeColor
      Response.Write "<SPAN style=""background:" & Value & ";width:31;height:25""></SPAN>"

    ' (Campo tipo Endereço HTTP)
    Case efFldTypeHTTP
      Response.Write "<A HREF=""http://" & Value & """>" & Value & "</A>"

    ' (Campo tipo E-Mail)
    Case efFldTypeEMail
      Response.Write "<A HREF=""mailto:" & Value & """>" & Value & "</A>"
    
    ' (Campo tipo Endereço FTP)
    Case efFldTypeFTP
      Response.Write "<A HREF=""ftp://" & Value & """>" & Value & "</A>"
    
    ' (Campo tipo Endereço Gopher)
    Case efFldTypeGopher
      Response.Write "<A HREF=""gopher://" & Value & """>" & Value & "</A>"

    ' (Campo tipo Arquivo de Imagem)
    Case efFldTypeImage
      ShowCaption , True, , 0
      
      If Value > "" Then
        sobjShow.Image "/" & FileRoot & "/" & Value, IIf(ID > -1, ID, ""), _
            "/" & FileRoot & "/" & Value, "'/" & FileRoot & "/" & Value & "' TARGET=__blank", _
            "", IIf(ImageWidth > -1, ImageWidth, ""), IIf(ImageHeight > -1, ImageHeight, "")

      End If

    ' (Campo tipo Arquivo)
    Case efFldTypeFile
      If Value > "" Then
        sobjShow.HTMLCR "<A HREF=""/" & FileRoot & "/" & Value & """ TARGET=""__blank"">" & Value & _
            FileSize(Initializer.SiteRootDir & FileRoot & "\" & Value) & "</A>"
    
      End If
    
    ' (Campo tipo Lookup)
    Case efFldTypeLookup
      Err.Raise 300, "HiperTools30.FormField.ShowAsText", "Não implementado ainda"

    ' (Campo tipo Checkbox)
    Case efFldTypeCheck
      'Response.Write "<B>" & EditCaption(False) & "</B> "
      
      'BeginMakeLink True
      
      If sobjString.ToString(Value) = sobjString.ToString(ValueChecked) Then
        Response.Write CaptionChecked

      Else
        Response.Write CaptionNotChecked

      End If
      
      'EndMakeLink

    ' (Campos tipos Select e Radio)
    Case efFldTypeRadio, efFldTypeSelect
      'Pensar na possibilidade de usar um Pack
      
      'BeginMakeLink True
      For Each typFieldItem In FieldItems
        If sobjString.ToString(Value) = sobjString.ToString(typFieldItem.Value) Then
          Response.Write typFieldItem.Caption

        End If

      Next
      'EndMakeLink

  End Select

End Sub

Public Function GetCaption(Optional ByVal strCaption As String = "", _
                       Optional ByVal blnWithTwoDots As Boolean = False, _
                       Optional ByVal blnBoldIt As Boolean = True, _
                       Optional ByVal blnDoUnderline As Boolean = True)
  Const lconAmpersAnd = "&"
  Dim intPos As Integer
  
  InstanceForm
  
  If strCaption > "" Then Caption = strCaption
  
  If Caption > "" Then
    intPos = InStr(Caption, lconAmpersAnd)
    If intPos > 0 Then
      If Parent.Mode = frModeEdit And blnDoUnderline Then
        GetCaption = Left(Caption, intPos - 1) & "<U>" & Mid(Caption, intPos + 1, 1) & "</U>" & Mid(Caption, intPos + 2)
        objPForm.AccessKey = sobjString.ChangeChar(Mid(Caption, intPos + 1, 1))
  
      Else
        GetCaption = Left(Caption, intPos - 1) & Mid(Caption, intPos + 1)
  
      End If
  
    Else
      GetCaption = Caption
    End If
    
    If blnBoldIt Then GetCaption = "<STRONG>" & GetCaption & "</STRONG>"
    If blnWithTwoDots Then GetCaption = GetCaption & " : "
  
  End If

End Function

Public Sub ShowCaption(Optional ByVal strCaption As String = "", _
                       Optional ByVal blnWithTwoDots As Boolean = False, _
                       Optional ByVal blnBoldIt As Boolean = True, _
                       Optional ByVal blnDoUnderline As Boolean = True)
  
  Response.Write GetCaption(strCaption, blnWithTwoDots, blnBoldIt, blnDoUnderline)

End Sub

Public Sub ShowBreak(Optional ByVal intNumNbsp As Integer = 1, Optional ByVal bytCaptionPos As Byte = 3)
  
  If bytCaptionPos < 3 Then CaptionPosition = bytCaptionPos
  
  If EditOption <> efFldTypeCheck Then
    If CaptionPosition = 0 Then
      sobjShow.BR
    Else
      sobjShow.Nbsp intNumNbsp
    End If
  End If

End Sub

Public Sub ShowImages()
' (Apresenta as imagens de campo requerido e erro de preenchimento)

' (Campo requerido)
  If RequireMessage > "" Or Required Then
    Response.Write " "
    sobjShow.Image Initializer.HiperToolsWorkPath & RequiredImage, "", "Campo obrigatório", "", "", 5, 5
        
  End If

' (Erro de preenchimento)
  If ValidationError > "" Then
    Response.Write " "
    sobjShow.Image Initializer.HiperToolsWorkPath & ErrorImage, "", ValidationError, "", "", 5, 5
    
    ValidationError = ""

  End If

End Sub

Public Function Valid(Optional vntValue As Variant = "") As Boolean
  ' (Faz a validação de um campo observando sua opção de validação)
  Dim blnIsValid As Boolean
  blnIsValid = True
  
  If (RequireMessage > "") Then
    If Not CStr(vntValue) > "" Then
      ValidationError = RequireMessage
      blnIsValid = False

    End If

  ElseIf Required Then
    If Not CStr(vntValue) > "" Then
      ValidationError = "O campo " & GetCaption(, , , False) & " deve ser preenchido"
      blnIsValid = False

    End If
  Else
    Select Case Validation
      Case efValOptCGC
        If Not Validate.IsCGC(vntValue) Then
          blnIsValid = False
          ValidationError = "O campo " & GetCaption(, , , False) & " é um CGC inválido"
  
        End If
  
      Case efValOptCPF
        If Not Validate.IsCPF(vntValue) Then
          blnIsValid = False
          ValidationError = "O campo " & GetCaption(, , , False) & " é um CPF inválido"
  
        End If
  
      Case efValOptDate
        If TypeName(vntValue) = "string" Then
          If Not Validate.IsValidDate(vntValue) Then
            blnIsValid = False
            ValidationError = "O campo " & GetCaption(, , , False) & " é uma Data inválida"

          End If
        End If
 
      Case efValOptTime
        If Not Validate.IsValidTime(vntValue) Then
          blnIsValid = False
          ValidationError = "O campo " & GetCaption(, , , False) & " é uma Hora inválida"
  
        Else
          vntValue = Validate.FormatTime(vntValue)
  
        End If
  
      Case efValOptDateMToday
        If Validate.IsValidDate(vntValue) Then
          If CDate(vntValue) > Now + 1 Then
            blnIsValid = False
            ValidationError = "O campo " & GetCaption(, , , False) & " é maior que hoje"

          Else
            vntValue = Validate.FormatDate(vntValue)
          End If
  
        Else
          blnIsValid = False
          ValidationError = "O campo " & GetCaption(, , , False) & " é uma Data inválida"
  
        End If
  
      Case efValOptEmail
        If Not Validate.IsEmail(vntValue) Then
          blnIsValid = False
          ValidationError = "O campo " & GetCaption(, , , False) & " é um Email inválido"
  
        End If
      
      Case efValOptSepDate
      
      Case efValOptCompareDates
        Rem If FindField(CurrentTable, Session(strField & "DoubleValidateField")) Then
        Rem   Select Case Validate.CompareDate(Session(strField & "vntValue"), Session(CurrentField & "vntValue"))
        Rem     Case 1
        Rem       blnIsValid = False
        Rem       If FindField(CurrentTable, Session(strField & "Field")) Then
        Rem         Session("Message") = Session("Message") & ErrorImage & "O campo <B>" & EditCaption(False) & "</B> foi preenchido de forma inválida</B><BR>"
        Rem         Session(CurrentField & "ValidationError") = "O campo " & EditCaption(False) & " foi preenchido de forma inválida"
        Rem
        Rem       End If
        Rem
        Rem     Case 2
        Rem       blnIsValid = False
        Rem       Session("Message") = Session("Message") & ErrorImage & "O campo <B>" & EditCaption(False) & "</B> foi preenchido de forma inválida</B><BR>"
        Rem       Session(CurrentField & "ValidationError") = "O campo " & EditCaption(False) & " foi preenchido de forma inválida"
        Rem
        Rem     Case 3
        Rem       blnIsValid = False
        Rem       If FindField(CurrentTable, Session(strField & "Field")) Then
        Rem         Session("Message") = Session("Message") & ErrorImage & "O campo <B>" & EditCaption(False) & "</B> foi preenchido de forma inválida</B><BR>"
        Rem         Session(CurrentField & "ValidationError") = "O campo " & EditCaption(False) & " foi preenchido de forma inválida"
        Rem         If FindField(CurrentTable, Session(CurrentField & "DoubleValidateField")) Then
        Rem           Session("Message") = Session("Message") & ErrorImage & "O campo <B>" & EditCaption(False) & "</B> foi preenchido de forma inválida</B><BR>"
        Rem           Session(CurrentField & "ValidationError") = "O campo " & EditCaption(False) & " foi preenchido de forma inválida"
        Rem
        Rem         End If
        Rem
        Rem       End If
        Rem
        Rem     Case 4
        Rem       blnIsValid = False
        Rem       Session("Message") = Session("Message") & ErrorImage & Session(strField & "DoubleValidateMessage")
        Rem       Session(CurrentField & "ValidationError") = Session(strField & "DoubleValidateMessage")
        Rem
        Rem   End Select
        Rem End If
  
    End Select
  End If

  Rem If Not blnIsValid And Validation <> efValOptCompareDates Then
      Rem ValidationError = "O campo " & GetCaption(, , , False) & " foi preenchido de forma inválida"

  Rem End If

  Valid = blnIsValid

End Function

Public Sub Read()
  Dim vntValue As Variant
  
  If Request.Form(Name) > "" Then
    vntValue = Request.Form(Name)
    
    If Valid(vntValue) Then Value = vntValue
  End If

End Sub

'Privadas:
Private Sub ShowTextField()
  Select Case EditOption
  Case efFldTypeText, efFldTypeHTTP, efFldTypeEMail
    ShowImages
    ShowBreak
    objPForm.Text Name, SizeToShow, _
    sobjString.ToString(DefaultValue), SizeToGet
  End Select
End Sub

Private Sub ShowFileField()
  Select Case EditOption
  Case efFldTypeFile, efFldTypeImage
    'Verifica o estado do Session, se estiver vazio não apresenta nenhum link
    If Parent.Mode = frModeEdit Then
        Response.Write "<A HREF=""\SendFile.asp?O=1&Folder=" & _
          Server.URLEncode(FileRoot) & "&Field=" & _
          Name & "&FileName=" & _
          Server.URLEncode(FileNamePrefix)
    End If
  
    'Implementar depois (verificar com chave dupla):
    'For i = 1 To Session(CurrentTable & "FieldNumber")
    '  If Session(CurrentTable & i & "IsKey") Then
    '     Response.Write Session(CurrentTable & i & "Value")
    '  End If
    'Next
    
    'Verifica o estado do Session, se estiver vazio não apresenta nenhum link
    
    If Parent.Mode = frModeEdit Then
        Response.Write """ TARGET=""__top"" CLASS=""SS"">"
    End If
  
    ShowCaption
    'Verifica o estado do Session, se estiver vazio não apresenta nenhum link
    If Parent.Mode = frModeEdit Then
        Response.Write "</A>"
    End If
  
    If Value > "" Then
      
      If EditOption = efFldTypeFile Then
        sobjShow.HTMLCR " - <A HREF='/" & FileRoot & "/" & Value & "' TARGET='__blank'>" & Value & FileSize(Initializer.SiteRootDir & FileRoot & "\" & Value) & "</A>"
      ElseIf EditOption = efFldTypeImage Then
        sobjShow.Image "/" & FileRoot & "/" & Value, IIf(ID > -1, ID, ""), "/" & FileRoot & "/" & Value, "'/" & FileRoot & "/" & Value & "' TARGET=__blank", "", IIf(ImageWidth > -1, ImageWidth, ""), IIf(ImageHeight > -1, ImageHeight, "")
      End If
    
    End If
  End Select
End Sub

Private Sub ShowRadioField()
  Dim typFieldItem As Variant
  Dim blnFirst As Boolean, i As Integer
  
  If EditOption = efFldTypeRadio Then
    ' (Monta uma tabela que organiza os radios)
    Table.BeginTable "100%", "", 1
    Table.ColumnVAlign = "MIDDLE"
    Table.BeginRow 1
    blnFirst = True
  
    i = 1
    ' (Para cada valor do radio)
    For Each typFieldItem In FieldItems
      ' (Faz as quebras de linha conforme o nº de itens por coluna definido no
      ' Session(CurrentField & "RadioColumns"))
      If (i - 1) Mod RadioColumns = 0 Then
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
      objPForm.Radio Name, typFieldItem.Value, "LEFT", BitBln(sobjString.ToString(Value) = sobjString.ToString(typFieldItem.Value))
      Table.EndColumn
      
  ' (Coloca o texto do item)
      Table.ColumnWidth = ""
      Table.BeginColumn
      Response.Write "&nbsp;" & typFieldItem.Caption
      Table.EndColumn
      
      i = i + 1
    Next
    
    Table.EndRow
    Table.EndTable
  End If
End Sub

Private Sub ShowSelectField()
  Dim typFieldItem As Variant
  
  If EditOption = efFldTypeSelect Then
    ShowImages
    ShowBreak
    
    objPForm.SelectBegin Name, 1, False
    If Not Value > "" Then
      objPForm.SelectOption "", "", True
    
    End If
    
    For Each typFieldItem In FieldItems
      objPForm.SelectOption typFieldItem.Value, typFieldItem.Caption, _
          (sobjString.ToString(Value) = sobjString.ToString(typFieldItem.Value))
    Next
    objPForm.SelectEnd
  End If
End Sub

Private Sub ShowEspecialSelectField()
  Select Case EditOption
  Case efFldTypeUF, efFldTypeCountry
    ShowImages
    ShowBreak

    If Not Enabled Then
      objPForm.SelectBegin Name, 1, False
      If Value > "" Then
        objPForm.SelectOption Value, Value, True
        
      Else
        objPForm.SelectOption "", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", True
        
      End If
      objPForm.SelectEnd
      
    Else
      If EditOption = efFldTypeUF Then
        objPForm.SelectState Name, 1, False, sobjString.ToString(Value)
      ElseIf EditOption = efFldTypeCountry Then
        objPForm.SelectCountry Name, 1, False, sobjString.ToString(Value)
      End If
    
    End If
  End Select
End Sub

Private Sub ShowSepDateField()
  Dim i As Integer
  
  If EditOption = efFldTypeSeparateDate Then
    ShowImages
    ShowBreak

    ' (Dia)
    objPForm.SelectBegin Name & "DayValue", 1, False
    For i = 1 To 31
      objPForm.SelectOption sobjString.LeadingZeroes(i, 2), sobjString.LeadingZeroes(i, 2), i = Day(Value)
    Next
    objPForm.SelectEnd

    Response.Write "<B>/</B>"

    ' (Mês)
    objPForm.SelectBegin Name & "MonthValue", 1, False
    For i = 1 To 12
      objPForm.SelectOption sobjString.LeadingZeroes(i, 2), sobjString.LeadingZeroes(i, 2), i = Month(Value)
    Next
    objPForm.SelectEnd

    Response.Write "<B>/</B>"

    ' (Ano)
    ' (As variáveis beginyear e endyear indicam o primeiro e o último ano do select)
    If BeginYear > 0 Then
      BeginYear = Year(Now) - 20

    End If
    
    If EndYear > 0 Then
      EndYear = Year(Now) + 20

    End If

    objPForm.SelectBegin Name & "YearValue", 1, False
    For i = BeginYear To EndYear
      objPForm.SelectOption i, i, i = Year(Value)
    Next
    objPForm.SelectEnd
  End If
End Sub

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

Private Function FileSize(ByVal FileName As String) As String
' (Transforma o tamanho do arquivo de bytes para a unidade mais apropriada)
Dim objFile As Object
Dim objFileSystem As Object
    
    Set objFileSystem = Server.CreateObject("Scripting.FileSystemObject")
    If objFileSystem.FileExists(FileName) Then
      Set objFile = objFileSystem.GetFile(FileName)
      FileSize = " (" & FileSizeFormat(objFile.Size) & ")"

    Else
      FileSize = " (Arquivo não encontrado)"

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
      strSizeUnit = "EB - Eta quanto Byte"

  End Select
  FileSizeFormat = sobjString.ZTINumber(Round(dblSize, 2)) & " " & strSizeUnit
  
End Function

Private Function BitBln(ByVal blnValor As Boolean) As Byte
' (Troca um valor booleano para 0 ou 1, parametro: Valor booleano)
  If blnValor Then
    BitBln = 1

  Else
    BitBln = 0

  End If

End Function

Private Sub InstanceForm()
  If objPForm Is Nothing Then Set objPForm = Server.CreateObject("HiperTools30.Form")
End Sub

Private Sub DestroyForm()
  If Not objPForm Is Nothing Then Set objPForm = Nothing
End Sub
