VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Validate"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =========================================================================
'  /HiperTools/Validate.inc
' -------------------------------------------------------------------------
'  Descricao   : Include de validacoes JavaScript
'  Criacao     : 08:30h 28/12/1998
'  Local       : Brasilia/DF
'  Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
'              : Eduardo Alves Cortes <edualves@zevallos.com.br>
'  Versao      : 1.0.0
'  Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------
'  ALTERACOES
' -------------------------------------------------------------------------
'  Responsavel : Ridai Govinda Pombo <ridai@zevallos.com.br>
'  Data/Hora   : 09/01/2000 15:20
'  Resumo      : Inclusão de novas Validações
' -------------------------------------------------------------------------

Private Request As Request, _
        Response As Response, _
        Server As Server, _
        Session As Session, _
        Application As Application, _
        ScriptContext As ScriptingContext

Private sobjString As HiperTools30.String, _
        Functions As HiperTools30.Pack, _
        sobjShow As HiperTools30.Show, _
        sblnEmptyShowed As Boolean, _
        sstrFormValidated As String

Private sstrEmptyFuncName As String, _
        sstrIsTimeFuncName As String, _
        sstrFrmName As String

Private sstrComments As String, _
        sstrTitle As String, _
        sstrVersion As String, _
        sstrCompanyName As String, _
        sstrProductName As String, _
        sstrLegalCopyright As String

' Permite ou não que o "alert('...';)" seja ativado ou não
' (colocar em todas as rotinas que apresentam JSAlerts...)
Public AllowJSAlert As Boolean
Public JSStart As Boolean

Option Explicit

Public Sub OnStartPage(ScriptContext As ScriptingContext)

  TimeBombX xtbValidate, 15656, 12934, 3080, 14624

  Set ScriptContext = ScriptContext

  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application
  JSStart = False
  sblnEmptyShowed = False
  AllowJSAlert = True
  sstrFormValidated = ""

  '(Ridai Govinda)
  'Nomes de funções:
  'Depois criar um collection privado que controle o nome das funções apresentadas
  'para isso criar as funções de validação que recebem parâmetros mudando o comportamento
  'atual do objeto...
  sstrEmptyFuncName = ""
  sstrIsTimeFuncName = ""
  sstrFrmName = ""
  
  Set sobjString = Server.CreateObject("HiperTools30.String")
  Set Functions = Server.CreateObject("HiperTools30.Pack")
  Set sobjShow = Server.CreateObject("HiperTools30.Show")

End Sub

Public Sub OnEndPage()
  Set ScriptContext = Nothing

  Set Request = Nothing
  Set Response = Nothing
  Set Server = Nothing
  Set Session = Nothing
  Set Application = Nothing

  Set sobjShow = Nothing

End Sub

Public Sub Class_Initialize()
    sstrComments = App.Comments
    sstrTitle = App.Title
    sstrVersion = App.Major & "." & _
                App.Minor & "." & _
                App.Revision
    sstrCompanyName = App.CompanyName
    sstrProductName = App.ProductName

End Sub

Public Property Get Comments() As String
  Comments = sstrComments
End Property

Public Property Get Title() As String
  Title = sstrTitle
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

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub JSBegin()
    BeginJS
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub BeginJS()
  
  If Not JSStart Then
    sobjShow.HTMLCR "<SCRIPT LANGUAGE=""JavaScript"">"
    JSStart = True
  End If

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub JSEnd()
    EndJS
End Sub


Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub EndJS()
  
  sobjShow.HTMLCR "</SCRIPT>"
  JSStart = False

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub JavaBegin(ByVal strFunctionName As String, _
                     ByVal strFormName As String)
    BeginJava strFunctionName, strFormName
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub BeginJava(ByVal strFunctionName As String, _
                     ByVal strFormName As String)

  JSBegin
  sobjShow.HTMLCR "function " & strFunctionName & "() {"
  sstrFrmName = strFormName

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub JavaEnd()
    EndJava
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub EndJava()
  sobjShow.HTMLCR "return true;"
  sobjShow.HTMLCR "}"
  JSEnd

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
'Criado (Ridai Govinda)
Public Sub IsFieldEmptyOnClient(Optional ByVal strFuncName As String = "")
  If Not strFuncName > "" Then strFuncName = "Vazio"
  
  sstrEmptyFuncName = strFuncName
  
  JSBegin
  sobjShow.HTMLCR "function " & strFuncName & "(input, text){"
  sobjShow.HTMLCR "if(input.value.length==0){"
  sobjShow.HTMLCR "alert(text);"
  sobjShow.HTMLCR "input.focus();"
  sobjShow.HTMLCR "return true;"
  sobjShow.HTMLCR "}}"
  sblnEmptyShowed = True

End Sub
'---------------------

Public Sub JavaEmpty()
  IsFieldEmptyOnClient
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
'Criado (Ridai Govinda)
Public Sub IsRequiredOnClient(ByVal strField As String, _
                              ByVal strMessage As String)
  If Not sblnEmptyShowed Then
    JavaEmpty

  End If
  sobjShow.HTMLCR "if(" & sstrEmptyFuncName & "(document." & sstrFrmName & "." & strField & ",'" & strMessage & "')){return false;}"

End Sub
'---------------------

Public Sub JavaRequired(ByVal strField As String, _
                        ByVal strMessage As String)
  IsRequiredOnClient strField, strMessage
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
'Criado (Ridai Govinda)
Public Sub CompareFieldsOnClient(ByVal strField As String, _
                                 ByVal strField2 As String, _
                                 ByVal strMessage As String)
  sobjShow.HTMLCR "if(document." & sstrFrmName & "." & strField & ".value != document." & sstrFrmName & "." & strField2 & ".value) {"
  sobjShow.HTMLCR "alert(""" & strMessage & """);"
  sobjShow.HTMLCR "document." & sstrFrmName & "." & strField2 & ".focus();"
  sobjShow.HTMLCR "return false;"
  sobjShow.HTMLCR "}"

End Sub
'---------------------

Public Sub JavaCompare(ByVal strField As String, _
                       ByVal strField2 As String, _
                       ByVal strMessage As String)
  CompareFieldsOnClient strField, strField2, strMessage
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
'Criado (Ridai Govinda)
Public Sub ClearFieldOnClient(ByVal strField As String)
  sobjShow.HTMLCR "document." & sstrFrmName & "." & strField & ".value = """""
End Sub
'---------------------

Public Sub JavaClearField(ByVal strField As String)
  ClearFieldOnClient strField
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
'Criado (Ridai Govinda)
Public Sub SetFocusOnClient(ByVal strField As String)
  sobjShow.HTMLCR "<SCRIPT LANGUAGE=""JavaScript"">"
  sobjShow.HTMLCR "document." & sstrFrmName & "." & strField & ".focus();"
  sobjShow.HTMLCR "</SCRIPT>"

End Sub
'---------------------

Public Sub JavaFocus(ByVal strField As String)
  SetFocusOnClient strField
End Sub

Rem =========================================================================
Rem Entra FormJavaValidMask
Rem -------------------------------------------------------------------------
Public Sub JSPutMask()
  Dim strFunction As String

  strFunction = "JSPutMask"
  JSBegin
  If Not Functions.Find(strFunction) Then
  
    sobjShow.HTMLCR "function htPutMask(mask,texto){"
    sobjShow.HTMLCR "  var i, j, result"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR "  i = 0; j = 0; result = '';"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR "  while(i<=mask.length){"
    sobjShow.HTMLCR "    switch(mask.charAt(i)){"
    sobjShow.HTMLCR "    case '0':"
    sobjShow.HTMLCR "      if(!isNaN(parseInt(texto.charAt(j)))) result = result.concat(parseInt(texto.charAt(j)));"
    sobjShow.HTMLCR "      else result = result.concat('0');"
    sobjShow.HTMLCR "      break;"
    sobjShow.HTMLCR "    case 'A':"
    sobjShow.HTMLCR "      if(((texto.charAt(j)>='a' && texto.charAt(j)<='z') || (texto.charAt(j)>='A' && texto.charAt(j)<='Z')))"
    sobjShow.HTMLCR "        result = result.concat(parseInt(texto.charAt(j)));"
    sobjShow.HTMLCR "      else result = result.concat('A');"
    sobjShow.HTMLCR "      break;"
    sobjShow.HTMLCR "    default:"
    sobjShow.HTMLCR "      result = result.concat(mask.charAt(i));"
    sobjShow.HTMLCR "      j--;"
    sobjShow.HTMLCR "    }"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR "    if(j<=texto.length) j++;"
    sobjShow.HTMLCR "    i++"
    sobjShow.HTMLCR "  }"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR "  return result;"
    sobjShow.HTMLCR "}"
    
    Functions.Add strFunction

  End If

End Sub

Rem =========================================================================
Rem Entra FormJavaValidMask
Rem -------------------------------------------------------------------------
Public Sub JSMaskVal()

  Dim strFunction As String

  strFunction = "JSMaskVal"
  JSBegin
  If Not Functions.Find(strFunction) Then

    JSBegin
    
    sobjShow.HTMLCR "function htMask(parMask, parReceive){"
    sobjShow.HTMLCR " var i, vchar, mchar, intCount;"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR " intCount = 0;"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR " for(i=0; i <= parMask.length; i++){"
    sobjShow.HTMLCR "   mchar = parMask.charAt(i);"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR "   if(i <= parReceive.length){"
    sobjShow.HTMLCR "    vchar = parReceive.charAt(i);"
    sobjShow.HTMLCR "   }"
    sobjShow.HTMLCR "   else{"
    sobjShow.HTMLCR "     vchar = """";"
    sobjShow.HTMLCR "   }"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR "   switch(mchar){"
    sobjShow.HTMLCR "   case '0':"
    sobjShow.HTMLCR "     if(isNaN(parseInt(vchar))) intCount++;"
    sobjShow.HTMLCR "     break;"
    sobjShow.HTMLCR "   case 'X':"
    sobjShow.HTMLCR "     if(vchar > """"){"
    sobjShow.HTMLCR "      if(isNaN(parseInt(vchar))) intCount++;"
    sobjShow.HTMLCR "     }"
    sobjShow.HTMLCR "     break;"
    sobjShow.HTMLCR "   case 'A':"
    sobjShow.HTMLCR "     if(!((vchar>='a' && vchar<='z') || (vchar>='A' && vchar<='Z'))) intCount++;"
    sobjShow.HTMLCR "     break;"
    sobjShow.HTMLCR "   default:"
    sobjShow.HTMLCR "     if(vchar!=mchar) intCount++;"
    sobjShow.HTMLCR "   }"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR " }"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR " if(intCount>0) return false; else return true;"
    sobjShow.HTMLCR "}"

    Functions.Add strFunction

  End If

End Sub

Rem =========================================================================
Rem Entra FormJavaValidSepDate
Rem -------------------------------------------------------------------------
'Criado (Ridai Govinda)
Public Sub SepDateOnClient(ByVal strField As String, _
                           ByVal strField2 As String, _
                           ByVal strField3 As String, _
                           ByVal strForm As String, _
                           ByVal strFunctionName As String, _
                           Optional ByVal blnMaiorQueHoje As Boolean = False, _
                           Optional ByVal strMessage As String = "")
  JSBegin
  JSIsDate
  JSStringToDate
  sobjShow.HTMLCR "function " & strFunctionName & "(){"
  sobjShow.HTMLCR "var dataatual, data, text;"

  sobjShow.HTMLCR "if((document." + strForm + "." & strField & ".value > """")  && (document." + strForm + "." & strField2 & ".value > """") && (document." + strForm + "." & strField3 & ".value > """")){"
  sobjShow.HTMLCR "text = document." + strForm + "." & strField & ".value + document." + strForm + "." & strField2 & ".value + document." + strForm + "." & strField3 & ".value;"
  sobjShow.HTMLCR "if(!isdate(text)){"
  If AllowJSAlert Then _
    sobjShow.HTMLCR "alert(""Data Inválida"");"
  sobjShow.HTMLCR "document." + strForm + "." & strField & ".focus();"
  sobjShow.HTMLCR "return false;"
  sobjShow.HTMLCR "}}"

  If blnMaiorQueHoje Then
    sobjShow.HTMLCR "dataatual  = new Date();"
    sobjShow.HTMLCR "data       = stringtodate(text);"
    sobjShow.HTMLCR "if(data.valueOf() + 90000000 < dataatual.valueOf()) {"
    
    If AllowJSAlert Then _
        sobjShow.HTMLCR "alert(""" & strMessage & """);"
    
    sobjShow.HTMLCR "document." + strForm + "." & strField & ".focus();"
    sobjShow.HTMLCR "return false;"
    sobjShow.HTMLCR "}"

  End If

Rem  If sstrLastValidate > "" Then
Rem    sobjShow.HTMLCR "return " & sstrLastValidate & "()"

Rem  End If
  sobjShow.HTMLCR "}"

End Sub
'---------------------

Public Sub FormJavaValidSepDate(ByVal strField As String, _
                                ByVal strField2 As String, _
                                ByVal strField3 As String, _
                                ByVal strForm As String, _
                                ByVal strFunctionName As String, _
                                ByVal blnMaiorQueHoje As Boolean, _
                                ByVal strMessage As String)

  SepDateOnClient strField, strField2, strField3, strForm, _
                  strFunctionName, blnMaiorQueHoje, strMessage
  
End Sub

Rem =========================================================================
Rem Entra FormJavaValidDate
Rem -------------------------------------------------------------------------
'Criado (Ridai Govinda)
Public Sub IsDateOnClient(ByVal strField As String, _
                          ByVal strForm As String, _
                          ByVal strFunctionName As String, _
                          Optional ByVal blnMaiorQueHoje As Boolean = False, _
                          Optional ByVal strMessage As String = "")

  JSBegin
  JSIsDate
  JSStringToDate
  JSDateToString
  
  sobjShow.HTMLCR "function " & strFunctionName & "(){"
  sobjShow.HTMLCR "var dataatual, data, text;"

  sobjShow.HTMLCR "if(document." + strForm + "." & strField & ".value > """"){"
  sobjShow.HTMLCR "text = document." + strForm + "." & strField & ".value;"
  sobjShow.HTMLCR "if(!isdate(text)){"
  If AllowJSAlert Then _
    sobjShow.HTMLCR "alert(""Data Inválida"");"
  sobjShow.HTMLCR "document." + strForm + "." & strField & ".focus();"
  sobjShow.HTMLCR "return false;"
  sobjShow.HTMLCR "}"
  sobjShow.HTMLCR "else"
  sobjShow.HTMLCR "{"
  sobjShow.HTMLCR "document." + strForm + "." & strField & ".value = datetostring(stringtodate(text));"
  sobjShow.HTMLCR "return true}"
  sobjShow.HTMLCR "}"

  If blnMaiorQueHoje Then
    sobjShow.HTMLCR "dataatual  = new Date();"
    sobjShow.HTMLCR "data       = stringtodate(text);"
    sobjShow.HTMLCR "if(data.valueOf() + 90000000 < dataatual.valueOf()) {"
    If AllowJSAlert Then _
        sobjShow.HTMLCR "alert(""" & strMessage & """);"
    sobjShow.HTMLCR "document." + strForm + "." & strField & ".focus();"
    sobjShow.HTMLCR "return false;"
    sobjShow.HTMLCR "}"
    sobjShow.HTMLCR "else{return true}"

  End If

  'Criado (Ridai Govinda)
  sobjShow.HTMLCR "if(!document." + strForm + "." & strField & ".value > """") return true"
  '-----------------------
  sobjShow.HTMLCR "}"

End Sub
'--------------------

Public Sub FormJavaValidDate(ByVal strField As String, _
                             ByVal strForm As String, _
                             ByVal strFunctionName As String, _
                             ByVal blnMaiorQueHoje As Boolean, _
                             ByVal strMessage As String)
  
  IsDateOnClient strField, strForm, strFunctionName, blnMaiorQueHoje, strMessage

End Sub

Rem =========================================================================
Rem Valida um campo de CEP
Rem -------------------------------------------------------------------------
'Criado (Ridai Govinda)
Public Sub PostalCodeOnClient(ByVal strField As String, _
                              ByVal strForm As String, _
                              ByVal strFunctionName As String)

    Dim strFieldVld As String
    
    strFieldVld = "document." & strForm & "." & strField & ".value"
    
    JSBegin
    JSMaskVal
    JSPutMask
    JSLeadingZeroes

    sobjShow.HTMLCR "function " & strFunctionName & "(){"
    sobjShow.HTMLCR "  var blnValid"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR "  blnValid = true;"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR "  if(" & strFieldVld & " > """" ) {"
    sobjShow.HTMLCR "    if((htMask('00.000-XXX'," & strFieldVld & ")"
    sobjShow.HTMLCR "       || htMask('00000-XXX'," & strFieldVld & "))){"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR "      if(!(htMask('00.000-000'," & strFieldVld & ")"
    sobjShow.HTMLCR "         || htMask('00000-000'," & strFieldVld & "))){"
    sobjShow.HTMLCR "        " & strFieldVld & " = htPutMask('00.000-000', " & strFieldVld & ");"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR "      }"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR "    }"
    sobjShow.HTMLCR "    else{"
    sobjShow.HTMLCR "      if(!htMask('00000XXX'," & strFieldVld & ")){"
    sobjShow.HTMLCR "        if(isNaN(parseInt(" & strFieldVld & "))){"
    sobjShow.HTMLCR "          " & strFieldVld & " = """";"
    sobjShow.HTMLCR "          blnValid = false;"
    sobjShow.HTMLCR "        }"
    sobjShow.HTMLCR "        else{"
    sobjShow.HTMLCR "          " & strFieldVld & " = LeadingZeroes(" & strFieldVld & ", 8);"
    sobjShow.HTMLCR "          " & strFieldVld & " = htPutMask('00.000-000', " & strFieldVld & ");"
    sobjShow.HTMLCR "          blnValid = false;"
    sobjShow.HTMLCR "        }"
    sobjShow.HTMLCR "      }"
    sobjShow.HTMLCR "      else{"
    sobjShow.HTMLCR "        " & strFieldVld & " = htPutMask('00.000-000', " & strFieldVld & ");"
    sobjShow.HTMLCR "      }"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR "    }"
    sobjShow.HTMLCR "  }"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR "  if(!blnValid){"
    
    If AllowJSAlert Then _
      sobjShow.HTMLCR "    alert(""Este CEP é inválido"");"
    
    sobjShow.HTMLCR "    document." & strForm & "." & strField & ".focus();"
    sobjShow.HTMLCR "  }"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR "  return blnValid"
    sobjShow.HTMLCR "}"

End Sub
'---------------------

Public Sub FormJavaValidCEP(ByVal strField As String, _
                            ByVal strForm As String, _
                            ByVal strFunctionName As String)

  PostalCodeOnClient strField, strForm, strFunctionName
End Sub

Rem =========================================================================
Rem Valida um campo de CGC
Rem -------------------------------------------------------------------------
'Criado/Alterado (Ridai Govinda)
Public Sub CGCOnClient(ByVal strField As String, _
                       ByVal strForm As String, _
                       ByVal strFunctionName As String)
  JSBegin

  JSNormalize
  JSLeadingZeroes
  sobjShow.HTMLCR "function " & strFunctionName & "(){"

  sobjShow.HTMLCR "  var cgc, soma, digito, digito2, multip, i, aux, numero;"
  sobjShow.HTMLCR "  soma   = 0;"
  sobjShow.HTMLCR "  multip = 5;"

  sobjShow.HTMLCR "  cgc = document." + strForm + "." & strField & ".value;"
  sobjShow.HTMLCR "  cgc     = normalize(cgc);"
  sobjShow.HTMLCR "  if(cgc.length < 1){ "
  sobjShow.HTMLCR "    document." + strForm + "." & strField & ".focus();"
  If AllowJSAlert Then _
    sobjShow.HTMLCR "    alert(""Este CGC é inválido"");"
  sobjShow.HTMLCR "    return false;"
  sobjShow.HTMLCR "  }"
  sobjShow.HTMLCR "  cgc     = LeadingZeroes(cgc, 11);"
  sobjShow.HTMLCR "  for(i = 0;i <= 11;i++){"
  sobjShow.HTMLCR "    number = parseInt(cgc.charAt(i));"
  sobjShow.HTMLCR "    soma  += (number * multip);"
  sobjShow.HTMLCR "    multip+= - 1;"
  sobjShow.HTMLCR "    if(multip < 2) multip = 9;"
  sobjShow.HTMLCR "  }"
  sobjShow.HTMLCR "  aux = soma % 11;"
  sobjShow.HTMLCR "  if((aux == 0) || (aux == 1))"
  sobjShow.HTMLCR "    digito = 0;"
  sobjShow.HTMLCR "  else"
  sobjShow.HTMLCR "    digito = 11 - aux;"
  sobjShow.HTMLCR "  if(parseInt(cgc.charAt(12)) == digito){"
  sobjShow.HTMLCR "    multip = 6;"
  sobjShow.HTMLCR "    soma   = 0;"
  sobjShow.HTMLCR "    for(i = 0;i <= 12;i++){"
  sobjShow.HTMLCR "      numero = parseInt(cgc.charAt(i));"
  sobjShow.HTMLCR "      soma  += (numero * multip);"
  sobjShow.HTMLCR "      multip+= - 1;"
  sobjShow.HTMLCR "      if(multip < 2) multip = 9;"
  sobjShow.HTMLCR "      "
  sobjShow.HTMLCR "    }"
  sobjShow.HTMLCR "    aux = soma % 11;"
  sobjShow.HTMLCR "    if((aux == 0) || (aux == 1))"
  sobjShow.HTMLCR "      digito = 0;"
  sobjShow.HTMLCR "    else"
  sobjShow.HTMLCR "      digito = 11 - aux;"
  sobjShow.HTMLCR "    if(parseInt(cgc.charAt(13)) == digito)"
  sobjShow.HTMLCR "      return true;"
  sobjShow.HTMLCR "    else{"
  sobjShow.HTMLCR "      document." + strForm + "." & strField & ".focus();"
  If AllowJSAlert Then _
    sobjShow.HTMLCR "      alert(""Este CGC é inválido"");"
  sobjShow.HTMLCR "      return false;"
  sobjShow.HTMLCR "    }"
  sobjShow.HTMLCR "  }"
  sobjShow.HTMLCR "  else{"
  sobjShow.HTMLCR "    document." + strForm + "." & strField & ".focus();"
  If AllowJSAlert Then _
    sobjShow.HTMLCR "    alert(""Este CGC é inválido"");"
  sobjShow.HTMLCR "    return false;"
  sobjShow.HTMLCR "  }"
  sobjShow.HTMLCR "}"

End Sub

Public Sub FormJavaValidCGC(ByVal strField As String, _
                            ByVal strForm As String, _
                            ByVal strFunctionName As String)
  CGCOnClient strField, strForm, strFunctionName
  
End Sub
'----------------------

Public Sub CPFOnClient(ByVal strField As String, _
                            ByVal strForm As String, _
                            ByVal strFunctionName As String)
  FormJavaValidCPF strField, strForm, strFunctionName

End Sub


Rem =========================================================================
Rem Valida um campo de CPF
Rem -------------------------------------------------------------------------
Public Sub FormJavaValidCPF(ByVal strField As String, _
                            ByVal strForm As String, _
                            ByVal strFunctionName As String)

  JSBegin

  JSNormalize
  JSLeadingZeroes
  sobjShow.HTMLCR "function " & strFunctionName & "(){"

  sobjShow.HTMLCR "  var cpf, soma, digito, digito2, multip, i, aux, numero;"
  sobjShow.HTMLCR "  soma    = 0;"
  sobjShow.HTMLCR "  multip  = 10;"

  sobjShow.HTMLCR "  cpf     = document." + strForm + "." & strField & ".value;"
  sobjShow.HTMLCR "  cpf     = normalize(cpf);"
  sobjShow.HTMLCR "  if(cpf.length < 1){ "
  sobjShow.HTMLCR "    document." + strForm + "." & strField & ".focus();"
  If AllowJSAlert Then _
    sobjShow.HTMLCR "    alert(""Este CPF é inválido"");"
  sobjShow.HTMLCR "    return false;"
  sobjShow.HTMLCR "  }"
  sobjShow.HTMLCR "  cpf     = LeadingZeroes(cpf, 11);"
  sobjShow.HTMLCR "  for(i = 0;i <= 8;i++){"
  sobjShow.HTMLCR "    number = parseInt(cpf.charAt(i));"
  sobjShow.HTMLCR "    soma  += (number * multip);"
  sobjShow.HTMLCR "    multip+= - 1;"
  sobjShow.HTMLCR "  }"
  sobjShow.HTMLCR "  aux = soma % 11;"
  sobjShow.HTMLCR "  if((aux == 0) || (aux == 1)) digito = 0;"
  sobjShow.HTMLCR "  else digito = 11 - aux;"
  sobjShow.HTMLCR "  if(parseInt(cpf.charAt(9)) == digito){"
  sobjShow.HTMLCR "    multip = 11;"
  sobjShow.HTMLCR "    soma   = 0;"
  sobjShow.HTMLCR "    for(i = 0;i <= 9;i++){"
  sobjShow.HTMLCR "      numero = parseInt(cpf.charAt(i));"
  sobjShow.HTMLCR "      soma  += (numero * multip);"
  sobjShow.HTMLCR "      multip+= - 1;"
  sobjShow.HTMLCR "    }"
  sobjShow.HTMLCR "    aux = soma % 11;"
  sobjShow.HTMLCR "    if((aux == 0) || (aux == 1)) digito = 0;"
  sobjShow.HTMLCR "    else digito = 11 - aux;"
  sobjShow.HTMLCR "    if(parseInt(cpf.charAt(10)) == digito) return true;"
  sobjShow.HTMLCR "    else{"
  sobjShow.HTMLCR "      document." + strForm + "." & strField & ".focus();"
  If AllowJSAlert Then _
    sobjShow.HTMLCR "      alert(""Este CPF é inválido"");"
  sobjShow.HTMLCR "      return false;"
  sobjShow.HTMLCR "    }"
  sobjShow.HTMLCR "  }"
  sobjShow.HTMLCR "  else{"
  sobjShow.HTMLCR "    document." + strForm + "." & strField & ".focus();"
  If AllowJSAlert Then _
    sobjShow.HTMLCR "    alert(""Este CPF é inválido"");"
  sobjShow.HTMLCR "    return false;"
  sobjShow.HTMLCR "  }"
  sobjShow.HTMLCR "}"

End Sub
Rem ---------------------------------------------------------------------
Rem Final da function FormJavaValidCPF

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub FormJavaValidCPFCGC(ByVal strField As String, _
                               ByVal strForm As String, _
                               ByVal strFunctionName As String, _
                               Optional ByVal bytWhichFirst As Byte = 0)

  Const lconCPFFunc = "_EspCPF"
  Const lconCGCFunc = "_EspCGC"
  
  Dim blnAllowJS As Boolean

  blnAllowJS = AllowJSAlert

  AllowJSAlert = False

  FormJavaValidCPF strField, strForm, strFunctionName & lconCPFFunc
  FormJavaValidCGC strField, strForm, strFunctionName & lconCGCFunc

  sobjShow.HTMLCR "function " & strFunctionName & "(){"
  sobjShow.HTMLCR "  var blnvalid;"
  
  Select Case bytWhichFirst
  Case vlCp_CPFFirst
    sobjShow.HTMLCR "  if(" & strFunctionName & lconCPFFunc & "())"
  Case vlCp_CGCFirst
    sobjShow.HTMLCR "  if(" & strFunctionName & lconCGCFunc & "())"
  End Select
  
  sobjShow.HTMLCR "    blnvalid = true;"
  sobjShow.HTMLCR "  else{"
  
  Select Case bytWhichFirst
  Case vlCp_CPFFirst
    sobjShow.HTMLCR "    if(" & strFunctionName & lconCGCFunc & "()) blnvalid = true; "
  Case vlCp_CGCFirst
    sobjShow.HTMLCR "    if(" & strFunctionName & lconCPFFunc & "()) blnvalid = true; "
  End Select
  
  sobjShow.HTMLCR "    else blnvalid = false;"
  sobjShow.HTMLCR "  }"
  
  If blnAllowJS Then
    Select Case bytWhichFirst
    Case vlCp_CPFFirst
      sobjShow.HTMLCR "  if(!blnvalid) alert(""CPF ou CGC inválido"");"
    Case vlCp_CGCFirst
      sobjShow.HTMLCR "  if(!blnvalid) alert(""CGC ou CPF inválido"");"
    End Select
  End If
  
  sobjShow.HTMLCR "  return blnvalid;"
  sobjShow.HTMLCR "}"

  AllowJSAlert = blnAllowJS

End Sub

Rem =========================================================================
Rem Valida um campo de e-mail
Rem -------------------------------------------------------------------------
'Criado (Ridai Govinda)
Public Sub EMailOnClient(ByVal strField As String, _
                         ByVal strForm As String, _
                         ByVal strFunctionName As String)
  
  'Criado/Alterado (Ridai):
  JSBegin
  JSIsAlpha
  JSASplit

  sobjShow.HTMLCR "function " & strFunctionName & "(){"
  sobjShow.HTMLCR "   var email, Valid, strDomain, strUser, i, intCount;"
  sobjShow.HTMLCR "   var arrAux;"
  sobjShow.HTMLCR "   "
  sobjShow.HTMLCR "   email = document." & strForm & "." & strField & ".value;"
  sobjShow.HTMLCR "   if(email == """") {"
  'sobjShow.HTMLCR "       //alert(""Campo vazio"");"
  sobjShow.HTMLCR "       return true;"
  sobjShow.HTMLCR "   }"
  sobjShow.HTMLCR "   else{"
  sobjShow.HTMLCR "       Valid = false;"
  sobjShow.HTMLCR "       "
  sobjShow.HTMLCR "       if(email.charAt(email.length - 1) == "".""){"
  sobjShow.HTMLCR "         email = email.substr(0, email.length - 1);"
  sobjShow.HTMLCR "         document." & strForm & "." & strField & ".value = email;"
  sobjShow.HTMLCR "       }"
  sobjShow.HTMLCR "       "
  sobjShow.HTMLCR "       if(email.indexOf(""@"") > 0){"
  sobjShow.HTMLCR "           strDomain = email.substr(email.indexOf(""@"") + 1, email.length - (email.indexOf(""@"") + 1));"
  sobjShow.HTMLCR "           strUser = email.substr(0, email.indexOf(""@""));"
  sobjShow.HTMLCR ""
  sobjShow.HTMLCR "           arrAux = ZHT_ASplit(strUser, ""."");"
  sobjShow.HTMLCR "       "
  sobjShow.HTMLCR "       intCount = -1;"
  sobjShow.HTMLCR "       for(i = 0; i <= arrAux.length - 1; i++){"
  sobjShow.HTMLCR "         if(arrAux[i] > """" && arrAux[i].indexOf(""@"") == -1){"
  sobjShow.HTMLCR "             intCount++;"
  sobjShow.HTMLCR "         }"
  sobjShow.HTMLCR "       }"
  sobjShow.HTMLCR "       "
  sobjShow.HTMLCR "       if(intCount == arrAux.length - 1) Valid = true;"
  sobjShow.HTMLCR "       "
  sobjShow.HTMLCR "       if(Valid){"
  sobjShow.HTMLCR "           "
  sobjShow.HTMLCR "         Valid = false;"
  sobjShow.HTMLCR "         arrAux = ZHT_ASplit(strDomain, ""."");"
  sobjShow.HTMLCR "        "
  sobjShow.HTMLCR "           intCount = -1;"
  sobjShow.HTMLCR "           for(i = 0; i <= arrAux.length - 1; i++){"
  sobjShow.HTMLCR "             "
  sobjShow.HTMLCR "             if(arrAux[i] > """" && ZHT_IsAlpha(arrAux[i].charAt(0)) && arrAux[i].indexOf(""@"") == -1){"
  sobjShow.HTMLCR "             intCount++;"
  sobjShow.HTMLCR "             }"
  sobjShow.HTMLCR "           }"
  sobjShow.HTMLCR "           "
  sobjShow.HTMLCR "         if(intCount == arrAux.length - 1) Valid = true;"
  sobjShow.HTMLCR "       "
  sobjShow.HTMLCR "         }"
  sobjShow.HTMLCR "       "
  sobjShow.HTMLCR "       }"
  sobjShow.HTMLCR "       if(!Valid) {"
  sobjShow.HTMLCR "           document." & strForm & "." & strField & ".focus();"

  If AllowJSAlert Then _
  sobjShow.HTMLCR "           alert(""Este e-mail é inválido"");"

  sobjShow.HTMLCR "           return false;"
  sobjShow.HTMLCR "       }"
  sobjShow.HTMLCR "       else {"
  sobjShow.HTMLCR "           return true;"
  sobjShow.HTMLCR "       }"
  sobjShow.HTMLCR "       "
  sobjShow.HTMLCR "   }"
  sobjShow.HTMLCR "}"

Rem sobjShow.HTMLCR "function " & strFunctionName & "(){"
Rem sobjShow.HTMLCR "var email;"
Rem sobjShow.HTMLCR "email = document." + strForm + "." & strField & ".value;"
Rem sobjShow.HTMLCR "if((email == """") || ((email.indexOf(""."") != -1) && (email.indexOf(""@"") != -1)))"
Rem sobjShow.HTMLCR "return true;"
Rem sobjShow.HTMLCR "else{"
Rem sobjShow.HTMLCR "document." + strForm + "." & strField & ".focus();"
Rem sobjShow.HTMLCR "alert(""Este e-mail é inválido"");"
Rem sobjShow.HTMLCR "return false;}"
Rem sobjShow.HTMLCR "}"

End Sub
'---------------------

Public Sub FormJavaValidEMail(ByVal strField As String, _
                              ByVal strForm As String, _
                              ByVal strFunctionName As String)

  EMailOnClient strField, strForm, strFunctionName

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub JSMid()
  Dim strFunction As String

  strFunction = "JSMid"
  JSBegin
  If Not Functions.Find(strFunction) Then
    sobjShow.HTMLCR "function mid(text, start, length){return text.substr(start - 1, length);}"
    Functions.Add strFunction

  End If

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub JSLeft()
Dim strFunction As String

  strFunction = "JSLeft"
  JSBegin
  If Not Functions.Find(strFunction) Then
    sobjShow.HTMLCR "function left(text, length){return text.substr(0, length);}"
    Functions.Add strFunction

  End If

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub JSMidToEnd()
Dim strFunction As String

  strFunction = "JSMidToEnd"
  JSBegin
  If Not Functions.Find(strFunction) Then
    sobjShow.HTMLCR "function midtoend(text, start){return text.substr(start - 1);}"
    Functions.Add strFunction

  End If

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub JSNormalize()
Dim strFunction As String

  strFunction = "JSNormalize"
  JSBegin
  JSMidToEnd
  If Not Functions.Find(strFunction) Then
    sobjShow.HTMLCR "function normalize(text)"
    sobjShow.HTMLCR "{var i;"
    sobjShow.HTMLCR "i = 0;"
    sobjShow.HTMLCR "while (i < text.length){"
    sobjShow.HTMLCR "if((text.substr(i, 1) < ""0"") || (text.substr(i, 1) > ""9"")){"
    sobjShow.HTMLCR "text=text.substr(0,i)+midtoend(text,i+2);}"
    sobjShow.HTMLCR "else{"
    sobjShow.HTMLCR "i++;}}"
    sobjShow.HTMLCR "return text;"
    sobjShow.HTMLCR "}"
    Functions.Add strFunction

  End If

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub JSIsAlpha(Optional ByVal blnAccents As Boolean = False)
  Dim strFunction As String

  strFunction = "JSIsAlpha"
  JSBegin

  If Not Functions.Find(strFunction) Then
    sobjShow.HTMLCR "function ZHT_IsAlpha(strChar){"
    sobjShow.HTMLCR "  var i, blnValid;"
    sobjShow.HTMLCR "  "
    sobjShow.HTMLCR "  blnValid = false;"
    sobjShow.HTMLCR "  for(i = 0; i < strChar.length; i++){"
    sobjShow.HTMLCR "    if((strChar.charAt(i) >= ""a"" && strChar.charAt(i) <= ""z"") "
    sobjShow.HTML "    || (strChar.charAt(i) >= ""A"" && strChar.charAt(i) <= ""Z"")"
    
    If blnAccents Then
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""á"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""ä"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""â"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""à"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""ã"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""é"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""ë"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""ê"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""è"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""í"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""ï"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""î"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""ì"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""ó"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""ö"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""ô"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""ò"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""õ"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""ú"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""ü"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""û"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""ù"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""ç"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""ñ"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""ý"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""ÿ"")"
    
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Á"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Ä"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Â"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""À"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Ã"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""É"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Ë"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Ê"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""È"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Í"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Ï"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Î"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Ì"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Ó"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Ö"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Ô"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Ò"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Õ"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Ú"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Ü"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Û"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Ù"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Ç"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Ñ"")"
        sobjShow.HTMLCR "    || (strChar.charAt(i) = ""Ý"") )"
    Else
        sobjShow.HTMLCR " ) "
    End If
    
    sobjShow.HTMLCR "    blnValid = true;"
    sobjShow.HTMLCR "  }"
    sobjShow.HTMLCR "  return blnValid;"
    sobjShow.HTMLCR "}"
    Functions.Add strFunction
  
  End If

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub JSASplit()
  Dim strFunction As String

  JSBegin
  strFunction = "JSASplit"

  If Not Functions.Find(strFunction) Then
    sobjShow.HTMLCR "function ZHT_ASplit(strString, charSep){"
    sobjShow.HTMLCR "  var arrAux = new Array(), i, j;"
    sobjShow.HTMLCR "  "
    sobjShow.HTMLCR "  "
    sobjShow.HTMLCR " if(charSep.length == 1){"
    sobjShow.HTMLCR " "
    sobjShow.HTMLCR "   j = 0;"
    sobjShow.HTMLCR "    arrAux[j] = """";"
    sobjShow.HTMLCR "   for(i = 0; i < strString.length; i++){"
    sobjShow.HTMLCR "     if(strString.charAt(i) != charSep){"
    sobjShow.HTMLCR "       arrAux[j] = arrAux[j] + strString.charAt(i);"
    sobjShow.HTMLCR "     }"
    sobjShow.HTMLCR "     else{"
    sobjShow.HTMLCR "       j++;"
    sobjShow.HTMLCR "       arrAux[j] = """";"
    sobjShow.HTMLCR "     }"
    sobjShow.HTMLCR "   }"
    sobjShow.HTMLCR ""
    sobjShow.HTMLCR "   return arrAux;"
    sobjShow.HTMLCR "  }"
    sobjShow.HTMLCR "  else return NaN;"
    sobjShow.HTMLCR "}"

    Functions.Add strFunction
  End If
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub JSLeadingZeroes()
  Dim strFunction As String

  strFunction = "JSLeadingZeroes"
  JSBegin
  If Not Functions.Find(strFunction) Then
    sobjShow.HTMLCR "function LeadingZeroes(text, tam){"
    sobjShow.HTMLCR "while (text.length < tam){"
    sobjShow.HTMLCR "text = ""0"" + text;}"
    sobjShow.HTMLCR "return text;"
    sobjShow.HTMLCR "}"
    Functions.Add strFunction

  End If

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub JSIsDate()
Dim strFunction As String

  strFunction = "JSIsDate"
  JSBegin
  JSNormalize
  JSMid
  JSMidToEnd
  JSLeft
  If Not Functions.Find(strFunction) Then
    sobjShow.HTMLCR "function isdate(text){"
    sobjShow.HTMLCR "var Dia, Mes, Ano, data;"
    sobjShow.HTMLCR "text = normalize(text);"
    sobjShow.HTMLCR "if(text.length > 1){"
    sobjShow.HTMLCR "data = new Date();"
    sobjShow.HTMLCR "Dia = parseFloat(left(text,2));"
    sobjShow.HTMLCR "if(text.length > 3){"
    sobjShow.HTMLCR "Mes = parseFloat(mid(text,3,2));"
    sobjShow.HTMLCR "if(text.length > 5)"
    sobjShow.HTMLCR "Ano = parseFloat(midtoend(text,5));"
    sobjShow.HTMLCR "else"
    sobjShow.HTMLCR "Ano = data.getFullYear();}"
    sobjShow.HTMLCR "else"
    sobjShow.HTMLCR "Mes = data.getMonth() + 1;"
    sobjShow.HTMLCR "if((Mes > 12) || (Mes < 1) || (Ano > data.getFullYear() + 60))"
    sobjShow.HTMLCR "return false;"
    sobjShow.HTMLCR "else{"
    sobjShow.HTMLCR "switch(Mes){"
    sobjShow.HTMLCR "case 1:"
    sobjShow.HTMLCR "case 3:"
    sobjShow.HTMLCR "case 5:"
    sobjShow.HTMLCR "case 7:"
    sobjShow.HTMLCR "case 8:"
    sobjShow.HTMLCR "case 10:"
    sobjShow.HTMLCR "case 12:"
    sobjShow.HTMLCR "if((Dia > 31) || (Dia < 1))"
    sobjShow.HTMLCR "return false;"
    sobjShow.HTMLCR "break;"
    sobjShow.HTMLCR "case 4:"
    sobjShow.HTMLCR "case 6:"
    sobjShow.HTMLCR "case 9:"
    sobjShow.HTMLCR "case 11:"
    sobjShow.HTMLCR "if((Dia > 30) || (Dia < 1))"
    sobjShow.HTMLCR "return false;"
    sobjShow.HTMLCR "break;"
    sobjShow.HTMLCR "case 2:"
    sobjShow.HTMLCR "if(Ano%4>0){"
    sobjShow.HTMLCR "if((Dia>28)||(Dia<1))"
    sobjShow.HTMLCR "return false;}"
    sobjShow.HTMLCR "else"
    sobjShow.HTMLCR "if((Dia>29)||(Dia<1))"
    sobjShow.HTMLCR "return false;"
    sobjShow.HTMLCR "break;}}}"
    sobjShow.HTMLCR "else"
    sobjShow.HTMLCR "return false;"
    sobjShow.HTMLCR "return true;"
    sobjShow.HTMLCR "}"
    Functions.Add strFunction

  End If

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub JSStringToDate()
Dim strFunction As String

  strFunction = "JSStringToDate"
  JSBegin
  JSNormalize
  JSMid
  JSMidToEnd
  JSLeft
  If Not Functions.Find(strFunction) Then
    sobjShow.HTMLCR "function stringtodate(text){"
    sobjShow.HTMLCR "var Dia, Mes, Ano, data;"
    sobjShow.HTMLCR "text = normalize(text);"
    sobjShow.HTMLCR "data = new Date();"
    sobjShow.HTMLCR "Dia = parseFloat(left(text, 2));"
    sobjShow.HTMLCR "if(text.length > 3)"
    sobjShow.HTMLCR "Mes = parseFloat(mid(text, 3,2));"
    sobjShow.HTMLCR "else"
    sobjShow.HTMLCR "Mes = data.getMonth() + 1;"
    sobjShow.HTMLCR "if(text.length > 5)"
    sobjShow.HTMLCR "Ano = parseFloat(midtoend(text, 5));"
    sobjShow.HTMLCR "else"
    sobjShow.HTMLCR "Ano = data.getFullYear();"
    sobjShow.HTMLCR "if(Ano < 100){"
    sobjShow.HTMLCR "if(Ano + 1900 > data.getFullYear() - 70)"
    sobjShow.HTMLCR "Ano+=1900;"
    sobjShow.HTMLCR "else"
    sobjShow.HTMLCR "Ano+=2000;}"
    sobjShow.HTMLCR "data = new Date(Ano, Mes - 1, Dia);"
    sobjShow.HTMLCR "return data;"
    sobjShow.HTMLCR "}"
    Functions.Add strFunction

  End If

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub JSDateToString()
Dim strFunction As String

  strFunction = "JSDateToString"
  JSBegin
  JSLeadingZeroes
  If Not Functions.Find(strFunction) Then
    sobjShow.HTMLCR "function datetostring(data){ return LeadingZeroes(data.getDate().toString(), 2) + ""/"" + LeadingZeroes((data.getMonth() + 1).toString(), 2) + ""/"" + LeadingZeroes(data.getFullYear(), 2);}"
    Functions.Add strFunction

  End If

End Sub

Rem =========================================================================
Rem Entra FormJavaVerEscolha
Rem -------------------------------------------------------------------------
Public Sub FormJavaVerEscolha(ByVal strField As String, _
                              ByVal strForm As String, _
                              ByVal strFunctionName As String)
  JSBegin

  sobjShow.HTMLCR "function " & strFunctionName & "(){"
  sobjShow.HTMLCR "if(document." & strForm & "." & strField & ".value == ""00""){"
  sobjShow.HTMLCR "alert(""Uma das opções deve ser escolhida"");"
  sobjShow.HTMLCR "document." & strForm & "." & strField & ".focus();"
  sobjShow.HTMLCR "return false;"
  sobjShow.HTMLCR "}"
  sobjShow.HTMLCR "return true;"
  sobjShow.HTMLCR "}"

End Sub

Rem =========================================================================
Rem Entra FormJavaVerEscolha
Rem -------------------------------------------------------------------------
Public Sub FormJavaDataEscolha(ByVal strField As String, _
                               ByVal strField2 As String, _
                               ByVal strField3 As String, _
                               ByVal strForm As String, _
                               ByVal strFunctionName As String, _
                               ByVal strMessage As String)

  FormJavaVerEscolha strField3, strForm, strFunctionName & "VerEscolha"
  FormJavaCompareDate strField, strField2, strForm, strFunctionName & "CompareData", strMessage

  JSBegin

  sobjShow.HTMLCR "function " & strFunctionName & "(){"
  sobjShow.HTMLCR "return (" & strFunctionName & "VerEscolha()&&" & strFunctionName & "CompareData())"
  sobjShow.HTMLCR "}"

End Sub

Rem =========================================================================
Rem Entra FormJavaCompareDate
Rem -------------------------------------------------------------------------
'Criado (Ridai Govinda)
Public Sub CompareDateOnClient(ByVal strField As String, _
                               ByVal strField2 As String, _
                               ByVal strForm As String, _
                               ByVal strFunctionName As String, _
                               ByVal strMessage As String)

  JSBegin
  JSIsDate
  JSStringToDate
  JSDateToString
  sobjShow.HTMLCR "function " & strFunctionName & "(){"
  sobjShow.HTMLCR "var datainicio, datafim, text, text2;"
  sobjShow.HTMLCR "datainicio = new Date();"
  sobjShow.HTMLCR "datafim    = new Date();"
  sobjShow.HTMLCR "text       = document." & strForm & "." & strField & ".value;"
  sobjShow.HTMLCR "text2      = document." & strForm & "." & strField2 & ".value;"
  sobjShow.HTMLCR "if(text > """")"
  sobjShow.HTMLCR "if(!isdate(text)){"
  sobjShow.HTMLCR "alert(""Data Inválida"");"
  sobjShow.HTMLCR "document." & strForm & "." & strField & ".focus();"
  sobjShow.HTMLCR "return false;"
  sobjShow.HTMLCR "}"
  sobjShow.HTMLCR "else{"
  sobjShow.HTMLCR "datainicio = stringtodate(text);"
  sobjShow.HTMLCR "document." & strForm & "." & strField & ".value = datetostring(datainicio);"
  sobjShow.HTMLCR "}"
  sobjShow.HTMLCR "if(text2 > """")"
  sobjShow.HTMLCR "if(!isdate(text2)){"
  sobjShow.HTMLCR "alert(""Data Inválida"");"
  sobjShow.HTMLCR "document." & strForm & "." & strField2 & ".focus();"
  sobjShow.HTMLCR "return false;"
  sobjShow.HTMLCR "}"
  sobjShow.HTMLCR "else{"
  sobjShow.HTMLCR "datafim = stringtodate(text2);"
  sobjShow.HTMLCR "document." & strForm & "." & strField2 & ".value = datetostring(datafim);"
  sobjShow.HTMLCR "}"
  sobjShow.HTMLCR "if((text > """") && (text2 > """")){"
  sobjShow.HTMLCR "if(datainicio.valueOf() > datafim.valueOf()) {"
  sobjShow.HTMLCR "alert(""" & strMessage & """);"
  sobjShow.HTMLCR "document." & strForm & "." & strField & ".focus();"
  sobjShow.HTMLCR "return false;"
  sobjShow.HTMLCR "}}"
  sobjShow.HTMLCR "return true;"
  sobjShow.HTMLCR "}"

End Sub

Public Sub FormJavaCompareDate(ByVal strField As String, _
                               ByVal strField2 As String, _
                               ByVal strForm As String, _
                               ByVal strFunctionName As String, _
                               ByVal strMessage As String)
  CompareDateOnClient strField, strField2, strForm, strFunctionName, strMessage
End Sub
'---------------------

Rem =========================================================================
Rem Entra FormJavaCompareDateSep
Rem -------------------------------------------------------------------------
'Criado/Alterado (Ridai Govinda)
Public Sub CompareSepDateOnClient(ByVal strField As String, _
                                  ByVal strField2 As String, _
                                  ByVal strField3 As String, _
                                  ByVal strField4 As String, _
                                  ByVal strField5 As String, _
                                  ByVal strField6 As String, _
                                  ByVal strForm As String, _
                                  ByVal strFunctionName As String, _
                                  ByVal strMessage As String, _
                                  Optional ByVal blnMaiorQueHoje As Boolean = False, _
                                  Optional ByVal strMessage2 As String)

  JSBegin
  JSIsDate
  sobjShow.HTMLCR "function " & strFunctionName & "(){"
  sobjShow.HTMLCR "var datainicio, datafim, dataatual, text, text2;"
  sobjShow.HTMLCR "datainicio = new Date();"
  sobjShow.HTMLCR "datafim    = new Date();"

  sobjShow.HTMLCR "if((document." & strForm & "." & strField & ".value > """")  && (document." & strForm & "." & strField2 & ".value > """") && (document." & strForm & "." & strField3 & ".value > """")){"
  sobjShow.HTMLCR "text       = document." & strForm & "." & strField & ".value + document." & strForm & "." & strField2 & ".value + document." & strForm & "." & strField3 & ".value;"
  sobjShow.HTMLCR "if(!isdate(text)){"
  
  If AllowJSAlert Then _
    sobjShow.HTMLCR "alert(""Data Inválida"");"
  
  sobjShow.HTMLCR "document." & strForm & "." & strField & ".focus();"
  sobjShow.HTMLCR "return false;"
  sobjShow.HTMLCR "}}"

  sobjShow.HTMLCR "if((document." & strForm & "." & strField4 & ".value > """")  && (document." & strForm & "." & strField5 & ".value > """") && (document." & strForm & "." & strField6 & ".value > """")){"
  sobjShow.HTMLCR "text2      = document." & strForm & "." & strField4 & ".value + document." & strForm & "." & strField5 & ".value + document." & strForm & "." & strField6 & ".value;"
  sobjShow.HTMLCR "if(!isdate(text2)){"
  
  If AllowJSAlert Then _
    sobjShow.HTMLCR "alert(""Data Inválida"");"
  
  sobjShow.HTMLCR "document." & strForm & "." & strField4 & ".focus();"
  sobjShow.HTMLCR "return false;"
  sobjShow.HTMLCR "}}"

  sobjShow.HTMLCR "if ((text > """") && (text2 > """")){"
  sobjShow.HTMLCR "datainicio = stringtodate(text);"
  sobjShow.HTMLCR "datafim    = stringtodate(text2);"

  If blnMaiorQueHoje Then
    sobjShow.HTMLCR "dataatual  = new Date();"
    sobjShow.HTMLCR "if (datainicio.valueOf() + 90000000 < dataatual.valueOf()) {"
    
    If AllowJSAlert Then _
      sobjShow.HTMLCR "alert(""" & strMessage2 & """);"
    
    sobjShow.HTMLCR "return false;"
    sobjShow.HTMLCR "}"

  End If

  sobjShow.HTMLCR "if (datainicio.valueOf() > datafim.valueOf()) {"
  
  If AllowJSAlert Then _
    sobjShow.HTMLCR "alert(""" & strMessage & """);"
  
  sobjShow.HTMLCR "document." & strForm & "." & strField & ".focus();"
  sobjShow.HTMLCR "return false;"
  sobjShow.HTMLCR "}}"

  sobjShow.HTMLCR "return true;"
  sobjShow.HTMLCR "}"

End Sub

Public Sub FormJavaCompareDateSep(ByVal strField As String, _
                                  ByVal strField2 As String, _
                                  ByVal strField3 As String, _
                                  ByVal strField4 As String, _
                                  ByVal strField5 As String, _
                                  ByVal strField6 As String, _
                                  ByVal strForm As String, _
                                  ByVal strFunctionName As String, _
                                  ByVal strMessage As String, _
                                  ByVal blnMaiorQueHoje As Boolean, _
                                  ByVal strMessage2 As String)

  CompareSepDateOnClient strField, strField2, strField3, strField4, strField5, strField6, _
                         strForm, strFunctionName, strMessage, blnMaiorQueHoje, strMessage2
End Sub

Rem =========================================================================
Rem Testa se e um CPF
Rem -------------------------------------------------------------------------
Public Function IsCPF(ByVal strCPF As String) As Boolean
  IsCPF = HTIsCPF(strCPF)
End Function

Rem =========================================================================
Rem Testa se e um CGC
Rem -------------------------------------------------------------------------
'Criado (Ridai Govinda)
Public Function CGC(ByVal strCGC As String) As Boolean
  CGC = HTIsCGC(strCGC)

End Function
'---------------------

Public Function IsCGC(ByVal strCGC As String) As Boolean
    IsCGC = HTIsCGC(strCGC)

End Function

Public Function IsCPF_CGC(ByVal strSomething As String, _
                          Optional ByVal bytWhichFirst As Byte = 0) As Boolean
  IsCPF_CGC = HTIsCPF_CGC(strSomething, bytWhichFirst)
End Function

Rem =========================================================================
Rem Testa um Email [Alterado (Ridai)]
Rem -------------------------------------------------------------------------
Public Function IsEmail(ByVal strEmail As String) As Boolean

  Dim blnValid As Boolean
  Dim strDomain As String, strUser As String
  Dim i As Integer, intCount As Integer
  Dim arrAux As Variant

  blnValid = True
  
  If strEmail > "" Then
    blnValid = False

    If Right(strEmail, 1) = "." Then
      strEmail = Left(strEmail, Len(strEmail) - 1)
    End If

    If InStr(strEmail, "@") > 0 Then
      strDomain = Mid(strEmail, InStr(strEmail, "@") + 1, Len(strEmail) - InStr(strEmail, "@") + 1)
      strUser = Left(strEmail, InStr(strEmail, "@"))

      arrAux = Split(strUser, ".")

      intCount = -1

      For i = 0 To UBound(arrAux)
        If arrAux(i) > "" And InStr(arrAux(i), "@") = 0 Then intCount = intCount + 1

      Next

      If intCount = UBound(arrAux) Then blnValid = True

      If blnValid Then

        blnValid = False
        arrAux = Split(strDomain, ".")

        intCount = -1
        For i = 0 To i <= Len(arrAux) - 1

          If arrAux(i) > "" And Not IsNumeric(Left(arrAux(i), 0)) And _
             InStr(arrAux(i), "@") = 0 Then
            intCount = intCount + 1
          End If
        Next
      End If

      If intCount = UBound(arrAux) - 1 Then blnValid = True

    End If

  End If

  IsEmail = blnValid

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
'Public Function IsEmail(ByVal strEmail)
  'IsEmail = ((InStr(strEmail, "@") <> 0) And (InStr(strEmail, ".") <> 0) And (InStr(strEmail, " ") = 0)) Or (Not strEmail > "")

'End Function

Rem =========================================================================
Rem Testa uma Data
Rem -------------------------------------------------------------------------
'Criado (Ridai Govinda)
Public Function IsDate(ByVal strDate As String) As Boolean
  Dim Dia, Mes, Ano, blnIsDate

  blnIsDate = True
  If Not LCase(TypeName(strDate)) = "date" Then
    If strDate > "" Then
      strDate = sobjString.NormalizeString(strDate)
      Mes = Month(Now)
      Ano = Year(Now)
      If Len(strDate) > 1 Then
        Dia = CInt(Left(strDate, 2))
        If Len(strDate) > 2 Then
          Mes = CInt(Mid(strDate, 3, 2))

        End If

        If Len(strDate) > 4 Then
          Ano = CInt(Mid(strDate, 5, 4))

        End If

        If (Mes > 12) Or (Mes < 1) Then
          blnIsDate = False

        End If

        If (Dia < 1) Or (Dia > MonthSize(Mes, Ano)) Then
          blnIsDate = False

        End If

        If (Ano > Year(Now) + 60) Then
          blnIsDate = False

        End If

      Else
        blnIsDate = False

      End If
    End If
  End If

  IsDate = blnIsDate

End Function
'---------------------

Public Function IsValidDate(ByVal strDate As String) As Boolean
  IsValidDate = IsDate(strDate)
End Function

Rem =========================================================================
Rem Testa uma Hora
Rem -------------------------------------------------------------------------
'Criado (Ridai Govinda)
Public Sub IsTimeOnClient(ByVal strField As String, _
                          ByVal strForm As String, _
                          ByVal strFunctionName As String)

End Sub

Rem =========================================================================
Rem Função em JavaScript genérica que valida uma hora no Cliente, o Browser
Rem -------------------------------------------------------------------------
Private Sub ShowFuncIsTime(Optional ByVal strFunctionName As String = "htVldTime")
  
  sstrEmptyFuncName = strFunctionName
  JSBegin
  If Not Functions.Find("IsTime") Then
    sobjShow.HTMLCR "function " & strFunctionName & "(strTime){"
    sobjShow.HTMLCR "var Hora, Minuto;"
    sobjShow.HTMLCR "  if(strTime > ""){"
    sobjShow.HTMLCR "    strTime = normalize(strTime);"
    sobjShow.HTMLCR "    Minuto  = 0;"
    sobjShow.HTMLCR "    Hora    = parseInt(strTime.substr(0, 2));"
    sobjShow.HTMLCR "    if(strTime.length > 2)"
    sobjShow.HTMLCR "      Minuto = parseInt(strTime.substr(2, 2));"
    sobjShow.HTMLCR "    if(Hora > 23)"
    sobjShow.HTMLCR "      return false;"
    sobjShow.HTMLCR "    if(Minuto > 59)"
    sobjShow.HTMLCR "      return false;"
    sobjShow.HTMLCR "  }"
    sobjShow.HTMLCR "  return true;"
    sobjShow.HTMLCR "}"
  End If
  
End Sub

Public Function IsTime(ByVal strTime As String) As Boolean
  Dim Hora, Minuto, blnIsTime

  blnIsTime = True
  If strTime > "" Then
    strTime = sobjString.NormalizeString(strTime)
    Minuto = 0
    Hora = CInt(Left(strTime, 2))
    If Len(strTime) > 2 Then
      Minuto = CInt(Mid(strTime, 3, 2))

    End If

    If (Hora > 23) Then
      blnIsTime = False

    End If

    If (Minuto > 59) Then
      blnIsTime = False

    End If

    strTime = sobjString.LeadingZeroes(Hora, 2) & ":" & sobjString.LeadingZeroes(Minuto, 2)

  End If

Rem  sobjShowJavaScriptMessage strTime

  IsTime = blnIsTime

End Function
'---------------------

Public Function IsValidTime(ByVal strTime As String) As Boolean
  IsValidTime = IsTime(strTime)
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function FormatDate(ByVal strDate As String) As String
Dim Dia, Mes, Ano As Integer

  strDate = sobjString.NormalizeString(strDate)
  Mes = Month(Now)
  Ano = Year(Now)
  If Len(strDate) > 1 Then
    Dia = CInt(Left(strDate, 2))
    If Len(strDate) > 2 Then
      Mes = CInt(Mid(strDate, 3, 2))

    End If

    If Len(strDate) > 4 Then
      Ano = CInt(Mid(strDate, 5, 4))

    End If

    If Ano < 100 Then
      If Ano + 1900 > Year(Now) - 70 Then
        Ano = Ano + 1900

      Else
        Ano = Ano + 2000

      End If
    End If
  End If
  FormatDate = sobjString.LeadingZeroes(Dia, 2) & "/" & sobjString.LeadingZeroes(Mes, 2) & "/" & sobjString.LeadingZeroes(Ano, 4)

End Function

Rem =========================================================================
Rem Formata uma Hora
Rem -------------------------------------------------------------------------
Public Function FormatTime(ByVal strTime As String) As String
Dim Hora, Minuto

  strTime = sobjString.NormalizeString(strTime)
  Minuto = 0
  Hora = CInt(Left(strTime, 2))
  If Len(strTime) > 2 Then
    Minuto = CInt(Mid(strTime, 3, 2))

  End If

  strTime = sobjString.LeadingZeroes(Hora, 2) & ":" & sobjString.LeadingZeroes(Minuto, 2)

  FormatTime = strTime

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function CompareDate(ByVal strDate As String, _
                            ByVal strDate2 As String)
Dim dtmDataInicio, dtmDataFim, intCompare

  intCompare = 0

  If strDate > "" Then
    If IsValidDate(strDate) Then
      dtmDataInicio = CDate(strDate)

    Else
      intCompare = 1

    End If
  End If

  If strDate2 > "" Then
    If IsValidDate(strDate2) Then
      dtmDataFim = CDate(strDate2)

    Else
      intCompare = intCompare + 2

    End If
  End If

  If intCompare = 0 Then
    If strDate > "" And strDate2 > "" Then
      If dtmDataInicio > dtmDataFim Then
        intCompare = 4

      End If
    End If
  End If

  CompareDate = intCompare

End Function

Rem =========================================================================
Rem Retorna o Tamanho de um Mes
Rem -------------------------------------------------------------------------
Public Function MonthSize(ByVal intMes As Integer, _
                          ByVal intAno As Integer)
  Select Case intMes
    Case 1, 3, 5, 7, 8, 10, 12
      MonthSize = 31

    Case 4, 6, 9, 11
      MonthSize = 30

    Case 2
      If intAno Mod 4 = 0 Then
        MonthSize = 29

      Else
        MonthSize = 28

      End If

  End Select

End Function

'Criado (Ridai Govinda)

Public Property Let JSStarted(ByVal blnFlag As Boolean)
  JSStart = blnFlag
End Property

Public Property Get JSStarted() As Boolean
  JSStarted = JSStart
End Property

Public Function IsCEP(ByVal strCEP As String) As Boolean
  Dim blnValid As Boolean

  blnValid = True

  If strCEP > "" Then
    If Not (htMask("00.000-XXX", strCEP) Or htMask("00000-XXX", strCEP)) Then
      If Not htMask("00000XXX", strCEP) Then
        blnValid = False
      End If

    End If
  End If

  IsCEP = blnValid
End Function

Public Function IsAlpha(ByVal strText As String, Optional ByVal blnAccents As Boolean = False) As Boolean
  Dim i As Integer, strTemp As String, strChar As Character
  
  IsAlpha = False

  If blnAccents Then
     strTemp = sobjString.NormalizeAccents(strText)
  Else
     strTemp = strText
  End If
  
  strTemp = UCase(strTemp)
  
  For i = 0 To Len(strTemp)
    strChar = Mid(strTemp, i, 1)
    If Not ((strChar >= "A") And (strChar <= "Z")) Then Exit Function
  Next
  
  IsAlpha = True

End Function


'Depois portar para o objeto String
Private Function htMask(ByVal strMask As String, ByVal strReceive As String) As Boolean
 Dim i As Integer, intCount As Integer
 Dim strRChar As String, strMChar As String

 intCount = 0

 For i = 1 To Len(strMask)
   strMChar = Mid(strMask, i, 1)

   If i <= Len(strReceive) Then
    strRChar = Mid(strReceive, i, 1)
   Else
     strRChar = ""
   End If

   Select Case strMChar
   Case "0"
     If Not IsNumeric(strRChar) Then intCount = intCount + 1
   
   Case "X"
     If strRChar > "" Then
      If Not IsNumeric(strRChar) Then intCount = intCount + 1
     End If
   
   Case "A"
     If Not ((strRChar >= "a" And strRChar <= "z") Or (strRChar >= "A" And strRChar <= "Z")) Then intCount = intCount + 1
   
   Case Else
     If strRChar <> strMChar Then intCount = intCount + 1
   
   End Select

 Next

 If intCount > 0 Then
  htMask = False
 Else
  htMask = True
 End If

End Function

Public Property Let FormName(ByVal strName As String)
  sstrFrmName = strName
End Property

Public Property Get FormName() As String
  FormName = sstrFrmName
End Property
'---------------

Public Function IsValidCreditCardNumber(ByVal sCardNo As String) As Boolean
   'Credit Card Number Validator
   'Must get rid of "-" and " " characters before calling this
   'function

   'alternatively, add Replace(sCardNo, "-", "") and
   'Replace  (sCardNo, " ", "")
   'to the beginning of the function
   
   Const MAX_DIGITS = 20  ' actually don't know any
                            'card using more than 16 digits
   
   Dim anDigits(1 To MAX_DIGITS) As Byte
   Dim nDigits As Long
   
   Dim ofsCurrentDigit As Long
   Dim ofsCurrentCharacter As Long
   
   Dim CurrentCharacter As String
   
   Dim Multiplier As Long
   Dim CheckSum As Long
   Dim DigitValue As Long
   Dim Result As Boolean
   
   Dim ValidDigits As String
   
   If Not IsNumeric(sCardNo) Then Exit Function
   If Len(Trim$(sCardNo)) < 1 Then
      Result = False
      GoTo Exit_Point
   End If
   
   ValidDigits = "0123456789"
   
   For ofsCurrentCharacter = 1 To Len(sCardNo)
      CurrentCharacter = Mid$(sCardNo, ofsCurrentCharacter, 1)
      If InStr(1, ValidDigits, CurrentCharacter, _
           vbBinaryCompare) Then
         nDigits = nDigits + 1
         If nDigits > MAX_DIGITS Then
            Result = False
            GoTo Exit_Point
         End If
         anDigits(nDigits) = Val(CurrentCharacter)
      End If
   Next ofsCurrentCharacter
   
   CheckSum = anDigits(nDigits)
   
   For ofsCurrentDigit = nDigits - 1 To 1 Step -1
      
      If Multiplier = 2 Then
         Multiplier = 1
      Else
         Multiplier = 2
      End If
      DigitValue = anDigits(ofsCurrentDigit) * Multiplier
      CheckSum = CheckSum + DigitValue
      If DigitValue > 9 Then
         CheckSum = CheckSum - 9
      End If
      
   Next ofsCurrentDigit
   
   Result = ((CheckSum Mod 10) = 0)
   
Exit_Point:
   
   IsValidCreditCardNumber = Result
   Exit Function
   
End Function

Public Function IsAlphaBetical(TestString As String) As Boolean
    
    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String
    
    'returns true if all characters in a string are alphabetical
    'returns false otherwise or for empty string
    
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[A-Za-z]" Then Exit Function
        Next
    
    IsAlphaBetical = True
    End If
    
End Function

Public Function IsAlphaNumeric(TestString As String) As Boolean

    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String
    
    'returns true if all characters in a string are alphabetical
    '   or numeric
    'returns false otherwise or for empty string
    
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[0-9A-Za-z]" Then Exit Function
        Next
    
    IsAlphaNumeric = True
    End If
    
End Function
Public Function IsNumericOnly(TestString As String) As Boolean
    
    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String
    
    'returns true if all characters in string are numeric
    'returns false otherwise or for empty string
    
    'this is different than VB's isNumeric
    'isNumeric returns true for something like 90.09
    'This function will return false
    
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[0-9]" Then Exit Function
        Next
    
    IsNumericOnly = True
    End If
    
End Function
