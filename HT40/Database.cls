VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =========================================================================
'  /HiperTools/Database.inc
' -------------------------------------------------------------------------
'  Descricao   : Include de Databases
'  Cria‡ao     : 11:23h 7/2/1998
'  Local       : Brasilia/DF
'  Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
'  Versao      : 1.0.0
'  Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------
'  ALTERACOES
' -------------------------------------------------------------------------
'  Responsavel : [Nome do executante da alteracao]
'  Data/Hora   : [Data e hora da alteracao]
'  Resumo      : [Resumo descritivo da alteracao executada]
' -------------------------------------------------------------------------

Private sblnPut As Boolean
Private ScriptContext As ScriptingContext

Private Enum DatabaseErrors
  dbErrorIdentityNoConn = 102
End Enum

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
    
Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  TimeBombX xtbDatabase, 15656, 12934, 3080, 14624
  
  Set ScriptContext = ScriptContext
  
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application
  
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

Public Sub CloseObject(ByRef obj As Object)
  If obj.State = 1 Then
    obj.Close

  End If

End Sub

Private Sub Class_Initialize()

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
Public Property Let AndState(ByVal blnValor As Boolean)
  sblnPut = blnValor
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get AndState() As Boolean
  AndState = sblnPut
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let OrState(ByVal blnValor As Boolean)
  sblnPut = blnValor
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get OrState() As Boolean
  OrState = sblnPut
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function PutAnd() As String

  If AndState Then
    PutAnd = " AND "
  Else
    PutAnd = " WHERE "
  End If

  AndState = True
  
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function PutOr() As String
  
  If OrState Then
    PutOr = " OR "
  Else
    PutOr = " WHERE "
  End If

  OrState = True
  
End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub ListFieldsSQL(ByVal objConnListFields As ADODB.Connection, _
                         ByVal strSQL As String)
  
  Dim objRSListFields
  
  Set objRSListFields = objConnListFields.Execute(strSQL)
  
  ListFields (objRSListFields)
  
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub ListFields(ByVal objRSSL As ADODB.RecordSet)
  Dim i
  
  On Error Resume Next
  
  sobjShow.HTMLCR "<table border=1>"
  sobjShow.HTMLCR "<tr bgcolor=lightblue><td colspan=4>Fields: " & objRSSL.Fields.Count & "</td></tr>"
  sobjShow.HTMLCR "<tr bgcolor=lightblue><td colspan=4>Records: " & objRSSL.RecordCount & "</td></tr>"

  sobjShow.HTMLCR "<tr bgcolor=lightblue><td>Id</td><td>Name</td><td>Type</td><td>Value</td></tr>"

  On Error GoTo 0

  For i = 0 To objRSSL.Fields.Count - 1
    If i Mod 2 = 0 Then
      Response.Write "<tr bgcolor=white>"
      
    Else
      Response.Write "<tr bgcolor=lightyellow>"
      
    End If
    
    sobjShow.HTMLCR "<td>" & i & _
               "</td><td>" & objRSSL(i).Name & _
               "</td><td>" & objRSSL(i).Type & _
               "</td><td>" & objRSSL(i).Value & _
               "</td></tr>"
    
  Next
  
  sobjShow.HTMLCR "</table>"

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub ListFieldsCurrent(ByVal objRSSL As ADODB.RecordSet, _
                             ByVal intID As Integer)

    sobjShow.Message "Id=" & intID & _
                " - Name=" & objRSSL(intID).Name & _
                " - Type=" & objRSSL(intID).Type & _
                " - Value=" & objRSSL(intID).Value

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub Lista(ByVal objRS As ADODB.RecordSet)
  Dim x, i As Integer
  
    sobjShow.HTMLCR "<p>"
    sobjShow.HTMLCR "<table border=1 cellpadding=0 cellspacing=0>"
    sobjShow.HTMLCR "<tr><td><b>Reg</b>"

    For i = 0 To objRS.Fields.Count - 1
      Response.Write "<td><b>" & objRS(i).Name & "</b>"
      
    Next

    objRS.MoveFirst

    x = 0

    Do While Not objRS.EOF
      x = x + 1
      
      sobjShow.HTMLCR "<TR><TD VALIGN=TOP>" & x
      
      For i = 0 To objRS.Fields.Count - 1
        Response.Write "<TD VALIGN=TOP>" & objRS(i)
        
      Next

      objRS.MoveNext

    Loop

    objRS.Close
    
    sobjShow.HTMLCR "</table>"

    sobjShow.Message "Registros Listados!!"

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub RSFieldLet(ByVal objConn As ADODB.Connection, _
                      ByVal objRS As ADODB.RecordSet, _
                      ByVal strField As String, _
                      ByVal strResult As String)
  On Error Resume Next
  
  objRS(strField) = strResult
  
  If ErrorConnHandler(objConn) <> 0 Then
    If objRS.EditMode Then
      objRS.CancelUpdate
      
    End If

    sobjShow.MessageError "Field=" & strField & " - Result='" & strResult & "' - Type=" & VarType(strResult) & " (" & TypeName(strResult) & ") - Lenght=" & Len(strResult)

    Err.Clear

    On Error GoTo 0

    objRS(strField) = strResult

  End If

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub RSUpdate(ByVal objConn As ADODB.Connection, _
                    ByVal objRS As ADODB.RecordSet)
  On Error Resume Next
  
  objRS.Update
  
  If ErrorConnHandler(objConn) <> 0 Then
    If objRS.EditMode Then
      objRS.CancelUpdate
      
    End If

    Err.Clear

    On Error GoTo 0

    objRS.Update

  End If


End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub RSOpen(ByVal objConn As ADODB.Connection, _
                  ByVal objRS As ADODB.RecordSet, _
                  ByVal strSQL As String, _
                  ByVal strPar1 As String, _
                  ByVal strPar2 As String)
  
  On Error Resume Next
  
  objRS.Open strSQL, objConn, strPar1, strPar1
  
  If ErrorConnHandler(objConn) <> 0 Then
    sobjShow.ItemListBegin
    sobjShow.ItemListError "SQL", strSQL
    sobjShow.ItemListError "objConn", objConn
    sobjShow.ItemListError "Parametro1", strPar1
    sobjShow.ItemListError "Parametro2", strPar2
    sobjShow.ItemListEnd
    
    Err.Clear

    On Error GoTo 0

    objRS.Open strSQL, objConn, strPar1, strPar1

  End If

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub RSOpenCommand(ByRef objRS As ADODB.RecordSet, _
                         ByVal objCMD As ADODB.Command, _
                         ByVal objConn As ADODB.Connection, _
                         ByVal strPar1 As String, _
                         ByVal strPar2 As String)
  On Error Resume Next
  
  objRS.Open objCMD, , strPar1, strPar1

  If ErrorConnHandler(objConn) <> 0 Then
    sobjShow.ItemListBegin
    sobjShow.ItemListError "objConn", objConn
    sobjShow.ItemListError "objCMD.CommandText", objCMD.CommandText
    sobjShow.ItemListError "Parametro1", strPar1
    sobjShow.ItemListError "Parametro2", strPar2
    sobjShow.ItemListEnd
    
    Err.Clear

    On Error GoTo 0

    objRS.Open objCMD, , strPar1, strPar1

  End If
  
End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub RSConnExecute(ByVal objConn As ADODB.Connection, _
                         ByVal strSQL As String, _
                         ByRef objRS As ADODB.RecordSet)
  On Error Resume Next
  
  Set objRS = objConn.Execute(strSQL)
  
  If ErrorConnHandler(objConn) <> 0 Then
    sobjShow.ItemListBegin
    sobjShow.ItemListError "SQL", strSQL
    sobjShow.ItemListError "objConn", objConn
    sobjShow.ItemListEnd
    
    Err.Clear

    On Error GoTo 0

    Set objRS = objConn.Execute(strSQL)

  End If

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub ConnExecute(ByVal objConn As ADODB.Connection, _
                       ByVal strSQL As String)
  On Error Resume Next
  
  objConn.Execute (strSQL)

  If ErrorConnHandler(objConn) <> 0 Then
    sobjShow.ItemListBegin
    sobjShow.ItemListError "SQL", strSQL
    sobjShow.ItemListError "objConn", objConn
    sobjShow.ItemListEnd
    
    Err.Clear

    On Error GoTo 0

    objConn.Execute (strSQL)

  End If

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub CreateTable(ByVal strSQL As String, _
                       ByVal strTableName As String)

  Session("CreateTableTableName") = strTableName
  
  On Error Resume Next

  ConnExecute sobjConn, "DROP TABLE " & strTableName

  On Error GoTo 0

  strSQL = "CREATE TABLE " & strTableName & " (" & strSQL & ")"

  ConnExecute sobjConn, strSQL

  sobjShow.Message "Criada a tabela " & strTableName

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub CreateIndex(ByVal strIndexName As String, _
                       ByVal strFields As String)

  On Error Resume Next

Rem  ConnExecute sobjConn, "DROP INDEX " & strIndexName & "_" &  Session("CreateTableTableName")

  On Error GoTo 0

  ConnExecute sobjConn, "CREATE INDEX " & strIndexName & "_" & Session("CreateTableTableName") & " ON " & _
              Session("CreateTableTableName") & "(" & strFields & ")"

  sobjShow.Message "Criado o Indice " & strIndexName & "_" & Session("CreateTableTableName")

End Sub

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Sub CreateIndexClustered(ByVal strIndexName As String, _
                                ByVal strFields As String)

  On Error Resume Next

  ConnExecute sobjConn, "DROP TABLE " & strTableName

  On Error GoTo 0

  ConnExecute sobjConn, "CREATE CLUSTERED INDEX " & strIndexName & "_" & Session("CreateTableTableName") & " ON " & _
              Session("CreateTableTableName") & "(" & strFields & ")"

  sobjShow.Message "Criado o Indice " & strIndexName & "_" & Session("CreateTableTableName")

End Sub
'Criado (Ridai)
Public Function ADOLock(ByVal intLock As Integer) As String
  
  Select Case intLock
  Case adLockReadOnly
    ADOLock = "adLockReadOnly"
  Case adLockPessimistic
    ADOLock = "adLockPessimistic"
  Case adLockOptimistic
    ADOLock = "adLockOptimistic"
  Case adLockBatchOptimistic
    ADOLock = "adLockBatchOptimistic"
  End Select

End Function

Public Function ADOOpenType(ByVal intOpen As Integer) As String
  
  Select Case intOpen
  Case adOpenKeyset
    ADOOpenType = "adOpenKeyset"
  Case adOpenForwardOnly
    ADOOpenType = "adOpenForwardOnly"
  Case adOpenDynamic
    ADOOpenType = "adOpenDynamic"
  Case adOpenStatic
    ADOOpenType = "adOpenStatic"
  End Select

End Function
'----

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function ErrorHandler() As Variant

  If Err.Number > 0 Then
    sobjShow.MessageError "Erro com o VBScript:"

    sobjShow.ItemListBegin
    sobjShow.ItemListError "Numero", Err.Number
    sobjShow.ItemListError "Descri&ccedil;&atilde;o.", Err.Description
    sobjShow.ItemListError "Help Context", Err.HelpContext
    sobjShow.ItemListError "Help Path", Err.HelpPath
    sobjShow.ItemListError "Native Error", Err.NativeError
    sobjShow.ItemListError "Source", Err.Source
    sobjShow.ItemListError "SQLState", Err.SQLState
    sobjShow.ItemListEnd

  End If
  
  ErrorHandler = Err.Number

End Function

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Function ErrorConnHandler(ByVal objConn As ADODB.Connection) As Variant
  Dim i
  
  On Error GoTo 0
  
  ErrorConnHandler = ErrorHandler
  
  If objConn.Errors.Count > 0 Then
    sobjShow.MessageError objConn.Errors.Count & " - Erro(s) com a Base de dados:"

    sobjShow.ItemListBegin

    For i = 0 To objConn.Errors.Count
      sobjShow.ItemListError i + 1 & " - N&uacute;mero", objConn.Errors(i).Number
      sobjShow.ItemListError "Descri&ccedil;&atilde;o", objConn.Errors(i).Description
     
    Next
    
    On Error Resume Next

    sobjShow.ItemListError "Source", objConn.Source
    sobjShow.ItemListError "SQLState", objConn.SQLState
    
    sobjShow.ItemListEnd

  End If

End Function

Public Function GetIdentity(ByVal strTable As String, ByVal strField As String, _
                            Optional objConn As ADODB.Connection) As Long
  Dim strSQL As String
  Dim lngIdentity As Long
  Dim objRS As New ADODB.RecordSet
  
  strSQL = ""
  
  If objConn Is Nothing Then Set objConn = New ADODB.Connection
  
  If Not objConn.State Then
    If Session("ConnectionString") > "" Then
      objConn.Open Session("ConnectionString")
    Else
      ShowError DatabaseErrors.dbErrorIdentityNoConn, "Database", "GetIdentity", "Não há conexão com o banco de dados ativa."
    End If
  End If
  
  Application.Lock
  
  If IsEmpty(Application("htDb" & strTable & "." & strField)) Then
    
    strSQL = "SELECT MAX(" & strField
    strSQL = strSQL & ") AS Codigo FROM " & strTable
    
    objRS.Open strSQL, objConn, adOpenDynamic, adLockPessimistic
    
    If Not IsNull(objRS("Codigo")) Then
      lngIdentity = CLng(objRS("Codigo")) + 1
    
    Else
      lngIdentity = 1
    
    End If
    objRS.Close
  
  Else
    lngIdentity = Application("htDb" & strTable & "." & strField)
    lngIdentity = lngIdentity + 1
    
  End If

  Application("htDb" & strTable & "." & strField) = lngIdentity
  Application.Unlock
  
  GetIdentity = lngIdentity

End Function
