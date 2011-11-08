VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Connection"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =========================================================================
' Connection.cls
' -------------------------------------------------------------------------
' Descricao   : Objeto da Gerencia os Connections
' Criacao     : 1/30/99 2:23AM
' Local       : Brasilia/DF
' Elaborado   : Ruben Zevallos Jr. - <ruben@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 1997, 1998, 1999 by Zevallos(r) Tecnologia em Informacao
'-------------------------------------------------------------------------

'    Description = "HiperTools.Connection - Dicionario de Dados - 1/31/99 10:47AM"
'    ProgId = "HiperTools.Connection"
'    Version = "1.01"
'    ClassID = "{E0408FF0-B920-11d2-B163-00A0C923ECAE}"

Option Explicit
Rem =========================================================================
Rem Variaveis Globais
Rem -------------------------------------------------------------------------
Private sintConnectionType As Integer
Private sintDriverID As Integer
Private sstrDBQ As String
Private sstrServerAddress As String
Private sstrDataBase As String
Private sstrUserID As String
Private sstrPassword As String
Private sstrDefaultDir As String
Private sintConnectionTimeout As Integer
Private sintCommandTimeout As Integer

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

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  TimeBombX xtbConnection, 15656, 12934, 3080, 14624
  
  Set ScriptContext = ScriptContext
  
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application
  
End Sub

Public Sub OnEndPage()

  Set ScriptContext = Nothing
  
  Set Request = Nothing
  Set Response = Nothing
  Set Server = Nothing
  Set Session = Nothing
  Set Application = Nothing
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

Rem -------------------------------------------------------------------------
Rem Propriedade ConnectionType
Rem -------------------------------------------------------------------------
Public Property Get ConnectionType() As String
    ConnectionType = sintConnectionType
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let ConnectionType(ByRef strValue As String)
    sintConnectionType = strValue
End Property

Rem -------------------------------------------------------------------------
Rem Propriedade DriverID
Rem -------------------------------------------------------------------------
Public Property Get DriverID() As String
    DriverID = sintDriverID
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let DriverID(ByRef strValue As String)
    sintDriverID = strValue
End Property

Rem -------------------------------------------------------------------------
Rem Propriedade DBQ
Rem -------------------------------------------------------------------------
Public Property Get DBQ() As String
    DBQ = sstrDBQ
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let DBQ(ByRef strValue As String)
    sstrDBQ = strValue
End Property

Rem -------------------------------------------------------------------------
Rem Propriedade ServerAddress
Rem -------------------------------------------------------------------------
Public Property Get ServerAddress() As String
    ServerAddress = sstrServerAddress
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let ServerAddress(ByRef strValue As String)
    sstrServerAddress = strValue
End Property

Rem -------------------------------------------------------------------------
Rem Propriedade DataBase
Rem -------------------------------------------------------------------------
Public Property Get Database() As String
    Database = sstrDataBase
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Database(ByRef strValue As String)
    sstrDataBase = strValue
End Property

Rem -------------------------------------------------------------------------
Rem Propriedade UserId
Rem -------------------------------------------------------------------------
Public Property Get UserID() As String
    UserID = sstrUserID
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let UserID(ByRef strValue As String)
    sstrUserID = strValue
End Property

Rem -------------------------------------------------------------------------
Rem Propriedade Password
Rem -------------------------------------------------------------------------
Public Property Get Password() As String
Attribute Password.VB_Description = "Let or Get the Password"
    Password = sstrPassword
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let Password(ByRef strValue As String)
    sstrPassword = strValue
End Property

Rem -------------------------------------------------------------------------
Rem Propriedade ConnectionTimeout
Rem -------------------------------------------------------------------------
Public Property Get ConnectionTimeout() As String
    ConnectionTimeout = sintConnectionTimeout
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let ConnectionTimeout(ByRef strValue As String)
    sintConnectionTimeout = strValue
End Property

Rem -------------------------------------------------------------------------
Rem Propriedade CommandTimeout
Rem -------------------------------------------------------------------------
Public Property Get CommandTimeout() As String
Attribute CommandTimeout.VB_Description = "Get or Set the Command Timeout"
    CommandTimeout = sintCommandTimeout
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let CommandTimeout(ByRef strValue As String)
    sintCommandTimeout = strValue
End Property

Rem -------------------------------------------------------------------------
Rem Metodo Clear
Rem -------------------------------------------------------------------------
Public Sub Clear()
Attribute Clear.VB_Description = "Clear all properties"
  sintConnectionType = 0
  sintDriverID = 0
  sstrDBQ = ""
  sstrServerAddress = ""
  sstrDataBase = ""
  sstrUserID = ""
  sstrPassword = ""
  sstrDefaultDir = ""
End Sub

Rem -------------------------------------------------------------------------
Rem Metodo da Conexao
Rem -------------------------------------------------------------------------
Public Function ConnectionString() As String
Attribute ConnectionString.VB_Description = "Return the Connection String"
Attribute ConnectionString.VB_UserMemId = 0
  
  Dim strCS

  strCS = ""
  
  Select Case sintConnectionType
    Case conConnSQL
      strCS = "{SQL Server}"
                            
    Case conConnAccess
      strCS = "{Microsoft Access Driver (*.mdb)}"

    Case conConnDBase
      strCS = "{Microsoft dBASE Driver (*.dbf)}"
 
    Case conConnFoxPro
      strCS = "{Microsoft FoxPro Driver (*.dbf)}"

    Case conConnParadox
      strCS = "{Microsoft Paradox Driver (*.db )}"

    Case conConnText
      strCS = "{Microsoft Text Driver (*.txt;*.csv)}"
                            
    Case conConnExcel
      strCS = "{Microsoft Excel Driver (*.xls)}"

  End Select

  If strCS > "" Then
    strCS = "DRIVER=" & strCS
  
    If sstrDBQ > "" Then
      strCS = strCS & ";DBQ=" & sstrDBQ
  
    End If
  
    If sstrServerAddress > "" Then
      strCS = strCS & ";SERVER=" & sstrServerAddress
  
    End If

    If sstrUserID > "" Then
      strCS = strCS & ";UID=" & sstrUserID
  
    End If
  
    If sstrPassword > "" Then
      strCS = strCS & ";PWD=" & sstrPassword
  
    End If

    If sintDriverID > 0 Then
      strCS = strCS & ";DriverID=" & sintDriverID
  
    End If
  
    If sstrDefaultDir > "" Then
      strCS = strCS & ";DefaultDir=" & sstrDefaultDir
  
    End If

    If sstrDataBase > "" Then
      strCS = strCS & ";DATABASE=" & sstrDataBase
  
    End If
  
    strCS = strCS & ";"
  End If

  ConnectionString = strCS

End Function

Rem =========================================================================
Rem Mantidas apenas por questão de compatibilidade
Rem ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Public Function Connection() As String 'Obsoleto
  Connection = ConnectionString
End Function

Public Property Get DefaultDir() As String 'Obsoleto
    DefaultDir = sstrDefaultDir
End Property

Public Property Let DefaultDir(ByRef strValue As String) 'Obsoleto
    sstrDefaultDir = strValue
End Property
