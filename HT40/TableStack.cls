VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TableStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' N�mero de colunas
Public ColumnNumber As Integer

' N�mero de linhas
Public RowNumber As Integer

' Indica se existe uma coluna aberta
Public ColumnOpened As Boolean

' Indica se existe uma linha aberta
Public RowOpened As Boolean

' Indica se existe uma segunda tabela
Public HaveSecondTable As Boolean

' N�mero de colunas da tabela
Public TableColumns As Integer

' Indica se a tabela possui um cabe�alho
Public HeaderRow As Boolean

' Indica se a tabela possui um rodap�
Public LastRow As Integer

' Indica se a tabela possui borda
Public HaveBorder As Boolean

' Indica se a tabela possui um t�tulo(?)
Public HaveTitle As Boolean
