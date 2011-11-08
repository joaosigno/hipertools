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

' Número de colunas
Public ColumnNumber As Integer

' Número de linhas
Public RowNumber As Integer

' Indica se existe uma coluna aberta
Public ColumnOpened As Boolean

' Indica se existe uma linha aberta
Public RowOpened As Boolean

' Indica se existe uma segunda tabela
Public HaveSecondTable As Boolean

' Número de colunas da tabela
Public TableColumns As Integer

' Indica se a tabela possui um cabeçalho
Public HeaderRow As Boolean

' Indica se a tabela possui um rodapé
Public LastRow As Integer

' Indica se a tabela possui borda
Public HaveBorder As Boolean

' Indica se a tabela possui um título(?)
Public HaveTitle As Boolean
