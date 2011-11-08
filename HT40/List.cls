VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------------------------------------------
' List.cls
'-------------------------------------------------------------------------
' Descricao   : Sistema de Listagem (tirado e componentizado do EditForm)
' Criacao     : 11:30h 22/02/1999
' Local       : Brasilia/DF
' Elaborado   : Ridai Govinda Pombo <ridai@zevallos.com.br>
'               Kelson Medeiros <kelson@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 2000 by Zevallos(r) Tecnologia em Informacao
'-------------------------------------------------------------------------
Option Explicit

'Propriedades simples:
Public AlphabeticIndex As Boolean

