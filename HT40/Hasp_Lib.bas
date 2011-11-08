Attribute VB_Name = "Hasp_Lib"
'-------------------------------------------------------------------------
' Hasp_Lib.bas
'-------------------------------------------------------------------------
' Descricao   : Definições de constantes e procedimento públicos
'               para a classe _Hasp_
' Criaçao     : 16/01/2000 18:55
' Local       : Brasilia/DF
' Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
'               Ridai Govinda Pombo <ridai@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 97-2000 by Zevallos(r) Tecnologia em Informacao
'-------------------------------------------------------------------------

Option Explicit

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

Rem
Rem Use this declaration to call the hasp() routine directly.
Rem

Public Declare Sub HaspLib Lib "haspvb32.dll" Alias "hasp" (ByVal Service&, ByVal seed&, ByVal LPT&, ByVal pass1&, ByVal pass2&, ByRef retcode1&, ByRef retcode2&, ByRef retcode3&, ByRef retcode4&)
Public Declare Sub WriteHaspBlock Lib "haspvb32.dll" (ByVal Service&, Buff As Any, ByVal Length&)
Public Declare Sub ReadHaspBlock Lib "haspvb32.dll" (ByVal Service&, Buff As Any, ByVal Length&)
