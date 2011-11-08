Attribute VB_Name = "APICalls"
Option Explicit

'// estrutura FileTimeToLocaleTime
Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

'// estrutura FileTimeToSystemTime
Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

' -------------------------------------------------------------------------
' Declaracoes usadas para data/hora
' -------------------------------------------------------------------------
Private Declare Function _
        FileTimeToLocalFileTime Lib "kernel32" _
        (lptFileTime As FILETIME, lptLocalFileTime As FILETIME) _
        As Long
        
Private Declare Function _
        FileTimeToSystemTime Lib "kernel32" _
        (lptFileTime As FILETIME, lptSystemTime As SYSTEMTIME) _
        As Long

Private Declare Function _
        GetDateFormat Lib "kernel32" Alias "GetDateFormatA" _
        (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, _
        ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) _
        As Long

' -------------------------------------------------------------------------
' Constante e declaracao usados em APIErrorText
' -------------------------------------------------------------------------
Private Declare Function FormatMessage _
   Lib "kernel32" Alias "FormatMessageA" _
      (ByVal dwFlags As Long, lpSource As Any, _
       ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
       ByVal lpBuffer As String, ByVal nSize As Long, _
       Arguments As Long) As Long

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000


' -------------------------------------------------------------------------
' Funcao que retorna a descricao das mensagens de erro da API Win32
' -------------------------------------------------------------------------
Public Function APIErrorText(ByVal lngErrNum As Long) As String
   Dim strMsg As String, _
        lngReturn As Long

   strMsg = Space$(1024)
   lngReturn = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, _
          ByVal 0&, lngErrNum, 0&, strMsg, Len(strMsg), ByVal 0&)
   If lngReturn Then
       APIErrorText = Left$(strMsg, lngReturn)
   Else
       APIErrorText = "Erro (" & lngErrNum & ") não definido."
   End If

End Function

Public Function FileTimeToDate(ByRef stFileTime As FILETIME) As Date
    Dim stSystemTime As SYSTEMTIME
    
    FileTimeToSystemTime stFileTime, stSystemTime
    
    FileTimeToDate = DateValue(Format(stSystemTime.wMonth, "00") & "/" & _
               Format(stSystemTime.wDay, "00") & "/" & _
               Format(stSystemTime.wYear, "0000")) + _
               TimeValue(Format(stSystemTime.wHour, "00") & ":" & _
               Format(stSystemTime.wMinute, "00") & ":" & _
               Format(stSystemTime.wSecond, "00"))
    
End Function

'Criado (Ridai)
Rem =========================================================================
Rem Levanta uma excecao
Rem -------------------------------------------------------------------------
Public Sub ShowError(ByVal lngError As Long, ByVal strObject As String, _
                     ByVal strCaller As String, Optional ByVal strDescription As String = "")
    Err.Raise vbObjectError + lngError, _
        App.Title & "." & strObject & "." & strCaller, _
        IIf(Not strDescription > "", APIErrorText(lngError), strDescription)
        
End Sub
