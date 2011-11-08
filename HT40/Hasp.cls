VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Hasp"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------------------------------------------
' Hasp.cls
'-------------------------------------------------------------------------
' Descricao   : HASP - HardLock da Aladdin
' Criaçao     : 1/10/0 12:55AM
' Local       : Brasilia/DF
' Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
'               Ridai Govinda Pombo <ridai@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 97-2000 by Zevallos(r) Tecnologia em Informacao
'-------------------------------------------------------------------------

Option Explicit

Public RC1
Public RC2
Public RC3
Public RC4

Public SeedCode
Public LptNum
Public Password1
Public Password2
Public ZYC
Public Query

Public P1 As Long
Public P2 As Long
Public P3 As Long
Public P4 As Long


' The HASP memory buffer.
'
Private Type HBuff
    txt As String * 500
End Type

'
' The TimeHASP memory buffer.
'
Private Type HTimeBuff
    txt As String * 20
End Type

Private MemoHaspBuffer As HBuff
Private TimeHaspBuffer As HTimeBuff

Private Sub Class_Initialize()
  LptNum = LPT_IBM_ALL_HASP25

End Sub

Public Function IsHasp() As Boolean

  On Error GoTo IsHaspError
    Call HaspLib(IS_HASP, SeedCode, LptNum, Password1, Password2, P1, P2, P3, P4)
    
  On Error GoTo 0
    
    IsHasp = (P1 <> 0)
    
  Exit Function
  
IsHaspError:

  Err.Raise 300, "HiperTools30.Hasp", "O HASP não esta instalado!"

End Function

Public Function IsMemoHASP() As Boolean
  
  Call HaspLib(GET_HASP_STATUS, SeedCode, LptNum, Password1, Password2, P1, P2, P3, P4)
  
  IsMemoHASP = (P2 = 1)

End Function

Public Function HaspID() As Long
    
    Call HaspLib(GET_ID_NUM, SeedCode, LptNum, Password1, Password2, P1, P2, P3, P4)
    
    HaspID = P1 + 65536 * P2
    
End Function

Public Function ReadBlock(ByVal intSize As Long) As String
  
  ReadBlock = vbNullString
  If IsHasp() And IsMemoHASP() Then
    P2 = intSize
    
    Call HaspLib(READ_MEMO_BLOCK, SeedCode, LptNum, Password1, Password2, P1, P2, P3, P4)
    
    If P3 = 0 Then ReadBlock = MemoHaspBuffer.txt
  End If
    
End Function

Public Function ReadWord(ByVal intAddress As Long) As Long
  
  ReadWord = vbNullString
  If IsHasp() And IsMemoHASP() Then
    P2 = -1
    Call HaspLib(READ_WORD, SeedCode, LptNum, Password1, Password2, intAddress, P2, P3, P4)
    
    If P3 = 0 Then ReadWord = P2
  End If
  
End Function

Public Function WriteBlock(ByVal strBlock As String) As Boolean
  
  If IsHasp() And IsMemoHASP() Then
    MemoHaspBuffer.txt = Left(strBlock, 500)
    P2 = Len(MemoHaspBuffer.txt)
    
    Call WriteHaspBlock(WRITE_MEMO_BLOCK, MemoHaspBuffer, P2)
    Call HaspLib(WRITE_MEMO_BLOCK, SeedCode, LptNum, Password1, Password2, P1, P2, P3, P4)
    
    WriteBlock = (P3 = 0)
  Else
    WriteBlock = False
  End If
  
End Function

Private Function WriteWord(ByVal intValue As Long, ByVal intAddress As Long) As Boolean
    If IsHasp() And IsMemoHASP() Then
      Call HaspLib(WRITE_WORD, SeedCode, LptNum, Password1, Password2, intAddress, intValue, P3, P4)
      WriteWord = (P3 = 0)
    
    Else
      WriteWord = False
    End If

End Function

Public Function MemorySize() As Integer
    
    Call HaspLib(GET_HASP_STATUS, SeedCode, LptNum, Password1, Password2, P1, P2, P3, P4)
    
    Select Case P1
        Case 1
            MemorySize = 55
        Case 4
            MemorySize = 247
        Case Else
            MemorySize = 0
            
    End Select
    
End Function

Public Function HaspType() As String
  
  Call HaspLib(GET_HASP_STATUS, SeedCode, LptNum, Password1, Password2, P1, P2, P3, P4)
  
  Select Case P2
  Case 0
      HaspType = "HASP-3"
  Case 1
      HaspType = "MemoHASP-1 Or MemoHASP-4"
  Case 3
      HaspType = "TimeHASP"
  Case 5
      HaspType = "TimeHASP-4"
  Case Else
      HaspType = ""

  End Select
    
End Function

Public Function LPT() As String
    
  Call HaspLib(GET_HASP_STATUS, SeedCode, LptNum, Password1, Password2, P1, P2, P3, P4)
  
  LPT = P3
    
End Function

Public Sub HaspCode(ByVal ZSeedCode As Long)
  Dim ZP1, ZP2, ZP3, ZP4
  
  If ZYC > 0 And Query > 0 Then
     Call HaspLib(GET_HASP_CODE, ZSeedCode, LptNum, ZYC - ZYC_SEED, Query + Query_SEED, ZP1, ZP2, ZP3, ZP4)
  Else
    Call HaspLib(GET_HASP_CODE, ZSeedCode, LptNum, Password1, Password2, ZP1, ZP2, ZP3, ZP4)
  
  End If
  
  RC1 = ZP1
  RC2 = ZP2
  RC3 = ZP3
  RC4 = ZP4
    
End Sub


