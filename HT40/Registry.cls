VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Objeto de acesso ao registro do Windows."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-------------------------------------------------------------------------
' Registry.cls
'-------------------------------------------------------------------------
' Descricao   : Include de acesso ao registro do Windows
' Criacao     : 09:30h 04/02/2000
' Local       : Brasilia/DF
' Elaborado   : Flavio Costa <flavio@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 2000 by Zevallos(r) Tecnologia em Informacao
'-------------------------------------------------------------------------
  
Option Explicit
  
'// Valores para Open/Create
Private Const REG_OPTION_RESERVED = 0&     '// Parameter is reserved
Private Const REG_OPTION_NON_VOLATILE = 0& '// Key is preserved when system is rebooted
Private Const REG_OPTION_VOLATILE = 1&     '// Key is not preserved when system is rebooted

'// tipos de valores
Private Const REG_NONE = 0                 '// No value type
Private Const REG_SZ = 1                   '// Unicode null terminated string
Private Const REG_EXPAND_SZ = 2            '// Unicode null terminated string with %environment% references
Private Const REG_BINARY = 3               '// Free form binary
Private Const REG_DWORD = 4                '// 32-bit number
Private Const REG_DWORD_LITTLE_ENDIAN = 4  '// 32-bit number (same as REG_DWORD)
Private Const REG_DWORD_BIG_ENDIAN = 5     '// 32-bit number
Private Const REG_LINK = 6                 '// Symbolic Link (unicode)
Private Const REG_MULTI_SZ = 7             '// Multiple Unicode strings
Private Const REG_RESOURCE_LIST = 8        '// Resource list in the resource map
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9 '// Resource list in the hardware description
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10 '// Resource requirements list

'// Registry hKey values
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006

'//return values
Private Const ERROR_SUCCESS = 0& '// The operation completed successfully.
Private Const ERROR_FILE_NOT_FOUND = 2& '// The system cannot find the file specified.
Private Const ERROR_ACCESS_DENIED = 5& '// Access is denied.
Private Const ERROR_MORE_DATA = 234& '// More data is available.
Private Const ERROR_NO_MORE_ITEMS = 259& '// No more data is available.
Private Const ERROR_BADDB = 1009& '// The configuration registry database is corrupt.
Private Const ERROR_BADKEY = 1010& '// The configuration registry key is invalid.

'// Registry Access values
Private Const DELETE = &H10000
Private Const READ_CONTROL = &H20000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20

Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or _
                        KEY_QUERY_VALUE Or _
                        KEY_ENUMERATE_SUB_KEYS Or _
                        KEY_NOTIFY) _
                         And (Not SYNCHRONIZE))

Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or _
                         KEY_SET_VALUE Or _
                         KEY_CREATE_SUB_KEY) _
                         And (Not SYNCHRONIZE))
                         
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
                           
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or _
                              KEY_QUERY_VALUE Or _
                              KEY_SET_VALUE Or _
                              KEY_CREATE_SUB_KEY Or _
                              KEY_ENUMERATE_SUB_KEYS Or _
                              KEY_NOTIFY Or _
                              KEY_CREATE_LINK) _
                              And (Not SYNCHRONIZE))
                              

'//comprimento máximo dos valores
Private Const MAX_VALUENAME_LEN = 256
Private Const MAX_REG_STRING_LEN = 256
Private Const MAX_REG_BINARY_LEN = 256
Private Const MAX_REG_DWORD_LEN = 4
Private Const MAX_DATE_LEN = 256

'//importação das funções da API Win32
Private Declare Function _
        ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" _
        (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) _
        As Long
        
Private Declare Function _
        RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long

Private Declare Function _
        RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" _
        (ByVal lpMachineName As String, _
        ByVal hKey As Long, phkResult As Long) _
        As Long
        
Private Declare Function _
        RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
        ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
        ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) _
        As Long

Private Declare Function _
        RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
        (ByVal hKey As Long, ByVal lpSubKey As String) _
        As Long
        
Private Declare Function _
        RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
        (ByVal hKey As Long, ByVal lpValueName As String) _
        As Long
        
Private Declare Function _
        RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
        (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
        lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, _
        lpcbClass As Long, lpftLastWriteTime As FILETIME) _
        As Long

Private Declare Function _
        RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
        (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
        lpcbValueName As Long, ByVal lpReserved As Long, _
        lpType As Long, lpData As Byte, lpcbData As Long) _
        As Long
        
Private Declare Function _
        RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
        ByVal samDesired As Long, phKeyResult As Long) _
        As Long
                             
Private Declare Function _
        RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
        (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, _
        ByVal lpReserved As Long, lpcSubKeys As Long, lpMaxSubKeyLen As Long, _
        lpMaxClassLen As Long, lpcValues As Long, _
        lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, _
        lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) _
        As Long
       
Private Declare Function _
        RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As Byte, lpcbData As Long) As Long

Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
        "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
        String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
        As String, lpcbData As Long) As Long
   
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
        "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
        String, ByVal lpReserved As Long, lpType As Long, lpData As _
        Long, lpcbData As Long) As Long
   
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
        ByVal dwType As Long, lpData As Byte, ByVal cbData As Long) _
        As Long

Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
        "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
        String, ByVal cbData As Long) As Long
   
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
        "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
        ByVal cbData As Long) As Long

Private Declare Function _
        WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" _
        (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Boolean) _
        As Long

Rem =========================================================================
Rem variaveis publicas que implementam propriedades
Rem -------------------------------------------------------------------------

Rem =========================================================================
Rem variaveis locais que implementam propriedades
Rem -------------------------------------------------------------------------
Private lngHKEY As Long 'ponteiro para HKEY aberta
Private lngKey As Long 'ponteiro para Key aberta
Private strKey As String 'HKEY atual
Private strMachine As String 'nome da maquina
Private blnAutoCreate As Boolean 'autocreate local
Private strValueAsString As String 'ValueAsString
Private blnValueAsBoolean As Boolean 'ValueAsBoolean
Private vntValueAsBinary As Variant 'ValueAsBinary
Private lngValueAsLong As Long 'ValueAsLong
Private objValues As RegistryValues 'coleção contendo os valores
Private objSubKeys As RegistrySubKeys 'coleção contendo as subchaves

Private Sub Class_Initialize()
  TimeBombX xtbRegistry, 15656, 12934, 3080, 14624
End Sub

Rem =========================================================================
Rem Levanta uma excecao
Rem -------------------------------------------------------------------------
Private Sub RaiseException(ByVal lngError As Long, ByVal strCaller As String)
    Err.Raise vbObjectError + lngError, _
        App.ProductName & ".Registry." & strCaller, _
        APIErrorText(lngError)
        
End Sub

Rem =========================================================================
Rem Abre/cria uma chave
Rem -------------------------------------------------------------------------
Private Function OpenKey(lngAccess As Long) As Long
    Dim lngReturn As Long
    
    If strMachine > "" Then
      If RegConnectRegistry(strMachine, lngHKEY, lngHKEY) <> ERROR_SUCCESS Then
        WNetCancelConnection2 strMachine & "\ipc$", 0&, False
      End If
    End If
    
    If blnAutoCreate Then
        lngReturn = RegCreateKeyEx(lngHKEY, strKey, 0&, vbNullString, _
                          REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                          0&, lngKey, 0&)
    Else
        lngReturn = RegOpenKeyEx(lngHKEY, strKey, 0, lngAccess, lngKey)
    
    End If
    
    OpenKey = lngReturn
End Function


Rem =========================================================================
Rem Fecha HKEY e Key
Rem -------------------------------------------------------------------------
Private Sub CloseAllKeys()
    RegCloseKey (lngKey)
    RegCloseKey (lngHKEY)
End Sub


Rem =========================================================================
Rem Acessa/cria um valor
Rem -------------------------------------------------------------------------
Private Function ValueData(ByVal strValue As String, ByVal vntData As Variant, ByRef lngValueType As Long, ByVal blnGetValue As Boolean) As Variant
    Dim arrValue() As Byte, _
        lngExistingValueType As Long, _
        lngReturn As Long, _
        lngValueLength As Long, _
        vntResult As Variant, _
        strExpanded As String, _
        strRegValue As String, _
        lngRegValue As Long, _
        i As Long, _
        iStartByte As Long, _
        iBit As Long, _
        intSignal As Integer
        
    lngReturn = OpenKey(KEY_READ)
    
    If lngReturn = ERROR_SUCCESS Then '//se não ocorrer erro
        If blnGetValue Then
            lngValueLength = MAX_REG_STRING_LEN
            ReDim arrValue(lngValueLength)
        
        End If
       
        If blnGetValue Then
            lngReturn = RegQueryValueEx(lngKey, strValue, 0&, lngExistingValueType, arrValue(0), lngValueLength)

            Select Case lngReturn
        
                Case ERROR_FILE_NOT_FOUND
                    If blnAutoCreate Then
                        lngReturn = RegSetValueEx(lngKey, strValue, 0&, lngValueType, vntResult, lngValueLength)
                        vntResult = Null
                    
                    End If
        
                Case ERROR_SUCCESS '//se não ocorrer erro
                    Select Case lngExistingValueType
                        Case REG_BINARY '// assume vntResult passed as byte array
                            ReDim vntResult(lngValueLength - 1)
                            For i = 0 To (lngValueLength - 1)
                                vntResult(i) = arrValue(i) '// set bytes arrays equal
                            Next
    
                        Case REG_DWORD
                            lngReturn = RegQueryValueExLong(lngKey, strValue, 0&, lngValueType, vntResult, lngValueLength)
                            
                        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
                            strRegValue = Space(lngValueLength)
                            lngReturn = RegQueryValueExString(lngKey, strValue, 0&, lngValueType, strRegValue, lngValueLength)
                            vntResult = strRegValue
        
                            If (lngValueLength = REG_EXPAND_SZ) Then
                                lngValueLength = ExpandEnvironmentStrings(vntResult, vbNullString, Null)
                                strExpanded = Space(lngValueLength) '//preenche com espaços
                                lngValueLength = ExpandEnvironmentStrings(vntResult, strExpanded, lngValueLength)
                                vntResult = strExpanded

                            End If
                    End Select
            
            End Select
        Else
            Select Case lngValueType
                Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
                    lngValueLength = Len(vntData) + 1
                    strRegValue = CStr(vntData)
                    lngReturn = RegSetValueExString(lngKey, strValue, 0&, lngValueType, strRegValue, lngValueLength)
                
                Case REG_DWORD
                    lngValueLength = MAX_REG_DWORD_LEN
                    lngRegValue = CLng(vntData)
                    lngReturn = RegSetValueExLong(lngKey, strValue, 0&, lngValueType, lngRegValue, lngValueLength)
            
                Case REG_BINARY
                    lngValueLength = UBound(vntData)
                    ReDim arrValue(lngValueLength + 1)
                    For i = 0 To lngValueLength
                        arrValue(i) = vntData(i) '// set bytes arrays equal
                        Debug.Print vntData(i)
                    Next
                    arrValue(lngValueLength + 1) = 0
                    
                    lngReturn = RegSetValueEx(lngKey, strValue, 0&, lngValueType, arrValue(0), lngValueLength)
            
            End Select
            
        End If
        
    End If
    CloseAllKeys
    
    If lngReturn <> ERROR_SUCCESS Then
        RaiseException lngReturn, "ValueData"
    
    End If
    
    If blnGetValue Then ValueData = vntResult

End Function

Rem =========================================================================
Rem Cria um valor
Rem -------------------------------------------------------------------------
Public Function Exists(Optional strKeyName As String) As Boolean
    Dim lngReturn As Long, _
    stFileTime As FILETIME

'    lngReturn = RegQueryInfoKey(lngKey, ByVal 0, ByVal 0, ByVal 0, ByVal 0, ByVal 0, ByVal 0, ByVal 0, ByVal 0, ByVal 0, ByVal 0, stFileTime)
    
    Debug.Print lngReturn & ":  " & APIErrorText(lngReturn)
    
End Function

Rem =========================================================================
Rem Cria um valor
Rem -------------------------------------------------------------------------
Public Function CreateValue(strValueName As String, Optional vntValueData As Variant, Optional intValueType As Long = REG_SZ) As Boolean
    Dim blnLocalAutoCreate As Boolean, _
    lngReturn As Long
    
    lngReturn = OpenKey(KEY_ALL_ACCESS)
    
    blnLocalAutoCreate = blnAutoCreate
    blnAutoCreate = True
    ValueData strValueName, vntValueData, intValueType, False
    blnAutoCreate = blnLocalAutoCreate
    
    CloseAllKeys
    CreateValue = (lngReturn = ERROR_SUCCESS)
End Function

Rem =========================================================================
Rem Cria a chave definida na propriedade Key.
Rem -------------------------------------------------------------------------
Public Function CreateKey() As Boolean
    Dim blnLocalAutoCreate As Boolean, _
    lngReturn As Long
    
    blnLocalAutoCreate = blnAutoCreate
    blnAutoCreate = True
    lngReturn = OpenKey(KEY_ALL_ACCESS)
    blnAutoCreate = blnLocalAutoCreate
    CloseAllKeys
    CreateKey = (lngReturn = ERROR_SUCCESS)
End Function

Rem =========================================================================
Rem Deleta um valor dentro da chave definida na propriedade Key.
Rem -------------------------------------------------------------------------
Public Function DeleteValue(strValueName As String) As Boolean
    Dim lngReturn As Long

    lngReturn = OpenKey(KEY_ALL_ACCESS)
    If lngReturn = ERROR_SUCCESS Then '//se não ocorrer erro
        lngReturn = RegDeleteValue(lngKey, strValueName)
    End If
    
    CloseAllKeys
    DeleteValue = (lngReturn = ERROR_SUCCESS)
End Function

Rem =========================================================================
Rem Deleta a chave definida na propriedade Key.
Rem -------------------------------------------------------------------------
Public Function DeleteKey() As Boolean
    Dim lngReturn As Long
    
    lngReturn = RegDeleteKey(lngHKEY, strKey)
    
    DeleteKey = (lngReturn = ERROR_SUCCESS)
End Function

Rem =========================================================================
Rem Propriedade SubKeys
Rem -------------------------------------------------------------------------
Public Function SubKeys() As RegistrySubKeys
    Dim i As Long, _
    lngReturn As Long, _
    intIndex As Integer, _
    lngValueLength As Long, _
    stLastWrite As FILETIME, _
    strName As String
            
    lngReturn = OpenKey(KEY_ENUMERATE_SUB_KEYS)
    
    If lngReturn = ERROR_SUCCESS Then '//se não ocorrer erro
        Set objSubKeys = New RegistrySubKeys
        intIndex = 0
        
        Do
            lngValueLength = MAX_REG_STRING_LEN
            strName = Space(lngValueLength)
            lngReturn = RegEnumKeyEx(lngKey, intIndex, strName, lngValueLength, ByVal 0, vbNullString, ByVal 0, stLastWrite)
    
            If lngReturn = ERROR_SUCCESS Then
                objSubKeys.Add Left(strName, lngValueLength), FileTimeToDate(stLastWrite), Me
                intIndex = intIndex + 1
            End If
            
        Loop While lngReturn = ERROR_SUCCESS
        
    End If

    If lngReturn <> ERROR_NO_MORE_ITEMS Then
        RaiseException lngReturn, "SubKeys"
        
    Else
        Set SubKeys = objSubKeys
        Set objSubKeys = Nothing
    
    End If

End Function


Rem =========================================================================
Rem Propriedade Values
Rem -------------------------------------------------------------------------
Public Property Get Values() As RegistryValues
    Dim i As Long, _
    lngReturn As Long, _
    intIndex As Integer, _
    lngValueLength As Long, _
    lngValueType As Long, _
    strName As String
            
    lngReturn = OpenKey(KEY_READ)
    
    If lngReturn = ERROR_SUCCESS Then '//se não ocorrer erro
        Set objValues = New RegistryValues
        intIndex = 0
        
        Do
            lngValueLength = MAX_REG_STRING_LEN
            strName = Space(lngValueLength)
            lngReturn = RegEnumValue(lngKey, intIndex, strName, lngValueLength, ByVal 0, lngValueType, ByVal 0, ByVal 0)
       
            If lngReturn = ERROR_SUCCESS Then
                objValues.Add Left(strName, lngValueLength), lngValueType, Me
                intIndex = intIndex + 1
            End If
            
        Loop While lngReturn = ERROR_SUCCESS
        
    End If

    If lngReturn <> ERROR_NO_MORE_ITEMS Then
        RaiseException lngReturn, "Values"
        
    Else
        Set Values = objValues
        Set objValues = Nothing
    
    End If

    
End Property

Rem =========================================================================
Rem Atribui um valor REG_DWORD
Rem -------------------------------------------------------------------------
Public Property Let ValueAsLong(ByVal strValue As String, ByVal lngData As Long)
    ValueData strValue, lngData, REG_DWORD, False
End Property


Rem =========================================================================
Rem Retorna um valor REG_DWORD
Rem -------------------------------------------------------------------------
Public Property Get ValueAsLong(ByVal strValue As String) As Long
    lngValueAsLong = CLng(ValueData(strValue, Null, REG_DWORD, True))
    ValueAsLong = lngValueAsLong
End Property

Rem =========================================================================
Rem Atribui o nome da maquina
Rem -------------------------------------------------------------------------
Public Property Let Machine(ByVal strValue As String)
    If Left(strValue, 2) = "\\" Then
      strMachine = strValue
    Else
      strMachine = "\\" & strValue
    End If
End Property


Rem =========================================================================
Rem Retorna o nome da maquina
Rem -------------------------------------------------------------------------
Public Property Get Machine() As String
    Machine = strMachine
End Property


Rem =========================================================================
Rem Atribui um valor REG_BINARY
Rem -------------------------------------------------------------------------
Public Property Let ValueAsBinary(ByVal strValue As String, ByVal vntData As Variant)
    ValueData strValue, vntData, REG_BINARY, False
End Property


Rem =========================================================================
Rem Retorna um valor REG_BINARY
Rem -------------------------------------------------------------------------
Public Property Get ValueAsBinary(ByVal strValue As String) As Variant
    vntValueAsBinary = ValueData(strValue, Null, REG_BINARY, True)
    ValueAsBinary = vntValueAsBinary
End Property


Rem =========================================================================
Rem Atribui um valor REG_DWORD como Boolean
Rem -------------------------------------------------------------------------
Public Property Let ValueAsBoolean(ByVal strValue As String, ByVal blnData As Boolean)
    Dim lngData As Long
    
    If blnData Then
        lngData = 1
    Else
        lngData = 0
    End If
    
    ValueData strValue, lngData, REG_DWORD, False
End Property


Rem =========================================================================
Rem Retorna um valor REG_DWORD como Boolean
Rem -------------------------------------------------------------------------
Public Property Get ValueAsBoolean(ByVal strValue As String) As Boolean
    blnValueAsBoolean = CBool(ValueData(strValue, Null, REG_DWORD, True))
    ValueAsBoolean = blnValueAsBoolean
End Property


Rem =========================================================================
Rem Atribui um valor REG_SZ
Rem -------------------------------------------------------------------------
Public Property Let ValueAsString(ByVal strValue As String, ByVal strData As String)
    ValueData strValue, strData, REG_SZ, False
End Property


Rem =========================================================================
Rem Retorna um valor REG_SZ
Rem -------------------------------------------------------------------------
Public Property Get ValueAsString(ByVal strValue As String) As String
    strValueAsString = CStr(ValueData(strValue, Null, REG_SZ, True))
    ValueAsString = strValueAsString
End Property


Rem =========================================================================
Rem Define a criação automática de chaves e valores
Rem -------------------------------------------------------------------------
Public Property Let AutoCreate(ByVal blnData As Boolean)
    blnAutoCreate = blnData
End Property


Rem =========================================================================
Rem Retorna o status da criação automática de chaves e valores
Rem -------------------------------------------------------------------------
Public Property Get AutoCreate() As Boolean
    AutoCreate = blnAutoCreate
End Property



Rem =========================================================================
Rem Define a chave a ser acessada
Rem -------------------------------------------------------------------------
Public Property Let Key(ByVal strData As String)
    strKey = strData
End Property


Rem =========================================================================
Rem Retorna a chave a ser acessada
Rem -------------------------------------------------------------------------
Public Property Get Key() As String
    Key = strKey
End Property


Rem =========================================================================
Rem Define o HKEY atual
Rem -------------------------------------------------------------------------
Public Property Let hKey(ByVal lngData As Long)
    lngHKEY = lngData
End Property

Rem =========================================================================
Rem Retorna uma string representando o HKEY atual
Rem -------------------------------------------------------------------------
Public Property Get hKeyName() As String
    Select Case lngHKEY
        Case HKEY_CLASSES_ROOT
            hKeyName = "HKEY_CLASSES_ROOT"
            
        Case HKEY_CURRENT_USER
            hKeyName = "HKEY_CURRENT_USER"
            
        Case HKEY_LOCAL_MACHINE
            hKeyName = "HKEY_LOCAL_MACHINE"
            
        Case HKEY_USERS
            hKeyName = "HKEY_USERS"
            
        Case HKEY_PERFORMANCE_DATA
            hKeyName = "HKEY_PERFORMANCE_DATA"
            
        Case HKEY_CURRENT_CONFIG
            hKeyName = "HKEY_CURRENT_CONFIG"
            
        Case HKEY_DYN_DATA
            hKeyName = "HKEY_DYN_DATA"
            
        Case Else
            hKeyName = Null
            
    End Select

End Property
