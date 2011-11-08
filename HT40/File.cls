VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "File"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =========================================================================
' File.cls
' -------------------------------------------------------------------------
' Descricao   : Include de File
' Criaçao     : 11:23h 7/2/1998
' Local       : Brasilia/DF
' Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------

Option Explicit

Const MAX_PATH = 260

Private Server As Server

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  TimeBombX xtbFile, 15656, 12934, 3080, 14624
  
  Set Server = ScriptContext.Server
 
End Sub

Public Sub OnEndPage()
  Set Server = Nothing
End Sub

Rem -------------------------------------------------------------------------
Rem Verifica se o arquivo existe
Rem -------------------------------------------------------------------------
Public Function Exists(ByVal strFullPath As String) As Boolean
  
  Dim objFS As Scripting.FileSystemObject

  Set objFS = New Scripting.FileSystemObject
 
  If objFS.GetDriveName(strFullPath) = "" Then
    strFullPath = Server.MapPath(strFullPath)
  
  End If
  
  Exists = objFS.FileExists(strFullPath)

  Set objFS = Nothing
  
End Function

Rem -------------------------------------------------------------------------
Rem Copia arquivos
Rem -------------------------------------------------------------------------
Public Sub Copy(ByVal strSource As String, ByVal strDestination As String, Optional ByVal blnOverwrite As Boolean = False)

  Dim objFS 'As Scripting.FileSystemObject
  Dim strFullPath As String

'  Set objFS = New Scripting.FileSystemObject
  Set objFS = Server.CreateObject("Scripting.FileSystemObject")

  If objFS.GetDriveName(strFullPath) = "" Then
    strFullPath = Server.MapPath(strFullPath)
  End If

  objFS.CopyFile strSource, strDestination, blnOverwrite

  Set objFS = Nothing
  
End Sub

Rem -------------------------------------------------------------------------
Rem Extrai o nome do arquivo do Path
Rem -------------------------------------------------------------------------
Public Function ExtractFileName(ByVal sFileName As String) As String
  
  Dim nIdx As Long
  Dim Path As HiperTools30.Path
  Dim strResult As String
  Dim strSlash As String
  
  Set Path = New HiperTools30.Path
  
  If sFileName > "" Then
    
    If InStr(sFileName, ".") > 0 Then
      strResult = sFileName
      strSlash = Path.SlashType(sFileName)
      
      If strSlash > "" Then
         strResult = Mid(sFileName, InStrRev(sFileName, strSlash) + 1)
      End If
    
    End If
  
  End If
  
  ExtractFileName = strResult
  
  Set Path = New HiperTools30.Path

End Function

Rem -------------------------------------------------------------------------
Rem Extrai o nome do arquivo do Path
Rem -------------------------------------------------------------------------
Public Function ExtractFilePath(ByVal strFullPath As String) As String
  
  Dim strResult

  strResult = ""
  
  If strFullPath > "" Then
    strResult = strFullPath
    
    If InStr(strFullPath, ".") > 0 Then
      strResult = Left(strFullPath, Len(strFullPath) - Len(ExtractFileName(strFullPath)))
      
    End If
  End If
   
  ExtractFilePath = strResult

End Function

Rem -------------------------------------------------------------------------
Rem Retorna a data do arquivo
Rem -------------------------------------------------------------------------
Public Function DateCreated(ByVal strFullPath As String) As Variant
  
  Dim objFS As Scripting.FileSystemObject, objFile

  Set objFS = New Scripting.FileSystemObject

  If objFS.GetDriveName(strFullPath) = "" Then
    strFullPath = Server.MapPath(strFullPath)
  End If

  Set objFile = objFS.GetFile(strFullPath)
 
  DateCreated = objFile.DateCreated

  Set objFile = Nothing
  Set objFS = Nothing

End Function

Rem -------------------------------------------------------------------------
Rem Retorna a data de atualizacao do arquivo
Rem -------------------------------------------------------------------------
Public Function DateLastModified(ByVal strFullPath As String) As String
  
  Dim objFS As Scripting.FileSystemObject, objFile

  Set objFS = New Scripting.FileSystemObject

  If objFS.GetDriveName(strFullPath) = "" Then
    strFullPath = Server.MapPath(strFullPath)
  
  End If

  Set objFile = objFS.GetFile(strFullPath)
 
  DateLastModified = objFile.DateLastModified

  Set objFile = Nothing

  Set objFS = Nothing

End Function

Rem -------------------------------------------------------------------------
Rem Retorna a data do arquivo
Rem -------------------------------------------------------------------------
Public Function Size(ByVal strFullPath As String) As Long

  Dim objFS As Scripting.FileSystemObject, objFile

  Set objFS = New Scripting.FileSystemObject

  If objFS.GetDriveName(strFullPath) = "" Then
    strFullPath = Server.MapPath(strFullPath)
  End If
 
  Set objFile = objFS.GetFile(strFullPath)

  Size = objFile.Size

  Set objFile = Nothing
  Set objFS = Nothing

End Function

Rem =========================================================================
Rem As propriedades descritas abaixo serão mantidas apenas por questão de
Rem compatibilidade com as versões anteriores
Rem ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Public Function IsFileExist(ByVal strFullPath As String) As Boolean 'Obsoleto
  IsFileExist = Exists(strFullPath)
End Function

'Public Function FileCopy(ByVal strSource As String, ByVal strDestination As String, ByVal blnOverwrite As Boolean) As Boolean 'Obsoleto
'  FileCopy = Copy(strSource, strDestination, blnOverwrite)
'End Function

'Public Function FileCopyOverwrite(ByVal strSource As String, ByVal strDestination As String) As Boolean 'Obsoleto
'  FileCopyOverwrite = Copy(strSource, strDestination, True)
'End Function

Public Function FileDateCreated(ByVal strFullPath As String) As String 'Obsoleto
  FileDateCreated = DateCreated(strFullPath)
End Function

Public Function FileDateLastModified(ByVal strFullPath As String) As String 'Obsoleto
  FileDateLastModified = DateLastModified(strFullPath)
End Function

Public Function FileSize(ByVal strFullPath As String) As String 'Obsoleto
  FileSize = Size(strFullPath)
End Function

Public Function GetTempFile() As String
    
    Dim lngRet As Long
    Dim strBuffer As String, strTempPath As String
    
    'Initialize buffer
    strBuffer = String$(MAX_PATH, 0)
    
    'Get temp path
    lngRet = GetTempPath(Len(strBuffer), strBuffer)
    
    'Validate return value
    If lngRet = 0 Then Exit Function
    
    'Remove trailing nulls
    strTempPath = Left$(strBuffer, lngRet)
    
    'Re-initialize buffer
    strBuffer = String$(MAX_PATH, 0)
    
    'Get temp filename
    lngRet = GetTempFileName(strTempPath, "tmp", 0&, strBuffer)
    
    'Validate return value
    If lngRet = 0 Then Exit Function
    
    'Trim nulls and then return filename
    lngRet = InStr(1, strBuffer, Chr(0))
    If lngRet > 0 Then
        GetTempFile = Left$(strBuffer, lngRet - 1)
    Else
        GetTempFile = strBuffer
    End If
    
End Function

