VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "UploadItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
' =========================================================================
'  /HiperTools30.UploadItem
' -------------------------------------------------------------------------
'  Descricao   : Itens de upload
'  Criaçao     : 11/01/2000 12:18
'  Local       : Brasilia/DF
'  Elaborado   : Ridai Govinda Pombo <ridai@zevallos.com.br>
'                Kelson da Costa Medeiros <kelson@zevallos.com.br>
'  Versao      : 2.0.0
'  Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
' -------------------------------------------------------------------------
Option Explicit

' Error Definition
Const ERR_FILESIZE_NOT_ALLOWED As Long = vbObjectError + 102, _
      ERR_FOLDER_DOES_NOT_EXIST As Long = vbObjectError + 103, _
      ERR_FILE_ALREADY_EXISTS As Long = vbObjectError + 104

Private File As HiperTools30.File

Private sintPosition As Long, _
        sstrName As String, _
        svntValue As Variant, _
        sstrUserFilename As String, _
        sobjUpload As Object, _
        sstrPath As String, _
        slngTotalBytes As Long, _
        sstrServerName As String

Private slngMaxBytes As Long, _
        sblnOverWriteFiles As Boolean
        
Private sstrContentDisposition As String, _
        sstrContentType As String
        
Private sstrContentTransferEncoding As String, _
        sdtmDateLastModified As Date, _
        sdtmDateLastModifiedUTC As Date, _
        sblnUseDateLastModified As Boolean

'ASP :
Private ScriptContext As ASPTypeLibrary.ScriptingContext, _
        Response As ASPTypeLibrary.Response, _
        Server As ASPTypeLibrary.Server

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  TimeBombX xtbUpload, 15656, 12934, 3080, 14624
  
  Set ScriptContext = ScriptContext
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  
End Sub

Public Sub OnEndPage()
  
  Set ScriptContext = Nothing
  Set Response = Nothing
  Set Server = Nothing

End Sub

Private Sub Class_Initialize()

  Set File = New HiperTools30.File

  sintPosition = 1

End Sub

Private Sub Class_Terminate()

  Set File = Nothing
  Set sobjUpload = Nothing

End Sub

Rem =========================================================================
Rem Índice deste cara
Rem -------------------------------------------------------------------------
Friend Property Get Position() As Long
  Position = sintPosition
End Property

Friend Property Let Position(ByVal Value As Long)
  sintPosition = Value
End Property

Rem =========================================================================
Rem Nome do campo do formulário através do qual o arquivo foi enviado
Rem -------------------------------------------------------------------------
Public Property Get Name() As String
  Name = sstrName
End Property

Friend Property Let Name(ByVal Value As String)
  sstrName = Value
End Property

Rem =========================================================================
Rem Arquivo enviado através do Upload
Rem -------------------------------------------------------------------------
Public Property Get Value() As Variant
Attribute Value.VB_Description = "Arquivo enviado através do Upload"
Attribute Value.VB_UserMemId = 0
  Value = svntValue
End Property

Friend Property Let Value(ByVal vntValue As Variant)
  svntValue = vntValue
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get UserFilename() As String
  UserFilename = sstrUserFilename
End Property

Friend Property Let UserFilename(ByVal Value As String)
  sstrUserFilename = Value
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get TotalBytes() As Long
  TotalBytes = slngTotalBytes
End Property

Friend Property Let TotalBytes(ByVal Value As Long)
  slngTotalBytes = Value
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ContentDisposition() As String
Attribute ContentDisposition.VB_Description = "The MIME Content Disposition of the data. This should always be ""form-data"" when using a browser that supports RFC 1867 uploads."
  ContentDisposition = sstrContentDisposition
End Property

Friend Property Let ContentDisposition(ByVal Value As String)
  sstrContentDisposition = Value
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ContentTransferEncoding() As String
Attribute ContentTransferEncoding.VB_Description = "This property is typically blank, but may contain the MIME Content Transfer Encoding."
  ContentTransferEncoding = sstrContentTransferEncoding
End Property

Friend Property Let ContentTransferEncoding(ByVal Value As String)
  sstrContentTransferEncoding = Value
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ContentType() As String
Attribute ContentType.VB_Description = "The MIME Content Type of the file, which is useful in determining the contents of the file."
  ContentType = sstrContentType
End Property

Friend Property Let ContentType(ByVal Value As String)
  sstrContentType = Value
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get DateLastModified() As Date
Attribute DateLastModified.VB_Description = "Show the last modification date of the file transmitted by the user. The date is displayed using the local time zone. For the UTC date of the file, see the DateLastModifiedUTC property."
  DateLastModified = sdtmDateLastModified
End Property

Friend Property Let DateLastModified(ByVal Value As Date)
  sdtmDateLastModified = Value
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get DateLastModifiedUTC() As Date
Attribute DateLastModifiedUTC.VB_Description = "Show the last modification date of the file transmitted by the user. The date is displayed using Universal Time (UTC, or GMT). For the local date of the file, see the DateLastModified property."
  DateLastModifiedUTC = sdtmDateLastModifiedUTC
End Property

Friend Property Let DateLastModifiedUTC(ByVal Value As Date)
  sdtmDateLastModifiedUTC = Value
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get UseDateLastModified() As Boolean
Attribute UseDateLastModified.VB_Description = "This property determines if an uploaded file will preserve the Last Modified Date of the user's original file. The file on the server will have the same creation and modification date as the user's original file. This only works if the client upload soft"
  UseDateLastModified = sblnUseDateLastModified
End Property

Public Property Let UseDateLastModified(ByVal Value As Boolean)
  sblnUseDateLastModified = Value
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Friend Property Get Path() As String
  ' Se este for um arquivo de upload
  If IsFile Then

    ' Se não foi definido um path específico para este arquivo.
    If sstrPath > "" Then
    
      ' Retorna o path específico deste arquivo
      Path = sstrPath

    Else

      ' Retorna o path da classe upload
      Path = Upload.Path

    End If

  End If

End Property

Friend Property Let Path(ByVal Value As String)
  sstrPath = Value
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Friend Property Get Upload() As Object
  Set Upload = sobjUpload
End Property

Friend Property Set Upload(ByVal Value As Object)
  Set sobjUpload = Value
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ServerName() As String
  ServerName = Path & File.ExtractFileName(UserFilename)
End Property

Rem =========================================================================
Rem Informa se este item é um campo de formulário apenas ou se é um arquivo para upload
Rem -------------------------------------------------------------------------
Public Property Get IsFile() As Boolean
  IsFile = (UserFilename > "")
End Property

Rem ===============================================================================
Rem Descrição:  This method saves an uploaded file to the web server's hard disk using
Rem the User's original filename..
Rem -------------------------------------------------------------------------------
Public Sub Save()

  ' Salva o arquivo com o nome original no caminho específicado pelo usuário
  WriteFile File.ExtractFileName(sstrServerName)

End Sub

Rem =================================================================================
Rem Descrição:  Este método salva um arquivo enviado por upload no servidor WEB.
Rem Parametros: strFilename -> Nome completo do arquivo a ser salvo.
Rem ---------------------------------------------------------------------------------
Public Sub SaveAs(ByVal strFileName As String)

  ' Salva o arquivo com o nome específicado pelo usuário que deverá ser passado em o
  ' caminho
  WriteFile strFileName

End Sub

Rem ================================================================================
Rem Descrição:  Salva o arquivo enviado no local específicado
Rem Parametros: intIndex -> Índice do item a ser salvo
Rem --------------------------------------------------------------------------------
Private Sub WriteFile(ByVal strFileName As String)

  ' Variables for file
  Dim fs As Scripting.FileSystemObject, _
      sFile As TextStream

  ' Check for maximum bytes allowed
  If Upload.MaxBytes <> 0 And Upload.MaxBytes < TotalBytes Then
     Err.Raise ERR_FILESIZE_NOT_ALLOWED, , "O arquivo [<b>" & UserFilename & "</b>] foi enviado com [<b>" & _
                                           TotalBytes & "</b>] bytes, sendo que o tamanha máximo é de " & _
                                           "[<b>" & Upload.MaxBytes & "</b>]."
  End If

  Set fs = New Scripting.FileSystemObject

  If Not Right(Path, 1) = "\" Then
    Path = Path & "\"
  End If

  If Not fs.FolderExists(Path) Then
    Err.Raise ERR_FOLDER_DOES_NOT_EXIST, , "O caminho '" & Path & "' não existe."
  End If

  If fs.FileExists(Path & strFileName) And (Not Upload.OverWriteFiles) Then
    Err.Raise ERR_FILE_ALREADY_EXISTS, , "O arquivo '" & strFileName & "' já existe."
  End If
  
  ' Create file
  Set sFile = fs.CreateTextFile(Path & strFileName, True)

  ' Write file
  sFile.Write svntValue

  ' Close File
  sFile.Close

  ' Destrói o cara
  Set sFile = Nothing
  Set fs = Nothing

End Sub
