VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Upload"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =================================================================================
' Upload.cls
' ---------------------------------------------------------------------------------
' Descri��o : Recebe arquivos atrav�s do prot�colo HTTP e salva estes
'             arquivos no servidor WEB
' Criacao   : 04/02/2000 10:50AM
' Local     : Brasilia/DF
' Elaborado : Kelson da Costa Medeiros. - kelson@zevallos.com.br
' Versao    : 1.0.0
' Copyright : 1999 by Zevallos(r) Tecnologia em Informacao
' ---------------------------------------------------------------------------------


Option Explicit

' Error Definition
Const ERR_NO_FILENAME As Long = vbObjectError + 100, _
      ERR_EMPTY_FILE As Long = vbObjectError + 101, _
      ERR_FILESIZE_NOT_ALLOWED As Long = vbObjectError + 102, _
      ERR_FOLDER_DOES_NOT_EXIST As Long = vbObjectError + 103, _
      ERR_FILE_ALREADY_EXISTS As Long = vbObjectError + 104, _
      ERR_DONT_HAVE_UPLOAD_FILE As Long = vbObjectError + 105

' Constantes
Const conQuotes = """"

Private ScriptContext As ASPTypeLibrary.ScriptingContext, _
        Server As ASPTypeLibrary.Server, _
        Request As ASPTypeLibrary.Request, _
        Session As ASPTypeLibrary.Session

Private Strings As HiperTools30.String, _
        File As HiperTools30.File, _
        sobjUploadItems As HiperTools30.UploadItems, _
        sstrPath As String, _
        sobjFirstFile As HiperTools30.UploadItem

Public MaxBytes As Long
Attribute MaxBytes.VB_VarDescription = "Define o n�mero m�ximo de bytes que podem ser escritos no disco ou dentro do banco de dados por arquivo. Se for igual zero, n�o h� limite."
Public OverWriteFiles As Boolean
Attribute OverWriteFiles.VB_VarDescription = "Determina se os arquivos devem sobrepor os arquivos j� existentes. O valor padr�o desta proprieade e ""True""."
Public Path As String

Public Sub OnStartPage(ScriptContext As ScriptingContext)

  TimeBombX xtbUpload, 15656, 12934, 3080, 14624
  
  Set ScriptContext = ScriptContext
  
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Request = ScriptContext.Request
  Set Strings = Server.CreateObject("HiperTools30.String")
  Set File = Server.CreateObject("HiperTools30.File")
  Set sobjUploadItems = Server.CreateObject("HiperTools30.UploadItems")
  
  ' L� os arquivos cabe�alho do arquivo HTTP e gera a lista de itens com os arquivos e campos de formul�rios
  ReadFiles

  ' Configura os valores padronizados das propriedades
  MaxBytes = 0
  OverWriteFiles = True
  Path = LCase(Request.ServerVariables("APPL_PHYSICAL_PATH"))

End Sub

Public Sub OnEndPage()

  Set Strings = Nothing
  Set File = Nothing
  Set sobjUploadItems = Nothing

  Set Server = Nothing
  Set Session = Nothing
  Set Request = Nothing
  Set ScriptContext = Nothing
  
End Sub

'Rem ===============================================================================
'Rem Descri��o:  Retorna o nome completo (caminho + nome) do arquivo a ser salvo no
'Rem             servidor WEB.
'Rem -------------------------------------------------------------------------------
'Public Property Get ServerName() As String
'
'  ' Invoca o m�todo "ServerName" do primeiro arquivo para upload
'  ServerName = GetFirstFile.ServerName
'
'End Property
'
'Rem ===============================================================================
'Rem Descri��o: Retorna o n�mero total de bytes que ser�o armazenados no disco.
'Rem -------------------------------------------------------------------------------
'Public Property Get TotalBytes() As Long
'
'  ' Invoca o m�todo "TotalBytes" do primeiro arquivo para upload
'  TotalBytes = GetFirstFile.TotalBytes
'
'End Property
'
'Rem ===============================================================================
'Rem Descri��o:  Retorna o nome original completo do arquivo na m�quina do usu�rio.
'Rem -------------------------------------------------------------------------------
'Public Property Get UserFilename() As Long
'
'  ' Invoca o m�todo "UserFilename" do primeiro arquivo para upload
'  UserFilename = GetFirstFile.UserFilename
'
'End Property

Rem ===============================================================================
Rem Descri��o:  Retorna a lista de campos do formul�rio e seus respectivos valores
Rem -------------------------------------------------------------------------------
Public Property Get Form() As HiperTools30.UploadItems
  Set Form = sobjUploadItems
End Property

Rem ===============================================================================
Rem Descri��o:  Retorna a lista de campos do formul�rio e seus respectivos valores
Rem -------------------------------------------------------------------------------
Public Property Get Items() As HiperTools30.UploadItems
  Set Items = sobjUploadItems
End Property

Rem ===============================================================================
Rem Descri��o:  This method saves an uploaded file to the web server's hard disk using
Rem the User's original filename..
Rem -------------------------------------------------------------------------------
Public Sub Save()

  ' Invoca o m�todo "Save" do primeiro arquivo para upload
  GetFirstFile.Save

End Sub

Rem =================================================================================
Rem Descri��o:  Este m�todo salva um arquivo enviado por upload no servidor WEB.
Rem Parametros: strFilename -> Nome completo do arquivo a ser salvo.
Rem ---------------------------------------------------------------------------------
Public Sub SaveAs(ByVal strFileName As String)

  ' Invoca o m�todo "SaveAs" do primeiro arquivo para upload
  GetFirstFile.SaveAs strFileName

End Sub

Rem ==================================================================================
Rem Descri��o:  L� os arquivos enviados e armazena em vari�veis para que o usu�rio
Rem             possa baixa-los depois.
Rem Parametros: <Nenhum>
Rem ----------------------------------------------------------------------------------
Private Sub ReadFiles()

  ' Vari�veis
  Dim intByteCount As Variant, _
      vntHTTPHeader As Variant, _
      vntDelimeter As Variant, _
      intCounter As Integer, _
      lngNameStart As Long, _
      lngNameEnd As Long, _
      strName As String, _
      lngValueStart As Long, _
      lngValueEnd As Long, _
      strValue As String, _
      strFileName As String, _
      lngFileDataStart As Long, _
      lngFileDataEnd As Long, _
      lngFileLength As Long, _
      i As Long
  Dim objUpFile As HiperTools30.UploadItem

  ' L� o tamanho do arquivo enviado pelo HTTP-header
  intByteCount = Request.TotalBytes

  ' L� o arquivo enviado pelo HTTP-header e converte para unicode
  vntHTTPHeader = StrConv(Request.BinaryRead(intByteCount), vbUnicode)
  
  ' Armazena o delimitador de campos deste HTTP-header
  vntDelimeter = LeftB(vntHTTPHeader, 76)

  ' Armazena a quantidade de FormFields que existem no HTTP-header
  intCounter = Strings.CountSubString(vntHTTPHeader, "; name=")

  ' Percorre todos os campos do formul�rio enviado a procura de campos de upload
  For i = 1 To intCounter

    ' Armazena o in�cio do pr�ximo nome de campo do formul�rio
    lngNameStart = InStrB(lngNameStart + 1, vntHTTPHeader, "; name=" & conQuotes) + 16

    ' Armazena o final do nome do campo
    lngNameEnd = InStrB(lngNameStart, vntHTTPHeader, conQuotes) + 2

    ' Armazena o nome do campo
    strName = MidB(vntHTTPHeader, lngNameStart, lngNameEnd - lngNameStart - 2)

    ' Verifica se este campo do formul�rio � um arquivo de upload
    If MidB(vntHTTPHeader, lngNameEnd, 24) = ("; filename=" & conQuotes) Then

      ' Armazena o tamanho do nome do arquivo
      lngValueStart = InStrB(lngNameEnd, vntHTTPHeader, "filename=" & conQuotes) + 20

      ' Armazena o tamanho do nome do arquivo
      lngValueEnd = InStrB(lngValueStart, vntHTTPHeader, conQuotes)

      ' Armazena o nome do arquivo armazenado neste campo do formul�rio
      strFileName = MidB(vntHTTPHeader, lngValueStart, lngValueEnd - lngValueStart)

      ' Se foi informado um nome de arquivo v�lido
      If strFileName >= "" Then
        
        ' Armazena o �nicio dos dados do arquivo
        lngFileDataStart = InStrB(lngValueEnd, vntHTTPHeader, vbCrLf & vbCrLf) + 8

        ' Armazena o final dos dados do arquivo
        lngFileDataEnd = InStrB(lngFileDataStart, vntHTTPHeader, vntDelimeter) - 4

        ' Armazena o tamanho do arquivo
        lngFileLength = lngFileDataEnd - lngFileDataStart

        ' Verifica se o arquivo est� em branco
        If lngFileLength > 2 Then
          
          ' Armazena o valor deste arquivo
          strValue = MidB(vntHTTPHeader, lngFileDataStart, lngFileLength)
          
          ' Adiciona o nome e o valor deste campo no pacote do formul�rio
          Set objUpFile = Items.Add(strName, strValue)
          objUpFile.UserFilename = strFileName
          objUpFile.TotalBytes = lngFileLength / 2
          Set objUpFile.Upload = Me
          
          If objUpFile.IsFile Then
            Set sobjFirstFile = objUpFile
            Set sobjUploadItems.FirstFile = objUpFile
          End If
          
          Set objUpFile = Nothing
          
        Else
        
          ' Adiciona o nome e o valor deste campo no pacote do formul�rio
          Items.Add strName, strFileName
          
        End If

      End If

    Else

      ' Armazena o tamanho do nome do arquivo
      lngValueStart = lngNameEnd + 8

      ' Armazena o tamanho do nome do arquivo
      lngValueEnd = InStrB(lngValueStart, vntHTTPHeader, vntDelimeter)

      ' Armazena o nome do arquivo armazenado neste campo do formul�rio
      strValue = MidB(vntHTTPHeader, lngValueStart, lngValueEnd - lngValueStart - 4)

      ' Adiciona o nome e o valor deste campo no pacote do formul�rio
      Items.Add strName, strValue

    End If

  Next

End Sub

Public Property Get GetFirstFile() As HiperTools30.UploadItem
Attribute GetFirstFile.VB_UserMemId = 0
  If Not sobjFirstFile Is Nothing Then
    Set GetFirstFile = sobjFirstFile
  Else
    Err.Raise ERR_DONT_HAVE_UPLOAD_FILE, , "N�o existe um arquivo para Upload."
  End If
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ServerName() As String
  ServerName = GetFirstFile.ServerName
End Property

Rem ===============================================================================
'Public Sub Delete( Optional ByVal strFilename As String = "" )
'Public Sub Flush ( )
'Public Sub SaveAsBlob   ( ByVal fldColumn As Field )
'Public Sub SaveInVirtual( ByVal strVirtualLocation As String )
'Public Sub TransferBlob ( ByVal fldColumn As Field )
'Public Sub TransferFile ( ByVal Filename As String )
Rem --------------------------------------------------------------------------------
