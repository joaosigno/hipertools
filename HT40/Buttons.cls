VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Buttons"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Criado (Ridai Govinda)
'Botões padrão do EditForm
Const constrAdd = "Incluir"
Const constrEdit = "Alterar"
Const constrDelete = "Excluir"
Const constrSave = "Salvar"
Const constrFind = "Localizar"
Const constrList = "Listagem"
Const constrSaveAdd = "Salvar e Incluir"
Const constrSaveCopy = "Salvar e Copiar"
Const constrFirst = " |<  "
Const constrPrevious = "  <  "
Const constrNext = "  >  "
Const constrLast = "  >| "
Const constrCancel = "Cancelar"
Const constrCopy = "Copiar"

Public FirstTitle As String
Public PreviousTitle As String
Public AddTitle As String
Public EditTitle As String
Public DeleteTitle As String
Public FindTitle As String
Public ListTitle As String
Public NextTitle As String
Public LastTitle As String
Public SaveTitle As String
Public SaveAddTitle As String
Public SaveCopyTitle As String
Public CancelTitle As String
Public CopyTitle As String

Public FirstCaption As String
Public PreviousCaption As String
Public AddCaption As String
Public EditCaption As String
Public DeleteCaption As String
Public FindCaption As String
Public ListCaption As String
Public NextCaption As String
Public LastCaption As String
Public SaveCaption As String
Public SaveAddCaption As String
Public SaveCopyCaption As String
Public CancelCaption As String
Public CopyCaption As String

Public FirstClass As String
Public PreviousClass As String
Public AddClass As String
Public EditClass As String
Public DeleteClass As String
Public FindClass As String
Public ListClass As String
Public NextClass As String
Public LastClass As String
Public SaveClass As String
Public SaveAddClass As String
Public SaveCopyClass As String
Public CancelClass As String
Public CopyClass As String

Public FirstStyleSheet As String
Public PreviousStyleSheet As String
Public AddStyleSheet As String
Public EditStyleSheet As String
Public DeleteStyleSheet As String
Public FindStyleSheet As String
Public ListStyleSheet As String
Public NextStyleSheet As String
Public LastStyleSheet As String
Public SaveStyleSheet As String
Public SaveAddStyleSheet As String
Public SaveCopyStyleSheet As String
Public CancelStyleSheet As String
Public CopyStyleSheet As String

Public FirstAccessKey As String
Public PreviousAccessKey As String
Public AddAccessKey As String
Public EditAccessKey As String
Public DeleteAccessKey As String
Public FindAccessKey As String
Public ListAccessKey As String
Public NextAccessKey As String
Public LastAccessKey As String
Public SaveAccessKey As String
Public SaveAddAccessKey As String
Public SaveCopyAccessKey As String
Public CancelAccessKey As String
Public CopyAccessKey As String

Public FirstImage As String
Public PreviousImage As String
Public AddImage As String
Public EditImage As String
Public DeleteImage As String
Public FindImage As String
Public ListImage As String
Public NextImage As String
Public LastImage As String
Public SaveImage As String
Public SaveAddImage As String
Public SaveCopyImage As String
Public CancelImage As String
Public CopyImage As String
'---------------------

Public ShowFirst As Boolean
Public ShowPrevious As Boolean
Public ShowAdd As Boolean
Public ShowCopy As Boolean
Public ShowEdit As Boolean
Public ShowDelete As Boolean
Public ShowFind As Boolean
Public ShowList As Boolean
Public ShowNext As Boolean
Public ShowLast As Boolean
Public ShowSave As Boolean
Public ShowSaveAdd As Boolean
Public ShowSaveCopy As Boolean
Public ShowCancel As Boolean

Private sobjButtons As Collection

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  ShowFirst = True
  ShowPrevious = True
  ShowAdd = True
  ShowCopy = True
  ShowEdit = True
  ShowDelete = True
  ShowFind = True
  ShowList = True
  ShowNext = True
  ShowLast = True
  ShowSave = True
  ShowSaveAdd = True
  ShowSaveCopy = True
  ShowCancel = True

  AddCaption = constrAdd
  CancelCaption = constrCancel
  DeleteCaption = constrDelete
  EditCaption = constrEdit
  FindCaption = constrFind
  FirstCaption = constrFirst
  LastCaption = constrLast
  ListCaption = constrList
  NextCaption = constrNext
  PreviousCaption = constrPrevious
  SaveAddCaption = constrSaveAdd
  SaveCaption = constrSave
  SaveCopyCaption = constrSaveCopy
  CopyCaption = constrCopy
  
  FirstTitle = "Primeiro Registro"
  PreviousTitle = "Registro Anterior"
  AddTitle = "Inclui um Registro"
  EditTitle = "Edita o registro corrente"
  DeleteTitle = "Exclui o registro corrente"
  FindTitle = "Entra em modo de localização de registros"
  ListTitle = "Entra no modo de visualização dos registros em lista"
  NextTitle = "Próximo Registro"
  LastTitle = "Último Registro"
  SaveTitle = "Salva as alterações efetuadas"
  SaveAddTitle = "Salva as alterações efetuadas e entra em modo de inclusão"
  SaveCopyTitle = "Salva as alterações efetuadas e entra em modo de inclusão copiando os dados do registro corrente"
  CancelTitle = "Cancela todas as alterações efetuadas"
  CopyTitle = "Inclui um registro copiando os dados do registro corrente"
  
  FirstAccessKey = ""
  PreviousAccessKey = ""
  AddAccessKey = "I"
  EditAccessKey = "A"
  DeleteAccessKey = "P"
  FindAccessKey = ""
  ListAccessKey = ""
  NextAccessKey = ""
  LastAccessKey = ""
  SaveAccessKey = "S"
  SaveAddAccessKey = ""
  SaveCopyAccessKey = ""
  CancelAccessKey = "C"
  CopyAccessKey = ""
  
  FirstStyleSheet = ""
  PreviousStyleSheet = ""
  AddStyleSheet = ""
  EditStyleSheet = ""
  DeleteStyleSheet = ""
  FindStyleSheet = ""
  ListStyleSheet = ""
  NextStyleSheet = ""
  LastStyleSheet = ""
  SaveStyleSheet = ""
  SaveAddStyleSheet = ""
  SaveCopyStyleSheet = ""
  CancelStyleSheet = ""
  CopyStyleSheet = ""
  
  FirstClass = ""
  PreviousClass = ""
  AddClass = ""
  EditClass = ""
  DeleteClass = ""
  FindClass = ""
  ListClass = ""
  NextClass = ""
  LastClass = ""
  SaveClass = ""
  SaveAddClass = ""
  SaveCopyClass = ""
  CancelClass = ""
  CopyClass = ""
  
  FirstImage = ""
  PreviousImage = ""
  AddImage = ""
  EditImage = ""
  DeleteImage = ""
  FindImage = ""
  ListImage = ""
  NextImage = ""
  LastImage = ""
  SaveImage = ""
  SaveAddImage = ""
  SaveCopyImage = ""
  CancelImage = ""
  CopyImage = ""
End Sub

Private Sub Class_Initialize()
    Set sobjButtons = New Collection
End Sub

Private Sub Class_Terminate()
    Set sobjButtons = Nothing
End Sub

Public Function Add(ByVal strName As String, ByVal strImage As String, _
                    ByVal strAccessKey As String, ByVal strTitle As String, _
                    ByVal strCaption As String, ByVal blnShow As Boolean, _
                    ByVal strCSS As String) As HiperTools30.Button
    Dim sobjButton As HiperTools30.Button
    Set sobjButton = New HiperTools30.Button

    sobjButton.Name = strName
    sobjButton.Image = strImage
    sobjButton.AccessKey = strAccessKey
    sobjButton.Title = strTitle
    sobjButton.Caption = strCaption
    sobjButton.StyleSheet = strCSS
    sobjButton.Show = blnShow
    sobjButtons.Add sobjButton, strName

    Set Add = sobjButton
    Set sobjButton = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As HiperTools30.Button
  Set Item = sobjButtons(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = sobjButtons.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    sobjButtons.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = sobjButtons.[_NewEnum]
End Property
