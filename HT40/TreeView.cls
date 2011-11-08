VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "TreeView"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Rem ==================================================================================================================================================
Rem  /TreeView.cls
Rem --------------------------------------------------------------------------------------------------------------------------------------------------
Rem  Descricao   : Objeto gerador de TreeView.
Rem  Criação     : 22/12/99 07:56:00 PM
Rem  Local       : Brasília/DF
Rem  Elaborado   : Kelson da Costa Medeiros <kelson@zevallos.com.br>
Rem              : Ruben Zevallos Jr. <ruben@zevallos.com.br>
Rem  Versao      : 1.4.2
Rem  Copyright   : 1999 by Zevallos® Tecnologia em Informação
Rem --------------------------------------------------------------------------------------------------------------------------------------------------

' Constantes do ambiente
'-----------------------

' Sufixo e extensão padrão para o nome das imagens
Const conOpened = "Open"
Const conClosed = "Close"

' Constantes com os nomes das imagens padrão
Const conLine = "/img/folder/Line.gif"
Const conLineBegin = "/img/folder/LineBegin.gif"
Const conLineBeginMinus = "/img/folder/LineBeginMinus.gif"
Const conLineBeginPlus = "/img/folder/LineBeginPlus.gif"
Const conLineEnd = "/img/folder/LineEnd.gif"
Const conLineEndMinus = "/img/folder/LineEndMinus.gif"
Const conLineEndPlus = "/img/folder/LineEndPlus.gif"
Const conLineMiddle = "/img/folder/LineMiddle.gif"
Const conLineMiddleOnly = "/img/folder/LineMiddleOnly.gif"
Const conLineMiddleMinus = "/img/folder/LineMiddleMinus.gif"
Const conLineMiddlePlus = "/img/folder/LineMiddlePlus.gif"
Const conSpace = "/img/folder/Space.gif"

' Constantes com os nomes "QueryStrings"
Const conItemClicked = "IC"
Const conItemState = "NS"
Const conCurrentLevel = "CL"
'Const conLevelID = "LI"

' Constantes com os nomes "QueryStrings"
Const conStateOpen = "O"
Const conStateClose = "C"

' Constantes con os nomes dos tipos de itens
Const conStatic = "Static"
Const conLookUp = "LookUp"

' Constantes com os prefixos dos nomes das propriedades do item
Const conItemType = "ItemType"
Const conItemParent = "ItemParent"
Const conItemIdent = "ItemIdent"
Const conItemText = "ItemText"
Const conItemURL = "ItemURL"
Const conItemOpened = "ItemOpened"
Const conItemIsNode = "ItemIsNode"
Const conItemLast = "ItemLast"
Const conItemChild = "ItemChild"
Const conItemChildCount = "ItemChildCount"
Const conItemImage = "ItemImage"
Const conItemBGColor = "ItemBGColor"
Const conItemFontFace = "ItemFontFace"
Const conItemFontColor = "ItemFontColor"
Const conItemFontSize = "ItemFontSize"
Const conItemBold = "ItemBold"
Const conItemItalic = "ItemItalic"
Const conItemUnderline = "ItemUnderline"
Const conItemHAlign = "ItemHAlign"
Const conItemVAlign = "ItemVAlign"
Const conItemTarget = "ItemTarget"
Const conItemNoWrap = "ItemNoWrap"
Const conItemHeight = "ItemHeight"
Const conItemWidth = "ItemWidth"
Const conItemTitle = "ItemTitle"
Const conItemMaxSize = "ItemMaxSize"
Const conItemStyle = "ItemStyle"
Const conItemCurrentLevel = "ItemCurrentLevel"
Const conItemIsLevelID = "IsLevelID"
Const conItemColumn = "ItemColumn"
Const conItemColumnCount = "ItemColumnCount"
Const conItemSub = "ItemSub"
Const conItemSubCount = "ItemSubCount"
Const conItemHaveMore = "ItemHaveMore"

' Variaveis globais privadas
' --------------------------------------------------------
Private sintItemsCount As Integer ' Quantidade de itens adicionados
Private sintImagesCount As Integer ' Quantidade de imagens adicionadas
Private sintMaxNodes As Integer ' Quantidade máxima de nós adicionados
Private sstrItemClicked As String ' Armazenam os valos dos "QueryStrings"
Private sstrItemState As String ' Armazenam os valos dos "QueryStrings"
Private sSetLevelID As String ' Armazena a identificação personalizada apenas do próximo item
Private sintChildsInRoot As Integer ' Quantidade de filhos na raíz
Private sarrCountChilds(100) As Integer ' Array com a quantidade de filhos para cada item aberto
Private arrLevelID(40) As String ' Array dos Levels ID.
Private sdicNodes As Object ' Pacote com os nomes dos nós
Private sdicItems As Object ' Pacote com as propriedades dos itens adicionados
Private sdicImages As Object ' Pacote com as propriedades das imagens adicionados

Private Table As HiperTools30.Table, _
        URL As HiperTools30.URL, _
        Initializer As HiperTools30.Initializer, _
        Show As HiperTools30.Show, _
        Strings As HiperTools30.String, _
        ScriptContext As ScriptingContext

Private sstrComments As String, _
        sstrTitle As String, _
        sstrVersion As String, _
        sstrCompanyName As String, _
        sstrProductName As String, _
        sstrLegalCopyright As String

Private Request As Request, _
        Response As Response, _
        Server As Server, _
        Session As Session, _
        Application As Application


' Propriedades
' --------------------------------------------------------

' Define qual é a imagem padrão a ser utilizada quando o usuário não informar o nome chave da imagem
Public DefaultImage As String

' Define qual é a cor de fundo para os próximos itens adicionados
Public ItemBGColor As String

' Define qual é a cor da fonte para os próximos itens adicionados
Public ItemFontColor As String

' Define qual é o nome da fonte para os próximos itens adicionados
Public ItemFontFace As String

' Define qual é o tamanho da fonte para os próximos itens adicionados
Public ItemFontSize As String

' Define se os próximos itens adicionados aparecerão com negrito
Public ItemBold As Boolean

' Define se os próximos itens adicionados aparecerão com itálico
Public ItemItalic As Boolean

' Define se os próximos itens adicionados aparecerão com subscrito
Public ItemUnderline As Boolean

' Define se o texto do item poderá ou não aparecer quebrado em multiplas linhas
Public ItemNoWrap As Boolean

' Define qual é o alinhamento horizontal para os próximos itens adicionados
Public ItemHAlign As String

' Define qual é o alinhamento vertical para os próximos itens adicionados
Public ItemVAlign As String

' Define qual é o alvo para os  próximos itens adicionados
Public ItemTarget As String

' Define qual é a largura para os próximos itens adicionados
Public ItemWidth As String

' Define qual é a altura para os próximos itens adicionados
Public ItemHeight As String

' Define qual é a altura para os próximos itens adicionados
Public TreeBGColor As String

' Define qual é o tamanho máximo de apresentação dos texto em número de caracteres para os próximos itens adicionados
Public ItemMaxSize As Integer

' Define qual é o estilo que deverá ser utilizado para os próximos itens adicionados
Public ItemStyle As String

' Define qual é o estilo do "Link" que deverá ser utilizado para todos os itens adicionados
Public TreeStyleLink As String

' Define qual é o estilo do "Active" que deverá ser utilizado para todos os itens adicionados
Public TreeStyleActive As String

' Define qual é o estilo do "Visited" que deverá ser utilizado para todos os itens adicionados
Public TreeStyleVisited As String

' Define qual é o estilo do "Hover" que deverá ser utilizado para todos os itens adicionados
Public TreeStyleHover As String

' Esta propriedade retorna a posição atual dos níveis
Public CurrentChild As Integer

' Esta propriedade recupera a identificação personalizada apenas do item atual
Public GetLevelID As String

' Esta propriedade retorna True se o strSetLevelID é o nível atual.
Public IsLevelID As String

' Propriedade que indica se o item possui mais itens mesmo que não tenham sido adicionados
Public ItemHaveMore As Boolean

' Propriedade que indica se o item utilizará as cores dos parentes
Public UseParentsColors As Boolean

Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  Dim i As Integer, arrTLevelID

  TimeBombX xtbTreeView, 15656, 12934, 3080, 14624

  Set ScriptContext = ScriptContext
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application
  Set Table = Server.CreateObject("HiperTools30.Table")
  Set URL = Server.CreateObject("HiperTools30.URL")
  Set Initializer = Server.CreateObject("HiperTools30.Initializer")
  Set Show = Server.CreateObject("HiperTools30.Show")
  Set Strings = Server.CreateObject("HiperTools30.String")
  Set sdicNodes = Server.CreateObject("Scripting.Dictionary")
  Set sdicItems = Server.CreateObject("Scripting.Dictionary")
  Set sdicImages = Server.CreateObject("Scripting.Dictionary")

  ' Configuração dos "QueryStrings"
  ' -------------------------------

  ' Armazena o valor do "QueryString" "conItemClicked", tipo de ação a ser executada
  sstrItemClicked = Request.QueryString(conItemClicked)

  ' Armazena o valor do "QueryString" "conItemState", item que executou a ação
  sstrItemState = Request.QueryString(conItemState)

  arrTLevelID = Split(sstrItemClicked, "-")

  For i = 0 To UBound(arrTLevelID)
    arrLevelID(i) = Trim(arrTLevelID(i))
  Next

  CurrentChild = 0

  sintChildsInRoot = 0

End Sub

Public Sub OnEndPage()

  Set Table = Nothing
  Set URL = Nothing
  Set Initializer = Nothing
  Set Show = Nothing
  Set Strings = Nothing
  Set sdicNodes = Nothing
  Set sdicItems = Nothing
  Set sdicImages = Nothing
  Set ScriptContext = Nothing
  Set Request = Nothing
  Set Response = Nothing
  Set Server = Nothing
  Set Session = Nothing
  Set Application = Nothing


End Sub

Private Sub Class_Initialize()

    sstrComments = App.Comments
    sstrTitle = App.Title
    sstrVersion = App.Major & "." & _
                App.Minor & "." & _
                App.Revision
    sstrCompanyName = App.CompanyName
    sstrProductName = App.ProductName

End Sub

Public Property Get Comments() As String
    Comments = sstrComments
End Property

Public Property Get Title() As String
    Title = sstrTitle
End Property

Public Property Get Version() As String
    Version = sstrVersion
End Property
Public Property Get CompanyName() As String
    CompanyName = sstrCompanyName
End Property

Public Property Get ProductName() As String
    ProductName = sstrProductName
End Property

Public Property Get LegalCopyright() As String
    LegalCopyright = sstrLegalCopyright
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ItemClicked() As String
    ItemClicked = sstrItemClicked
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Get ItemState() As String
    ItemState = sstrItemState
End Property

Rem =========================================================================
Rem
Rem -------------------------------------------------------------------------
Public Property Let SetLevelID(ByVal strSetLevelID As String)

  sSetLevelID = strSetLevelID

  IsLevelID = False

  If strSetLevelID > "" Then
    IsLevelID = strSetLevelID = Trim(arrLevelID(CurrentChild))

  End If

End Property

Rem =========================================================================
Rem Inicializa a definição do "TreeView"
Rem -------------------------------------------------------------------------
Public Sub BeginTree()

  ' Inicializa os contadores
  sintItemsCount = 0
  sintImagesCount = 0
  sintMaxNodes = 0

End Sub

Rem =========================================================================
Rem Adiciona um nó a este "TreeView".
Rem -------------------------------------------------------------------------
Public Sub BeginNode(ByVal strText As String, _
                     Optional ByVal strURL As String = "", _
                     Optional ByVal strItemTitle As String = "", _
                     Optional ByVal strImage As String = "", _
                     Optional ByVal blnOpened As Boolean = False)

  ' Adiciona o item no pacote
  AddStatic strText, strURL, blnOpened, True, strImage, strItemTitle

  ' Adiciona este item na lista de nós abertos
  sdicNodes.Add conItemParent & sdicNodes.Count + 1, sintItemsCount

  ' Se a quantidade de nós para o próximo item for maior que a quantidade máxima de nós
  If sdicNodes.Count > sintMaxNodes Then

     ' Armazena a nova quantidade máxima de nós
     sintMaxNodes = sdicNodes.Count

  End If

  CurrentChild = CurrentChild + 1

  GetLevelID = arrLevelID(CurrentChild - 1)

  IsLevelID = False

  If sSetLevelID > "" Then
    IsLevelID = sSetLevelID = GetLevelID

  End If

  ' Limpa o valor da propriedade "CurrentLevel"
  sSetLevelID = ""

End Sub

Rem =========================================================================
Rem Adicina um item estático ao "TreeView".
Rem -------------------------------------------------------------------------
Public Sub EndNode()

  ' Remove o último item da lista de nós abertos
  sdicNodes.Remove (conItemParent & sdicNodes.Count)

  CurrentChild = CurrentChild - 1

  GetLevelID = arrLevelID(CurrentChild)

  IsLevelID = False

End Sub

Rem =========================================================================
Rem Adiciona um item a este "TreeView".
Rem -------------------------------------------------------------------------
Public Sub AddItem(ByVal strText As String, _
                   Optional ByVal strURL As String = "", _
                   Optional ByVal strItemTitle As String = "", _
                   Optional ByVal strImage As String = "")

  ' Adiciona o item no pacote
  AddStatic strText, strURL, False, False, strImage, strItemTitle

End Sub

Rem =========================================================================
Rem Adiciona um sub-item a este "TreeView".
Rem -------------------------------------------------------------------------
Public Sub AddSubCaption(ByVal strText As String, _
                         Optional ByVal strURL As String = "", _
                         Optional ByVal strTitle As String = "", _
                         Optional ByVal strTarget As String = "")
 
  ' Variável temporária
  Dim strSufix As String

  ' Se já foram adicionados itens
  If sintItemsCount > 0 Then

     ' Incrementa a quantidade de colunas adicionadas neste item
     sdicItems(conItemSubCount & sintItemsCount) = sdicItems(conItemSubCount & sintItemsCount) + 1

     ' Escreve o nome deste item
     strSufix = conItemSub & sdicItems(conItemSubCount & sintItemsCount)
  
     ' Adiciona as propriedades padronizadas da coluna
     AddDefaultProperties strText, strURL, strTitle, strTarget, strSufix
  
  End If

End Sub

Rem =========================================================================
Rem Adiciona uma coluna ao último item. Não será aplicado aos próximos itens.
Rem -------------------------------------------------------------------------
Public Sub AddColumn(ByVal strText As String, _
                     Optional ByVal strURL As String = "", _
                     Optional ByVal strTitle As String = "", _
                     Optional ByVal strTarget As String = "")

  Dim strSufix As String

  ' Se já foram adicionados itens
  If sintItemsCount > 0 Then

     ' Incrementa a quantidade de colunas adicionadas neste item
     sdicItems(conItemColumnCount & sintItemsCount) = sdicItems(conItemColumnCount & sintItemsCount) + 1

     ' Escreve o nome deste item
     strSufix = conItemColumn & sdicItems(conItemColumnCount & sintItemsCount)
  
     ' Adiciona as propriedades padronizadas da coluna
     AddDefaultProperties strText, strURL, strTitle, strTarget, strSufix

     ' Se a identação mais a quantidade de colunas deste item for maior que a quantidade máxima de nós
     If sdicItems(conItemColumnCount & sintItemsCount) + sdicItems(conItemIdent & sintItemsCount) > sintMaxNodes Then

        ' Armazena a nova quantidade máxima de nós
        sintMaxNodes = sdicItems(conItemColumnCount & sintItemsCount) + sdicItems(conItemIdent & sintItemsCount)

     End If

  End If

End Sub

Rem =========================================================================
Rem Adiciona uma imagem à lista de imagens disponíveis aos itens. Desta forma
Rem ao se adicionar o nome os itens não é necessário informar o nome da
Rem imagem completo novamente, bastando informar o nome chave da imagem
Rem -------------------------------------------------------------------------
Public Sub AddImage(ByVal strInitial As String, _
                    ByVal strAddress As String, _
                    Optional ByVal strOpen As String = "Open", _
                    Optional ByVal strClose As String = "Close", _
                    Optional ByVal strExtension As String = ".gif")

  ' Se o usuário tentar adicionar uma imagem que já exista
  If sdicImages.Exists(strInitial & conOpened) Then

     ' Informa ao mesmo que esta imagem já existe
     Show.Message "Já foi adicionada uma imagem com o nome """ & strInitial & """."

     ' Não continua com a execução do procedimento
     Exit Sub

  End If
  
  ' Verifica se já existe o ponto antes da extenção
  If Left(strExtension, 1) <> "." Then
     strExtension = "." & strExtension
  End If

  ' Incrementa o contador de items adicionados
  sintImagesCount = (sintImagesCount + 1)

  ' Adiciona no pacote as propriedades deste item
  sdicImages.Add strInitial & conOpened, strAddress & strInitial & strOpen & strExtension
  sdicImages.Add strInitial & conClosed, strAddress & strInitial & strClose & strExtension

End Sub

Rem =========================================================================
Rem Faz o Igual para parametros
Rem -------------------------------------------------------------------------
Public Function Equal(ByVal strOption, _
                      ByVal strValue) As String

  ' Retorno da função
  Equal = "&" & strOption & "=" & strValue

End Function

Rem =========================================================================
Rem Finaliza o "TreeView" e inicia o código que o constrói.
Rem -------------------------------------------------------------------------
Public Sub EndTree()

  ' Prepara os items para depois incluílos.
  PrepareTree

  ' Mostra todos os items do "TreeView".
  MakeTree

  ' Inicializa as propriedades
  InitializeProperties

End Sub

Rem =========================================================================
Rem Adicina um item estático ao "TreeView".
Rem -------------------------------------------------------------------------
Private Sub AddStatic(ByVal strText As String, _
                      ByVal strURL As String, _
                      ByVal blnOpened As Boolean, _
                      ByVal bnlIsNode As Boolean, _
                      ByVal strImage As String, _
                      ByVal strItemTitle As String)

  ' Variável para controle do "Loop" "For"
  Dim i, intCurrentLevel As Integer
  Dim strParentCurrentLevel As String

  ' Incrementa o contador de items adicionados
  sintItemsCount = sintItemsCount + 1

  ' Adiciona as propriedades padronizadas deste item
  AddDefaultProperties strText, strURL, strItemTitle, ItemTarget, ""
  
  ' Adiciono as propriedades específicas do item
  sdicItems.Add conItemIdent & sintItemsCount, sdicNodes.Count
  sdicItems.Add conItemIsNode & sintItemsCount, bnlIsNode
  sdicItems.Add conItemOpened & sintItemsCount, blnOpened
  sdicItems.Add conItemChild & sintItemsCount, False
  sdicItems.Add conItemChildCount & sintItemsCount, 0
  sdicItems.Add conItemImage & sintItemsCount, strImage
  sdicItems.Add conItemColumnCount & sintItemsCount, 0
  sdicItems.Add conItemSubCount & sintItemsCount, 0
  sdicItems.Add conItemHaveMore & sintItemsCount, ItemHaveMore

  ' Adiciona no pacote as propriedades deste item
  sdicItems.Add conItemType & sintItemsCount, conStatic
  If sdicNodes.Count > 0 Then
     sdicItems.Add conItemParent & sintItemsCount, sdicNodes(conItemParent & sdicNodes.Count)
     sdicItems(conItemChildCount & sdicItems(conItemParent & sintItemsCount)) = sdicItems(conItemChildCount & sdicItems(conItemParent & sintItemsCount)) + 1
     intCurrentLevel = sdicItems(conItemChildCount & sdicItems(conItemParent & sintItemsCount))
  Else
     sdicItems.Add conItemParent & sintItemsCount, 0
     sintChildsInRoot = sintChildsInRoot + 1
     intCurrentLevel = sintChildsInRoot
  End If

  ' Se este item possuí um pai, armazena o "CurrenLevel" do item pai
  If sdicItems(conItemParent & sintItemsCount) > 0 Then
     strParentCurrentLevel = sdicItems(conItemCurrentLevel & sdicItems(conItemParent & sintItemsCount))
  End If

  ' Se o usuário estiver configurando um "CurrentLevel" deste cara
  If sSetLevelID > "" Then

     ' Armazena no "CurrentLevel" deste cara o "CurrentLevel" do pai mais o "CurrentLevel" informado pelo usuário
     sdicItems.Add conItemCurrentLevel & sintItemsCount, strParentCurrentLevel & sSetLevelID & "-"

  Else

     ' Armazena no "CurrentLevel" deste cara o "CurrentLevel" do pai mais o "CurrentLevel" calculado automáticamente
     sdicItems.Add conItemCurrentLevel & sintItemsCount, sdicItems(conItemCurrentLevel & sdicItems(conItemParent & sintItemsCount)) & intCurrentLevel & "-"

  End If

  ' Informa se este cara é um "LevelID"
  sdicItems.Add conItemIsLevelID & sintItemsCount, IsLevelID

  ' Informa ao pai deste item que ele agora tem filhos
  sdicItems(conItemChild & sdicItems(conItemParent & sintItemsCount)) = True

  ' Percorre todos os item
  For i = 1 To sintItemsCount - 1

    ' Verifica qual tem o mesmo pai do item que eu vou adicionar agora
    If sdicItems(conItemParent & i) = sdicItems(conItemParent & sintItemsCount) Then

       ' Indica que ele já não é mais o último
       sdicItems(conItemLast & i) = False

    End If

  Next

  ' Informa que este novo item é o último deste pai
  sdicItems.Add conItemLast & sintItemsCount, True

End Sub

Rem =========================================================================
Rem Adicina as propriedades padronizadas a todos os itens do "TreeView".
Rem -------------------------------------------------------------------------
Private Sub AddDefaultProperties(ByVal strText As String, _
                                 ByVal strURL As String, _
                                 ByVal strTitle As String, _
                                 ByVal strTarget As String, _
                                 ByVal strSufix As String)

  ' Variavel temporária
  Dim strThisItem As String
  
  ' Escreve o nome a ser aplicado a todas as propriedades
  strThisItem = sintItemsCount & strSufix

  ' Adiciona as propriedades
  sdicItems.Add conItemText & strThisItem, strText
  sdicItems.Add conItemURL & strThisItem, strURL
  sdicItems.Add conItemTitle & strThisItem, strTitle
  sdicItems.Add conItemTarget & strThisItem, strTarget
  sdicItems.Add conItemBGColor & strThisItem, ItemBGColor
  sdicItems.Add conItemFontFace & strThisItem, ItemFontFace
  sdicItems.Add conItemFontColor & strThisItem, ItemFontColor
  sdicItems.Add conItemFontSize & strThisItem, ItemFontSize
  sdicItems.Add conItemBold & strThisItem, ItemBold
  sdicItems.Add conItemItalic & strThisItem, ItemItalic
  sdicItems.Add conItemUnderline & strThisItem, ItemUnderline
  sdicItems.Add conItemHAlign & strThisItem, ItemHAlign
  sdicItems.Add conItemVAlign & strThisItem, ItemVAlign
  sdicItems.Add conItemNoWrap & strThisItem, ItemNoWrap
  sdicItems.Add conItemWidth & strThisItem, ItemWidth
  sdicItems.Add conItemHeight & strThisItem, ItemHeight
  sdicItems.Add conItemMaxSize & strThisItem, ItemMaxSize
  sdicItems.Add conItemStyle & strThisItem, ItemStyle

End Sub

Rem =========================================================================
Rem Prepara os items para depois incluílos.
Rem -------------------------------------------------------------------------
Private Sub PrepareTree()

  ' Variável para controle do "Loop" "For"
  Dim i, j

  ' Percorre todos os item
  For i = 1 To sintItemsCount

    ' Verifica se existe um nó sem filho
    If (sdicItems(conItemIsNode & i)) And (Not sdicItems(conItemChild & i)) Then

       ' Já que este nó não tem filho, tornarei este nó fechado
       sdicItems(conItemOpened & i) = False

    End If


    ' Se este é o item clicado
    If Strings.CropRight(sdicItems(conItemCurrentLevel & i), 1) = sstrItemClicked Then

       ' Copia o índice deste item para a variável "j"
       j = i

       ' Percorro todos os pais deste item
       Do While sdicItems(conItemParent & j) > 0

          ' Informa que este item está sendo fechado
          sdicItems(conItemOpened & j) = True

          ' Informa que este item está aberto
          sdicItems(conItemOpened & sdicItems(conItemParent & j)) = True

          ' Aponta o índice para o pai deste item
          j = sdicItems(conItemParent & j)

       Loop

       ' Se o usuário esta abrindo este item
       If sstrItemState = conStateOpen Then

          ' Informa que este item está sendo fechado
          sdicItems(conItemOpened & i) = True

       ' Se o usuário está fechando este item
       Else

          ' Informa que este item está sendo fechado
          sdicItems(conItemOpened & i) = False

       End If

    End If

  Next

End Sub

Rem =========================================================================
Rem Mostra todos os items do "TreeView".
Rem -------------------------------------------------------------------------
Private Sub MakeTree()

  ' Variável para controle do "Loop" "For"
  Dim intIndex

  ' Retirando o espaço entre as celulas da tabela
  Table.Padding = 0
  Table.Spacing = 0

  ' Configurando a cor de fundo do "TreeView"
  Table.Style.Color2 = TreeBGColor

  ' Se alguma das propriedades de estilo do TreeView foi configurada
  If (TreeStyleLink & TreeStyleActive & TreeStyleVisited & TreeStyleHover > "") Then

     ' Começa a escrever o estilo
     Show.HTMLCR "<STYLE>"

     ' Se o usuário configurou a propriedade "Link", escreve o estilo "Link"
     Show.StyleClass "A:LINK", TreeStyleLink

     ' Se o usuário configurou a propriedade "Active", escreve o estilo "Active"
     Show.StyleClass "A:ACTIVE", TreeStyleActive

     ' Se o usuário configurou a propriedade "Visited", escreve o estilo "Visited"
     Show.StyleClass "A:VISITED", TreeStyleVisited

     ' Se o usuário configurou a propriedade "Hover", escreve o estilo "Hover"
     Show.StyleClass "A:HOVER", TreeStyleHover

     ' Finalizo a escrita do estilo
     Show.HTMLCR "</STYLE>"

  End If

  ' Inicializa uma nova tabela onde será formatado o "TreeView"
  Table.BeginTable "100%"

  ' Percorre todos os itens do "TreeView" para construílo
  For intIndex = 1 To sintItemsCount

      ' Se o item for do tipo "Static"
      If sdicItems(conItemType & intIndex) = conStatic Then

         ' Desenha um item estático do "TreeView".
         ShowStaticItem (intIndex)

      ' Se o item for do tipo "LookUp"
      Else

        ' Por enquanto não faz nada

      End If

  Next

  ' Finaliza a tabela utilizada na construção do "TreeView"
  Table.EndTable

End Sub

Rem =========================================================================
Rem Desenha um item estático do "TreeView".
Rem -------------------------------------------------------------------------
Private Sub ShowStaticItem(ByVal intIndex As Integer)

  ' Se o pai não estiver aberto
  If Not sdicItems(conItemOpened & sdicItems(conItemParent & intIndex)) And (sdicItems(conItemParent & intIndex) > 0) Then

     ' Informa que este item está sendo fechado
     sdicItems(conItemOpened & intIndex) = False

     ' Caio fora do procedimento
     Exit Sub

  End If

  ' Adiciona um novo item, ou seja, inicializa uma nova linha na tabela
  Table.BeginRow

    ' Desenha a indentação do item do "TreeView".
    ShowIdent (intIndex)

    ' Mostra a image que indica o estado do item
    ShowStateImage (intIndex)

    ' Mostra o ícone do item
    ShowIcon (intIndex)

    ' Mostra o título do item do "TreeView".
    ShowCaption (intIndex)

  ' Finaliza a linha de desenho do item
  Table.EndRow

End Sub

Rem =========================================================================
Rem Desenha a indentação do item do "TreeView".
Rem -------------------------------------------------------------------------
Private Sub ShowIdent(ByVal intIndex As Integer)

  ' Variáveis para controle dos "Loop"
  Dim i As Integer, j As Integer, k As Integer

  ' Pacote que armazena o nomes das imagens que deverão ser mostradas
  Dim dicParents As Object

  'Variável que armazena o caminho para o diretório virtual onde se encontra
  'os arquivos do HiperTools
  Dim strPath As String
  strPath = Session("htDefaultHiperToolsWorkPath")

  ' Define o tamanho das colunas anteriores à que aparece o texto para que fique com tamanho igua a 1%
  Table.ColumnWidth = "1%"

  ' Copia o índice deste item para a variável "j"
  j = intIndex

  ' Inicializa a variável será o que armazenará o valor
  i = 1
  
  ' Contador de identação
  k = 0

  ' Iniciliza o pacote
  Set dicParents = Server.CreateObject("Scripting.Dictionary")

  ' Percorre todos os itens pai deste item verificando a necessidade de desenhar a linha
  ' de continuação
  Do While sdicItems(conItemParent & j) > 0

     ' Incrementa está variável
     k = k + 1

     ' Verifica se este item pai é o último item do nó
     If sdicItems(conItemLast & sdicItems(conItemParent & j)) Then

        ' Já que este item é o último item deste nó, não desenha a linha de continuação
        dicParents.Add "Image" & k, strPath & conSpace

     ' Já que este item ainda não é o último item deste nó
     Else

        ' Desenha a linha de continuação
        dicParents.Add "Image" & k, strPath & conLine

     End If
     
     dicParents.Add "Color" & k, sdicItems(conItemBGColor & sdicItems(conItemParent & j))

     ' Aponta o índice para o pai deste item
     j = sdicItems(conItemParent & j)

  Loop

  ' Percorre o pacote com os nomes dos objetos de traz para frente, desta forma
  For i = k To 1 Step -1

      ' Se for para utilizar as cores do ancestral
      If (UseParentsColors) Then

         ' Utiliza a cor deste ancestral
         Table.ColumnColor = dicParents("Color" & i)

      Else

         ' Utiliza a cor deste item
         Table.ColumnColor = sdicItems(conItemBGColor & intIndex)

      End If
      
      ' Adiciona a coluna responsável pelo incremento
      Table.Column "<img src=""" & dicParents("Image" & i) & """ border=0>"

      ' Devolvendo a cor padrão
      Table.ColumnColor = ""

  Next

  ' Destrói o pacote
  Set dicParents = Nothing

End Sub

Rem =========================================================================
Rem Desenha o estado do item do "TreeView".
Rem -------------------------------------------------------------------------
Private Sub ShowStateImage(ByVal intIndex As Integer)

  'Variável que armazena o caminho para o diretório virtual onde se encontra
  'os arquivos do HiperTools
  Dim strPath As String
  strPath = Session("htDefaultHiperToolsWorkPath")

  ' Utiliza a cor deste item
  Table.ColumnColor = sdicItems(conItemBGColor & intIndex)

  ' Inicializa a coluna onde será mostrado o estado do item
  Table.BeginColumn

    ' Se este item for um nó
    If ((sdicItems(conItemIsNode & intIndex)) And (sdicItems(conItemChild & intIndex) Or sdicItems(conItemHaveMore & intIndex))) Then

       ' Começa a construção do link que irá me reinvocar executando a ação a mim relacionada
       URL.BeginURL Initializer.ScriptName

         ' Adiciona no link o ponteiro para o item clicado
         URL.Add conItemClicked, Strings.CropRight(sdicItems(conItemCurrentLevel & intIndex), 1)

         ' Se este item estiver aberto
         If sdicItems(conItemOpened & intIndex) Then

            ' Informa no link o novo estado deste item
            URL.Add conItemState, conStateClose

            ' Se este for o último item da arvore
            If sdicItems(conItemLast & intIndex) Then

               ' Adiciono a imagem [-] sem continuação de linha
               URL.Show "<img src=""" & strPath & conLineEndMinus & """ border=0>"

            ' Se este item estiver fechado
            Else

               ' Adiciono a imagem [-] com continuação de linha
               URL.Show "<img src=""" & strPath & conLineMiddleMinus & """ border=0>"

            End If

         ' Se este não for o último item da lista
         Else

            ' Informa no link o novo estado deste item
            URL.Add conItemState, conStateOpen

            ' Se este for o último item da arvore
            If sdicItems(conItemLast & intIndex) Then

               ' Adiciono a imagem [+] sem continuação de linha
               URL.Show "<img src=""" & strPath & conLineEndPlus & """ border=0>"

            ' Se este item estiver fechado
            Else

               ' Adiciono a imagem [+] com continuação de linha
               URL.Show "<img src=""" & strPath & conLineMiddlePlus & """ border=0>"

            End If

         End If

      ' Finaliza a construção da URL
      URL.EndURL

    ' Se este for um item comum
    Else

       ' Se este for o primeiro item
       If (intIndex = 1) Then

          ' Se este for o último item ligado à raiz
          If sdicItems(conItemLast & intIndex) Then

             ' Adiciono uma coluna onde ficará apenas o traço no meio com continuação de linha
             Show.Image strPath & conLineMiddleOnly
             
          ' Se não
          Else

             ' Adiciono uma coluna onde ficará apenas o traço no meio com continuação de linha
             Show.Image strPath & conLineBegin
          
          End If

       ' Se este for um item do meio da lista
       ElseIf Not sdicItems(conItemLast & intIndex) Then

          ' Adiciono uma coluna onde ficará apenas o traço no meio com continuação de linha
          Show.Image strPath & conLineMiddle

       ' Se este for o último item da arvore
       Else

          ' Adiciono uma coluna onde ficará apenas o traço no meio sem continuação de linha
          Show.Image strPath & conLineEnd

       End If

    End If

  ' Finaliza esta coluna
  Table.EndColumn

  ' Devolvendo a cor padrão
  Table.ColumnColor = ""

End Sub

Rem =========================================================================
Rem Desenha o ícone do item do "TreeView".
Rem -------------------------------------------------------------------------
Private Sub ShowIcon(ByVal intIndex As Integer)

  ' Variável que armazena o nome da imagem que deverá ser aberta
  Dim strImageName

  ' Se este item estiver aberto
  If sdicItems(conItemOpened & intIndex) Then

    ' Configura o nome da imagem que deverá ser mostrada para a padrão aberta
    strImageName = DefaultImage & conOpened

  ' Se este item estiver fechado
  Else

    ' Configura o nome da imagem que deverá ser mostrada para a padrão fechada
    strImageName = DefaultImage & conClosed

  End If

  ' Se foi informado o nome da imagem para este item
  If sdicItems(conItemImage & intIndex) > "" Then

     ' Se este item estiver aberto
     If sdicItems(conItemOpened & intIndex) Then

        ' Configura o nome da imagem que deverá ser mostrada a imagem aberta
        strImageName = sdicItems(conItemImage & intIndex) & conOpened

     ' Se este item estiver fechado
     Else

        ' Configura o nome da imagem que deverá ser mostrada a imagem fechada
        strImageName = sdicItems(conItemImage & intIndex) & conClosed

     End If

  End If

  ' Se esta imagem existir na lista de imagens
  If sdicImages.Exists(strImageName) Then

     ' Utiliza a cor deste item
     Table.ColumnColor = sdicItems(conItemBGColor & intIndex)

     ' Adiciono uma coluna onde ficará a imagem de estado
     Table.Column "<img src=""" & sdicImages(strImageName) & """ border=0>"

      ' Devolvendo a cor padrão
      Table.ColumnColor = ""

  End If

End Sub

Rem =========================================================================
Rem Mostra o título do item do "TreeView".
Rem -------------------------------------------------------------------------
Private Sub ShowCaption(ByVal intIndex As Integer)

  Dim i As Integer

  ' Como vou começar a definição do texto do item, redefino a largura da coluna para a padrão
  Table.ColumnWidth = ""

  ' Espande a coluna para que se espanda até que ocupe a quantidade de colunas restantes da tabela
  Table.ColumnColSpan = ((sintMaxNodes * 2) - sdicItems(conItemIdent & intIndex) - sdicItems(conItemColumnCount & intIndex))
'  Table.ColumnColSpan = ((sintMaxNodes + 1) - sdicItems(conItemIdent & intIndex) - sdicItems(conItemColumnCount & intIndex))

  ' Configura a cor de fundo deste item
  Table.ColumnColor = sdicItems(conItemBGColor & intIndex)

  ' Informa que está coluna não pode quebrar
  Table.ColumnNoWrap = sdicItems(conItemNoWrap & intIndex)

  ' Informa qual será o alinhamento horizontal desta coluna
  Table.ColumnAlign = sdicItems(conItemHAlign & intIndex)

  ' Informa qual será o alinhamento vertical desta coluna
  Table.ColumnVAlign = sdicItems(conItemVAlign & intIndex)

  ' Informa qual será a altura desta coluna
  Table.ColumnHeight = sdicItems(conItemHeight & intIndex)

  ' Informa qual será a largura desta coluna
  Table.ColumnWidth = sdicItems(conItemWidth & intIndex)

  ' Inicializa a coluna onde fica o texto do item
  Table.BeginColumn

    ' Se este item tiver um tamanho máximo de apresentação
    If (sdicItems(conItemMaxSize & intIndex) > 0) And (Len(sdicItems(conItemText & intIndex)) > sdicItems(conItemMaxSize & intIndex)) Then

       ' Reconfigura o título
       sdicItems(conItemTitle & intIndex) = sdicItems(conItemText & intIndex)

       ' Alterando o texto para que fique do tamanho máximo definido pelo usuário e com três "pontinhos"
       sdicItems(conItemText & intIndex) = Left(sdicItems(conItemText & intIndex), sdicItems(conItemMaxSize & intIndex) - 3) & "..."

    End If

    ' Mostra o espaço separador antes do texto do item
    Show.HTML "&nbsp;"

    ' Inicializa as "tags" necessárias para mostrar o texto formatado do
    ' modo definido pelo usuário
    BeginShowCaption intIndex, ""

    ' Mostra o texto
    Show.HTML sdicItems(conItemText & intIndex)

    ' Finaliza as "tags" que foram abertas
    EndShowCaption intIndex, ""

    ' Percorre todas as colunas deste item
    For i = 1 To sdicItems(conItemSubCount & intIndex)

        ' Inicializa as "tags" necessárias para mostrar o texto formatado do
        ' modo definido pelo usuário
        BeginShowCaption intIndex, conItemSub & i

          ' Mostra o texto
          Show.HTML sdicItems(conItemText & intIndex & conItemSub & i)

        ' Finaliza as "tags" que foram abertas
        EndShowCaption intIndex, conItemSub & i
        
    Next

  ' Finaliza a coluna onde fica o texto do item
  Table.EndColumn

  ' Configura o valor padrão para a propriedade "ColumnHeight"
  Table.ColumnHeight = ""

  ' Configura o valor padrão para a propriedade "ColumnWidth"
  Table.ColumnWidth = ""

  ' Configura o valor padrão para a propriedade "ColumnColSpan"
  Table.ColumnColSpan = ""

  ' Percorre todas as colunas deste item
  For i = 1 To sdicItems(conItemColumnCount & intIndex)
  
      ' Inicializa esta coluna
      Table.BeginColumn

        ' Inicializa as "tags" necessárias para mostrar o texto formatado do
        ' modo definido pelo usuário
        BeginShowCaption intIndex, conItemColumn & CStr(i)
        
        ' Mostra o texto da coluna
        Show.HTML sdicItems(conItemText & intIndex & conItemColumn & CStr(i))

        ' Finaliza as "tags" que foram abertas
        EndShowCaption intIndex, conItemColumn & CStr(i)

      ' Finaliza esta coluna
      Table.EndColumn

  Next

End Sub

Rem =========================================================================
Rem Inicia a configuração do texto.
Rem -------------------------------------------------------------------------
Private Sub BeginShowCaption(ByVal intIndex As Integer, _
                             ByVal strSufixName As String)

  ' Se este item deve aparecer em negrito
  If sdicItems(conItemBold & intIndex & strSufixName) Then

     ' Inicializa o negrito
     Show.Bold

  End If

  ' Se este item deve aparecer em italico
  If sdicItems(conItemItalic & intIndex & strSufixName) Then

     ' Inicializa o italico
     Show.Italic

  End If

  ' Se este item deve aparecer em subscrito
  If sdicItems(conItemUnderline & intIndex & strSufixName) Then

     ' Inicializa o subscrito
     Show.Underline

  End If

  ' Se alguma das propriedades da fonte foi configurada
  If (sdicItems(conItemFontFace & intIndex & strSufixName) & sdicItems(conItemFontColor & intIndex & strSufixName) & sdicItems(conItemFontSize & intIndex & strSufixName) > "") Then

     ' Inicializa a "tag" de fonte
     Show.HTML "<font"

     ' Se o nome da fonte foi informado
     Show.TagParameter "face", sdicItems(conItemFontFace & intIndex & strSufixName)

     ' Se a cor da fonte foi informada
     Show.TagParameter "color", sdicItems(conItemFontColor & intIndex & strSufixName)

     ' Se o tamanho da fonte foi informado
     Show.TagParameter "size", sdicItems(conItemFontSize & intIndex & strSufixName)

     ' Adiciona ao HTML o sinal ">" da tag "<font>"
     Show.HTML ">"

  End If

  ' Se o "link" foi informado
  If sdicItems(conItemURL & intIndex & strSufixName) > "" Then

    ' Inicializa "link" do texto
    Show.HTML "<a href=""" & sdicItems(conItemURL & intIndex & strSufixName) & """"

    ' Se houver um alvo configurado
    Show.TagParameter "target", sdicItems(conItemTarget & intIndex & strSufixName)
    
    ' Se houver um alvo configurado
    Show.TagParameter "title", sdicItems(conItemTitle & intIndex & strSufixName)
    
    ' Se houver um alvo configurado
    Show.TagParameter "style", sdicItems(conItemStyle & intIndex & strSufixName)
    
    ' Adiciona ao HTML o sinal ">" da tag "<a>"
    Show.HTML ">"

  End If

End Sub

Rem =========================================================================
Rem Finaliza a configuração do texto.
Rem -------------------------------------------------------------------------
Private Sub EndShowCaption(ByVal intIndex As Integer, _
                           ByVal strSufixName As String)

  ' Se o "link" não foi informado
  If sdicItems(conItemURL & intIndex & strSufixName) > "" Then

    ' Inicializa "link" do texto
    Show.HTML "</a>"

  End If

  If (sdicItems(conItemFontFace & intIndex & strSufixName) > "") Or _
     (sdicItems(conItemFontColor & intIndex & strSufixName) > "") Or _
     (sdicItems(conItemFontSize & intIndex & strSufixName) > "") Then

     Show.HTML "</font>"

  End If


  ' Se este item deve aparecer em subscrito
  If sdicItems(conItemUnderline & intIndex & strSufixName) Then

     ' Inicializa o subscrito
     Show.UnderlineEnd

  End If

  ' Se este item deve aparecer em italico
  If sdicItems(conItemItalic & intIndex & strSufixName) Then

     ' Inicializa o italico
     Show.ItalicEnd

  End If

  ' Se este item deve aparecer em negrito
  If sdicItems(conItemBold & intIndex & strSufixName) Then

     ' Inicializa o negrito
     Show.BoldEnd

  End If

End Sub

Rem =========================================================================
Rem Inicializa o valor padrão das propriedades.
Rem -------------------------------------------------------------------------
Private Sub InitializeProperties()

  ' Inicializa o valor padrão das propriedades
  ItemTarget = ""
  ItemNoWrap = False
  ItemBold = False
  ItemItalic = False
  ItemUnderline = False
  ItemBackColor = ""
  ItemHAlign = ""
  ItemVAlign = ""
  ItemWidth = ""
  ItemHeight = ""
  ItemFontColor = ""
  ItemFontFace = ""
  ItemFontSize = ""
  TreeBGColor = ""
  ItemMaxSize = 0
  ItemStyle = ""
  TreeStyleLink = ""
  TreeStyleActive = ""
  TreeStyleVisited = ""
  TreeStyleHover = ""
  sSetLevelID = ""
  
  ' Limpa as listas de itens nós e imagens
  sdicNodes.RemoveAll
  sdicItems.RemoveAll
  sdicImages.RemoveAll

End Sub

' TreeView.BeginTree
'
'    TreeView.AddImage( strInitial, strAddress, [strOpen], [strClose], [strExtension] )
'
'    TreeView.DefaultImage     = strImageInitial
'    TreeView.ItemTarget       = strTarget
'    TreeView.ItemNoWrap       = bnlNoWrap
'    TreeView.ItemBGColor      = strColor
'    TreeView.ItemBold         = bnlBold
'    TreeView.ItemItalic       = bnlItalic
'    TreeView.ItemUnderline    = bnlUnderline
'    TreeView.ItemHAlign       = strHAlign
'    TreeView.ItemVAlign       = strVAlign
'    TreeView.ItemWidth        = strWidth
'    TreeView.ItemHeight       = strHeight
'    TreeView.ItemFontColor    = strColor
'    TreeView.ItemFontFace     = strFace
'    TreeView.ItemFontSize     = strSize
'    TreeView.TreeBGColor      = strColor
'    TreeView.ItemMaxSize      = intMaxSize
'    TreeView.ItemStyle        = strStyle
'    TreeView.TreeStyleLink    = strStyleLink
'    TreeView.TreeStyleActive  = strStyleActive
'    TreeView.TreeStyleVisited = strStyleVisited
'    TreeView.TreeStyleHover   = strStyleHover
'    TreeView.CurrentChild     = intChild
'    TreeView.SetLevelID       = strSetLevelID
'    TreeView.GetLevelID       = strGetLevelID
'    TreeView.IsLevelID        = strIsLevelID
'
'    Tree.BeginNode( strText, strURL, [strItemTitle], [strImageInitial], [blnOpened] )
'
'      TreeView.AddItem( strText, strURL, [strItemTitle], [strImageInitial] )
'
'      strEqual = TreeView.Equal( strOption, strValue)
'
'      > TreeView.ColumnStyle = intAlignStyle
'
'      > TreeView.AddColumn ( strText )
'
'      ? TreeView.AddLookup ( strTable, strKeyField, strDisplayField, intMaxSize, intMaxLines, bnlWordWrap, strOptionURL )
'      ? TreeView.LookupFilterAdd
'
'    TreeView.EndNode
'
' TreeView.EndTree
'<a href=#new onclick="window.open('URL','Caption','status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no,width=600,height=450,top=0,left=0');">
