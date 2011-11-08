<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE VIRTUAL="/HiperTools/HiperTools30.inc"-->
<!--#INCLUDE VIRTUAL="/HiperTools/Objects.inc"-->
<!--#INCLUDE FILE="Constants.inc"-->
<%
REM =========================================================================
REM /ZEUSIIIMenu.asp
REM -------------------------------------------------------------------------
REM Nome     : Menu do Entregas
REM Descricao: Menu para o Sistema Entregas
REM Home     : http://www.hipertools.com.br/
REM Criacao  : 4/21/0 9:31PM
REM Autor    : Zevallos Tecnologia em Informacao
REM Versao   : 2.0.0.0
REM Local    : Brasília - DF
REM Companhia: 2000 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

  Private sobjRS

  Private sobjConn

  Private sstrSN

  sstrSN = "Entregas.Asp"

  Private sstrIC
  Private sstrNS

  sstrIC = Request.QueryString("IC")
  sstrNS = Request.QueryString("NS")

  Main

REM =========================================================================
REM Rotina Principal do Sistema
REM -------------------------------------------------------------------------
Private Sub Main

  Server.ScriptTimeOut = conScriptTimeout
  Session.TimeOut      = conSessionTimeout

  Set sobjRS = Server.CreateObject("ADODB.RecordSet")

  Set sobjConn  = Server.CreateObject("ADODB.Connection")

  sobjConn.ConnectionTimeout = 300
  sobjConn.CommandTimeout = 300
  sobjConn.Open Session("ConnectionString")

  MainBody

  Server.ScriptTimeOut = Session("ScriptTimeOut")

  Set sobjRS = nothing

  Set sobjConn  = nothing

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Main
REM =========================================================================

REM =========================================================================
REM Monta o Menu em Tree-View
REM -------------------------------------------------------------------------
Sub ShowTreeViewMenu

  Session("htDefaultHiperToolsWorkPath") = "/HiperTools"

  Show.HTML "<IMG SRC=""/HiperTools/img/Folder/home.gif"" ALT=""Página Inicial"" BORDER=""0"">"

  TreeView.ItemFontSize     = 1
  TreeView.ItemTarget       = "Body"
  TreeView.ItemMaxSize      = 15
  TreeView.TreeStyleLink    = "font-style: normal; text-decoration: none; color: blue"
  TreeView.TreeStyleVisited = "font-style: normal; text-decoration: none; color: blue"
  TreeView.TreeStyleActive  = "font-style: normal; text-decoration: underline; color: red"
  TreeView.TreeStyleHover   = "font-style: normal; text-decoration: underline; color: red"
  TreeView.ItemNoWrap       = True
  TreeView.ItemHaveMore     = True

  TreeView.BeginTree

    TreeView.BeginNode "Entregas", , , True
   
      TreeView.AddItem "Produtos", ConfigItem(sstrSN, efQueryStrAction, efQSActionList, TreeView.Equal(efQueryStrWhat,conOptionProdutos)) & efQueryStrEditableStr
         TreeView.AddSubCaption " ("
         TreeView.AddSubCaption "Incluir", ConfigItem(sstrSN, efQueryStrAction, efQSActionAdd, TreeView.Equal(efQueryStrWhat,conOptionProdutos)) & efQueryStrEditableStr, "Incluir um Produto", TreeView.ItemTarget
         TreeView.AddSubCaption ","
         TreeView.AddSubCaption "1", ConfigItem(sstrSN, efQueryStrAction, efQSActionSummary, TreeView.Equal(efQueryStrWhat,conOptionProdutos) & TreeView.Equal(efQueryStrList, 1) ), "Relatórios por Índice", TreeView.ItemTarget
         TreeView.AddSubCaption ","
         TreeView.AddSubCaption "2", ConfigItem(sstrSN, efQueryStrAction, efQSActionSummary, TreeView.Equal(efQueryStrWhat,conOptionProdutos) & TreeView.Equal(efQueryStrList, 2) ), "Relatório por Fabricante", TreeView.ItemTarget
         TreeView.AddSubCaption ","
         TreeView.AddSubCaption "3", ConfigItem(sstrSN, efQueryStrAction, efQSActionSummary, TreeView.Equal(efQueryStrWhat,conOptionProdutos) & TreeView.Equal(efQueryStrList, 2) ), "Relatório por Disponibilidade", TreeView.ItemTarget
         TreeView.AddSubCaption ")"

      TreeView.AddItem "Fornecedores", ConfigItem(sstrSN, efQueryStrAction, efQSActionList, TreeView.Equal(efQueryStrWhat,conOptionFornecedores)) & efQueryStrEditableStr
         TreeView.AddSubCaption " ("
         TreeView.AddSubCaption "Incluir", ConfigItem(sstrSN, efQueryStrAction, efQSActionAdd, TreeView.Equal(efQueryStrWhat,conOptionFornecedores) ) & efQueryStrEditableStr, "Incluir um Fornecedor", TreeView.ItemTarget
         TreeView.AddSubCaption ","
         TreeView.AddSubCaption "1", ConfigItem(sstrSN, efQueryStrAction, efQSActionSummary, TreeView.Equal(efQueryStrWhat,conOptionFornecedores) & TreeView.Equal(efQueryStrList, 1) ), "Relatórios por Cidade", TreeView.ItemTarget
         TreeView.AddSubCaption ","
         TreeView.AddSubCaption "2", ConfigItem(sstrSN, efQueryStrAction, efQSActionSummary, TreeView.Equal(efQueryStrWhat,conOptionFornecedores) & TreeView.Equal(efQueryStrList, 2) ), "Relatório por UF", TreeView.ItemTarget
         TreeView.AddSubCaption ")"

      TreeView.AddItem "Funcionários", ConfigItem(sstrSN, efQueryStrAction, efQSActionList, TreeView.Equal(efQueryStrWhat,conOptionFuncionarios)) & efQueryStrEditableStr
         TreeView.AddSubCaption " ("
         TreeView.AddSubCaption "Incluir", ConfigItem(sstrSN, efQueryStrAction, efQSActionAdd, TreeView.Equal(efQueryStrWhat,conOptionFuncionarios) ) & efQueryStrEditableStr, "Incluir um Funcionário", TreeView.ItemTarget
         TreeView.AddSubCaption ","
         TreeView.AddSubCaption "1", ConfigItem(sstrSN, efQueryStrAction, efQSActionSummary, TreeView.Equal(efQueryStrWhat,conOptionFuncionarios) & TreeView.Equal(efQueryStrList, 1) ), "Relatórios por Cidade", TreeView.ItemTarget
         TreeView.AddSubCaption ","
         TreeView.AddSubCaption "2", ConfigItem(sstrSN, efQueryStrAction, efQSActionSummary, TreeView.Equal(efQueryStrWhat,conOptionFuncionarios) & TreeView.Equal(efQueryStrList, 2) ), "Relatório por UF", TreeView.ItemTarget
         TreeView.AddSubCaption ","
         TreeView.AddSubCaption "3", ConfigItem(sstrSN, efQueryStrAction, efQSActionSummary, TreeView.Equal(efQueryStrWhat,conOptionFuncionarios) & TreeView.Equal(efQueryStrList, 3) ), "Relatório por Orgão Expedidor da RG", TreeView.ItemTarget
         TreeView.AddSubCaption ","
         TreeView.AddSubCaption "4", ConfigItem(sstrSN, efQueryStrAction, efQSActionSummary, TreeView.Equal(efQueryStrWhat,conOptionFuncionarios) & TreeView.Equal(efQueryStrList, 4) ), "Relatório por Estado Civil", TreeView.ItemTarget
         TreeView.AddSubCaption ","
         TreeView.AddSubCaption "5", ConfigItem(sstrSN, efQueryStrAction, efQSActionSummary, TreeView.Equal(efQueryStrWhat,conOptionFuncionarios) & TreeView.Equal(efQueryStrList, 5) ), "Relatório por Sexo", TreeView.ItemTarget
         TreeView.AddSubCaption ")"

      TreeView.AddItem "Clientes", ConfigItem(sstrSN, efQueryStrAction, efQSActionList, TreeView.Equal(efQueryStrWhat,conOptionClientes)) & efQueryStrEditableStr
         TreeView.AddSubCaption " ("
         TreeView.AddSubCaption "Incluir", ConfigItem(sstrSN, efQueryStrAction, efQSActionAdd, TreeView.Equal(efQueryStrWhat,conOptionClientes) ) & efQueryStrEditableStr, "Incluir um Cliente", TreeView.ItemTarget
         TreeView.AddSubCaption ","
         TreeView.AddSubCaption "1", ConfigItem(sstrSN, efQueryStrAction, efQSActionSummary, TreeView.Equal(efQueryStrWhat,conOptionClientes) & TreeView.Equal(efQueryStrList, 1) ), "Relatórios por Cidade", TreeView.ItemTarget
         TreeView.AddSubCaption ","
         TreeView.AddSubCaption "2", ConfigItem(sstrSN, efQueryStrAction, efQSActionSummary, TreeView.Equal(efQueryStrWhat,conOptionClientes) & TreeView.Equal(efQueryStrList, 2) ), "Relatório por UF", TreeView.ItemTarget
         TreeView.AddSubCaption ","
         TreeView.AddSubCaption "3", ConfigItem(sstrSN, efQueryStrAction, efQSActionSummary, TreeView.Equal(efQueryStrWhat,conOptionClientes) & TreeView.Equal(efQueryStrList, 2) ), "Relatório por Função Social", TreeView.ItemTarget
         TreeView.AddSubCaption ")"

      TreeView.BeginNode "Tabelas", , , True
        TreeView.AddItem "Funções", ConfigItem(sstrSN, efQueryStrAction, efQSActionList, TreeView.Equal(efQueryStrWhat,conOptionFuncoes)) & efQueryStrEditableStr
         TreeView.AddSubCaption " ("
         TreeView.AddSubCaption "Incluir", ConfigItem(sstrSN, efQueryStrAction, efQSActionAdd, TreeView.Equal(efQueryStrWhat,conOptionFuncoes) ) & efQueryStrEditableStr, "Incluir uma Função", TreeView.ItemTarget
         TreeView.AddSubCaption ")"

        TreeView.AddItem "Índices", ConfigItem(sstrSN, efQueryStrAction, efQSActionList, TreeView.Equal(efQueryStrWhat,conOptionIndices)) & efQueryStrEditableStr
         TreeView.AddSubCaption " ("
         TreeView.AddSubCaption "Incluir", ConfigItem(sstrSN, efQueryStrAction, efQSActionAdd, TreeView.Equal(efQueryStrWhat,conOptionIndices) ) & efQueryStrEditableStr, "Incluir um Índice", TreeView.ItemTarget
         TreeView.AddSubCaption ")"

        TreeView.AddItem "Valores", ConfigItem(sstrSN, efQueryStrAction, efQSActionList, TreeView.Equal(efQueryStrWhat,conOptionValores)) & efQueryStrEditableStr
         TreeView.AddSubCaption " ("
         TreeView.AddSubCaption "Incluir", ConfigItem(sstrSN, efQueryStrAction, efQSActionAdd, TreeView.Equal(efQueryStrWhat,conOptionValores) ) & efQueryStrEditableStr, "Incluir um Valor", TreeView.ItemTarget
         TreeView.AddSubCaption ")"
  
      TreeView.EndNode

    TreeView.EndNode

  TreeView.EndTree

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowTreeViewMenu
REM =========================================================================

REM =========================================================================
REM Monta os Ítens do Menu
REM -------------------------------------------------------------------------

Private Function ConfigItem(ByVal strLocation, ByVal strConstant, ByVal strValue, ByVal strEqual )

  If strLocation > "" Then

     If strConstant > "" Then
        ConfigItem = strLocation & "?" & strConstant & "=" & strValue
     End If

     If strEqual > "" Then
        If ConfigItem > "" Then
           ConfigItem = ConfigItem & strEqual
        Else
           ConfigItem = strLocation & "?" & strEqual
        End If
     End If

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Sub ConfigMenuItem
REM =========================================================================

REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub MainBody

  Default.HTMLBegin
  Response.Write "<TITLE>Menu do Entregas v1.0</TITLE>"

  Response.Write "<BODY>"
  Response.Write "<BASEFONT FACE=""Arial, Helvetica, Sans-Serif"">"

  Table.Style.Color2 = ""
  Table.BorderColor = "Black"

  Table.Style.Color1 = "#FFFFFF"
  Table.Style.Color2 = "#FFFFFF"

  Table.BeginTable "100%"                         'Início Tabela Principal

  Table.BeginRow 4
  Table.Column "<B>Entregas v1.0</B>"
  Table.EndRow

  Table.BeginRow 2                                'Inicio Segunda linha
  Table.BeginColumn                               'Primeira Coluna (Menu)


  'Chama o Menu no Tree View

  ShowTreeViewMenu


  Table.EndColumn                                 'Fim Primeira Coluna (Menu)
  Table.EndRow                                    'Fim Segunda linha
  Table.EndTable                                  'Fim Tabela Principal

  Response.Write "</BODY>"

  Default.HTMLEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================

REM -------------------------------------------------------------------------
REM Fim do ZEUSIIIMenu.asp
REM =========================================================================
%>