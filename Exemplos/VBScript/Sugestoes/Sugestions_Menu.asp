<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE VIRTUAL="/HiperTools/HiperTools30.inc"-->
<!--#INCLUDE VIRTUAL="/HiperTools/Objects.inc"-->
<!--#INCLUDE FILE="Sugestions.inc"-->
<!--#INCLUDE FILE="Constants.inc"-->
<%
REM =========================================================================
REM /SugestMenu.asp
REM -------------------------------------------------------------------------
REM Nome     : Menu do Sugestões para o HiperTools 30
REM Descricao: Menu para o Sistema Sugestões para o HiperTools 30
REM Home     : http://http://sugestions.zevallos2.com.br/
REM Criacao  : 2000/04/27 12:18AM
REM Autor    : Ruben Zevallos
REM          : Eduardo Gonçalves
REM Versao   : 1.1.0.0
REM Local    : Brasília - DF
REM Companhia: Zevallos
REM -------------------------------------------------------------------------

  Main

REM =========================================================================
REM Rotina Principal do Sistema
REM -------------------------------------------------------------------------
Private Sub Main

  Server.ScriptTimeOut = conScriptTimeout

  MainBody

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Main
REM =========================================================================

REM =========================================================================
REM Monta o Menu em Tree-View
REM -------------------------------------------------------------------------
Sub ShowTreeViewMenu

  Session("htDefaultHiperToolsWorkPath") = "/HiperTools"

  TreeView.DefaultImage     = "Folder"
  TreeView.TreeStyleLink    = "font-style: normal; text-decoration: none; color:orange;font-size: 11"
  TreeView.TreeStyleActive  = "font-style: normal; text-decoration: none; color:orange;font-size: 11"
  TreeView.TreeStyleVisited = "font-style: normal; text-decoration: none; color:orange;font-size: 11"
  TreeView.TreeStyleHover   = "font-style: normal; text-decoration: underline; color:orange;font-size: 11"
  TreeView.ItemTarget       = "Body"
  TreeView.ItemMaxSize      = 42
  TreeView.ItemNoWrap       = True

  TreeView.AddImage "Folder", "/HiperTools/Img/Folder/"


  TreeView.BeginTree
    TreeView.ItemBGColor = "White"
    TreeView.BeginNode "Sugestões", sstrSN, , , True
      TreeView.AddItem "Bugs",          ConfigEditForm( sstrSN, conWhatBugs, efQSActionList, "", "", "", "", "", "" )
         TreeView.AddColumn "Incluir",    ConfigEditForm( sstrSN, conWhatBugs, efQSActionAdd, "", "", "", "", "", "" ), "Incluir um Bug", TreeView.ItemTarget

      TreeView.AddItem "Idéias",        ConfigEditForm( sstrSN, conWhatIdeias, efQSActionList, "", "", "", "", "", "" )
         TreeView.AddColumn "Incluir",    ConfigEditForm( sstrSN, conWhatIdeias, efQSActionAdd, "", "", "", "", "", "" ), "Incluir uma Ideia", TreeView.ItemTarget

      TreeView.AddItem "Melhorias",     ConfigEditForm( sstrSN, conWhatMelhorias, efQSActionList, "", "", "", "", "", "" )
         TreeView.AddColumn "Incluir",    ConfigEditForm( sstrSN, conWhatMelhorias, efQSActionAdd, "", "", "", "", "", "" ), "Incluir uma Melhoria", TreeView.ItemTarget

      TreeView.AddItem "Classes",       ConfigEditForm( sstrSN, conWhatClasses, efQSActionList, "", "", "", "", "", "" )
         TreeView.AddColumn "Incluir",    ConfigEditForm( sstrSN, conWhatClasses, efQSActionAdd, "", "", "", "", "", "" ), "Incluir uma Classe", TreeView.ItemTarget

      TreeView.AddItem "Posição",       ConfigEditForm( sstrSN, conWhatPosition, efQSActionList, "", "", "", "", "", "" )
         TreeView.AddColumn "Incluir",    ConfigEditForm( sstrSN, conWhatPosition, efQSActionAdd, "", "", "", "", "", "" ), "Incluir Posição de uma Classe", TreeView.ItemTarget

      TreeView.AddItem "Pessoas",       ConfigEditForm( sstrSN, conWhatPessoas, efQSActionList, "", "", "", "", "", "" )
         TreeView.AddColumn "Incluir",    ConfigEditForm( sstrSN, conWhatPessoas, efQSActionAdd, "", "", "", "", "", "" ),"Incluir uma Pessoa", TreeView.ItemTarget
    TreeView.EndNode
  TreeView.EndTree

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowTreeViewMenu
REM =========================================================================

REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub MainBody

  Default.HTMLBegin
  Response.Write "<TITLE>Sugestões para o HiperTools 30</TITLE>"

  Response.Write "<BODY>"
  Response.Write "<BASEFONT FACE=""Arial, Helvetica, Sans-Serif"">"

  Table.Border = 2
  Table.Style.Color2 = ""
  Table.BorderColor = "Black"

  Table.Style.Color1 = "#FFFFFF"
  Table.Style.Color2 = "#FFFFFF"

  Table.BeginTable "100%"                         'Início Tabela Principal

  Table.ColumnAlign = "LEFT"
  Table.BeginRow 4
  Table.BeginColumn

  Show.Image "Img/HTLogo.gif"

  Table.EndColumn
  Table.EndRow

  Table.BeginRow 4
  Table.EndRow
  Table.ColumnAlign = "LEFT"

    Table.BeginRow 2                                'Inicio Segunda linha
  Table.BeginColumn                               'Primeira Coluna (Menu)


REM  Chama o Menu no Tree View
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
REM Fim do Home.asp
REM =========================================================================
%>