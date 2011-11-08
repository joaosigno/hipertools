<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE VIRTUAL="/HiperTools/HiperTools30.inc"-->
<!--#INCLUDE VIRTUAL="/HiperTools/Objects.inc"-->
<!--#INCLUDE FILE="Constants.inc"-->
<%
REM =========================================================================
REM /ZEUSIIIMenu.asp
REM -------------------------------------------------------------------------
REM Nome     : Menu do ZEUS III
REM Descricao: Menu para o Sistema ZEUS III
REM Home     : http://www.hipertools.com.br/
REM Criacao  : 2/12/0 5:14PM
REM Autor    : Zevallos Tecnologia em Informacao
REM Versao   : 1.1.0.0
REM Local    : Brasília - DF
REM Copyright: 97-2000 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

  Private sobjRS
  Private sobjRSEmp
  Private sobjRSSis
  Private sobjRSTab
  Private sobjRSCam

  Private sobjConn

  Private sstrSN

  sstrSN = "ZEUSIII.Asp"

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

  Set sobjRSEmp = Server.CreateObject("ADODB.RecordSet")
  Set sobjRSSis = Server.CreateObject("ADODB.RecordSet")
  Set sobjRSTab = Server.CreateObject("ADODB.RecordSet")
  Set sobjRSCam = Server.CreateObject("ADODB.RecordSet")

  Set sobjConn  = Server.CreateObject("ADODB.Connection")

  sobjConn.ConnectionTimeout = 300
  sobjConn.CommandTimeout = 300
  sobjConn.Open Session("ConnectionString")

  MainBody

  Server.ScriptTimeOut = Session("ScriptTimeOut")

  Set sobjRS = nothing

  Set sobjRSEmp = nothing
  Set sobjRSSis = nothing
  Set sobjRSTab = nothing
  Set sobjRSCam = nothing
  Set sobjConn  = nothing

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Main
REM =========================================================================

REM =========================================================================
REM Monta o Menu em Tree-View
REM -------------------------------------------------------------------------
Sub ShowTreeViewMenu

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

    TreeView.BeginNode "Empresas", ConfigItem(sstrSN, efQueryStrAction, efQSActionList, TreeView.Equal(efQueryStrWhat,conWhatEmpresa)), , True
      TreeView.AddSubCaption " ("
        TreeView.AddSubCaption "Incluir", ConfigItem(sstrSN, efQueryStrAction, efQSActionAdd, TreeView.Equal(efQueryStrWhat,conWhatEmpresa) ) & efQueryStrEditableStr, "Incluir uma empresa", TreeView.ItemTarget
        TreeView.AddSubCaption ", "
        TreeView.ItemStyle = "font: bold 8pt"
        TreeView.AddSubCaption "Relatório", ConfigItem(sstrSN, efQueryStrAction, efQSActionSummary, TreeView.Equal(efQueryStrWhat,conWhatEmpresa) & TreeView.Equal(efQueryStrList, conListEmpresaUF ) ), "Relatório por unidade federativa", TreeView.ItemTarget
        TreeView.ItemStyle = ""
      TreeView.AddSubCaption ")"

      ShowEmpresas

    TreeView.EndNode

  TreeView.EndTree

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowTreeViewMenu
REM =========================================================================

REM =========================================================================
REM Monta os Ítens do Menu "Empresas"
REM -------------------------------------------------------------------------
Sub ShowEmpresas

  Dim sql, strNome

  sql = "SELECT empCodigo,empNome, empRazao_Social FROM zeuEmpresas"

  sobjRSEmp.Open sql, sobjConn, adOpenForwardOnly, adLockReadOnly

  If Not sobjRSEmp.EOF Then
    Do While Not sobjRSEmp.EOF
      If sobjRSEmp("empNome") > "" Then
        strNome = sobjRSEmp("empNome")

        TreeView.SetLevelID = sobjRSEmp("empCodigo")
        TreeView.BeginNode strNome, ConfigItem(sstrSN, efQueryStrAction, efQSActionEditor, TreeView.Equal(efQueryStrWhat,conWhatEmpresa) & TreeView.Equal(efQueryStrFilter, Server.UrlEncode("empCodigo=" & sobjRSEmp("empCodigo"))) & efQueryStrEditableStr ), sobjRSEmp("empRazao_Social")

          TreeView.SetLevelID = sobjRSEmp("empCodigo")
          TreeView.BeginNode "Sistemas", ConfigItem(sstrSN, efQueryStrAction, efQSActionList, TreeView.Equal(efQueryStrWhat,conWhatSistema) & TreeView.Equal(efQueryStrFilter, Server.UrlEncode("sisEmpresa=" & sobjRSEmp("empCodigo"))) & efQueryStrEditableStr ), "Lista dos sistemas para empresa (" & sobjRSEmp("empNome") & ")"
            TreeView.AddSubCaption " ("
              TreeView.AddSubCaption "Incluir", ConfigItem( sstrSN, efQueryStrAction, efQSActionAdd, TreeView.Equal(efQueryStrWhat,conWhatSistema) ) & TreeView.Equal(efQueryStrFilter, Server.UrlEncode("sisEmpresa=" & sobjRSEmp("empCodigo"))) & TreeView.Equal(efQueryStrDefaults, Server.UrlEncode("sisEmpresa=" & sobjRSEmp("empCodigo"))) & efQueryStrEditableStr, "Incluir um sistema", TreeView.ItemTarget
            TreeView.AddSubCaption ")"

            ShowSistemas sobjRSEmp("empCodigo")

          TreeView.EndNode
        TreeView.EndNode

      End If

      sobjRSEmp.MoveNext

    Loop

  End If

  sobjRSEmp.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowEmpresas
REM =========================================================================

REM =========================================================================
REM Monta os Ítens do Menu "Sistemas"
REM -------------------------------------------------------------------------
Sub ShowSistemas(ByVal strEmpresa)
Dim sql
Dim strSigla

  sql = "SELECT sisCodigo,sisNome, sisSigla, sisVersao FROM zeuSistemas" & _
        " WHERE sisEmpresa=" & strEmpresa

  sobjRSSis.Open sql, sobjConn, adOpenForwardOnly, adLockReadOnly

  If Not sobjRSSis.EOF Then
    If TreeView.IsLevelID Then
      Do While Not sobjRSSis.EOF
        strSigla = sobjRSSis("sisSigla")

        If strSigla > "" Then


          TreeView.SetLevelID = sobjRSSis("sisCodigo")

          TreeView.ItemStyle = "font: bold 8pt"

          TreeView.BeginNode strSigla, ConfigItem(sstrSN, efQueryStrAction, efQSActionEditor, TreeView.Equal(efQueryStrWhat,conWhatSistema) & TreeView.Equal(efQueryStrFilter, Server.UrlEncode("sisCodigo=" & sobjRSSis("sisCodigo"))) & efQueryStrEditableStr ), sobjRSSis("sisNome")
            TreeView.ItemStyle = ""
            TreeView.AddSubCaption " ("
              TreeView.AddSubCaption "Importar", ConfigItem(sstrSN, conOptions, conOptionInitImport, TreeView.Equal(conImportCompany, strEmpresa ) & TreeView.Equal(conImportSystem, sobjRSSis("sisCodigo") ) ), "Importar tabelas para o sistema (" & sobjRSSis("sisNome") & ")", TreeView.ItemTarget
            TreeView.AddSubCaption ", "
              TreeView.AddSubCaption "Construir", ConfigItem(sstrSN, conToMakeSystem, sobjRSSis("sisCodigo"), ""), "Construir o sistema (" & sobjRSSis("sisNome") & ")", TreeView.ItemTarget
            TreeView.AddSubCaption ")"

            TreeView.SetLevelID = sobjRSSis("sisCodigo")
            TreeView.BeginNode "Tabelas", ConfigItem(sstrSN, efQueryStrAction, efQSActionList, TreeView.Equal(efQueryStrWhat,conWhatTabela) & TreeView.Equal(efQueryStrFilter, Server.UrlEncode("tabSistema=" & sobjRSSis("sisCodigo"))) & efQueryStrEditableStr )
              TreeView.AddSubCaption " ("
                TreeView.AddSubCaption "Incluir", ConfigItem( sstrSN, efQueryStrAction, efQSActionAdd, TreeView.Equal(efQueryStrWhat,conWhatTabela) ) & TreeView.Equal(efQueryStrFilter, Server.UrlEncode("tabSistema=" & sobjRSSis("sisCodigo"))) & TreeView.Equal(efQueryStrDefaults, Server.UrlEncode("tabSistema=" & sobjRSSis("sisCodigo"))) & efQueryStrEditableStr, "Incluir uma tabela", TreeView.ItemTarget
              TreeView.AddSubCaption ")"

              ShowTabelas sobjRSSis("sisCodigo")

            TreeView.EndNode
          TreeView.EndNode

        End If

        sobjRSSis.MoveNext

      Loop

    End If

  End If

  sobjRSSis.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowSistemas
REM =========================================================================

REM =========================================================================
REM Monta os Ítens do Menu "Tabelas"
REM -------------------------------------------------------------------------
Sub ShowTabelas(ByVal strSistema)
Dim sql
Dim strSigla

  sql = "SELECT tabCodigo, tabNome, tabSigla FROM zeuTabelas" & _
        " WHERE tabSistema=" & strSistema

  sobjRSTab.Open sql, sobjConn, adOpenForwardOnly, adLockReadOnly

  If Not sobjRSTab.EOF Then
    If TreeView.IsLevelID Then
      Do While Not sobjRSTab.EOF
        strSigla = sobjRSTab("tabSigla")

        If strSigla > "" Then

          TreeView.SetLevelID = sobjRSTab("tabCodigo")

          TreeView.BeginNode strSigla, ConfigItem(sstrSN, efQueryStrAction, efQSActionEditor, TreeView.Equal(efQueryStrWhat,conWhatTabela) & TreeView.Equal(efQueryStrFilter, Server.UrlEncode("tabCodigo=" & sobjRSTab("tabCodigo"))) & efQueryStrEditableStr ), sobjRSTab("tabNome")

            TreeView.SetLevelID = sobjRSTab("tabCodigo")
            TreeView.BeginNode "Campos", ConfigItem(sstrSN, efQueryStrAction, efQSActionList, TreeView.Equal(efQueryStrWhat,conWhatCampo) & TreeView.Equal(efQueryStrFilter, Server.UrlEncode("camTabela=" & sobjRSTab("tabCodigo"))) & efQueryStrEditableStr )
              TreeView.AddSubCaption " ("
              TreeView.AddSubCaption "Incluir", ConfigItem( sstrSN, efQueryStrAction, efQSActionAdd, TreeView.Equal(efQueryStrWhat,conWhatCampo) ) & TreeView.Equal(efQueryStrFilter, Server.UrlEncode("camTabela=" & sobjRSTab("tabCodigo"))) & TreeView.Equal(efQueryStrDefaults, Server.UrlEncode("camTabela=" & sobjRSTab("tabCodigo"))) & efQueryStrEditableStr, "Incluir um campo", TreeView.ItemTarget
              TreeView.AddSubCaption ","
              TreeView.AddSubCaption "1", ConfigItem(sstrSN, efQueryStrAction, efQSActionSummary, TreeView.Equal(efQueryStrWhat,conWhatCampo) & TreeView.Equal(efQueryStrList, conListCampoTipo ) ), "Relatórios por tipos de campos", TreeView.ItemTarget
              TreeView.AddSubCaption ","
              TreeView.AddSubCaption "2", ConfigItem(sstrSN, efQueryStrAction, efQSActionSummary, TreeView.Equal(efQueryStrWhat,conWhatCampo) & TreeView.Equal(efQueryStrList, conListCampoDelimitador ) ), "Relatório por tipos de delimitadores", TreeView.ItemTarget
              TreeView.AddSubCaption ")"

              ShowCampos sobjRSTab("tabCodigo")

            TreeView.EndNode
          TreeView.EndNode

        End If

        sobjRSTab.MoveNext

      Loop

    End If
  End If

  sobjRSTab.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowTabelas
REM =========================================================================

REM =========================================================================
REM Monta os Ítens do Menu "Tabelas"
REM -------------------------------------------------------------------------
Sub ShowCampos(ByVal strTabela)
Dim sql
Dim strNome
Dim strTipo

  sql = "SELECT camCodigo, camNome, camTipo, camTamanho FROM zeuCampos" & _
        " WHERE camTabela=" & strTabela

  sobjRSCam.Open sql, sobjConn, adOpenForwardOnly, adLockReadOnly

  If Not sobjRSCam.EOF Then
    If TreeView.IsLevelID Then
      Do While Not sobjRSCam.EOF
        strNome = sobjRSCam("camNome")

        If strNome > "" Then

          Select Case sobjRSCam("camTipo")
            Case 0
             strTipo = "Texto"
            Case 1
             strTipo = "Texto variável"
            Case 2
             strTipo = "Data e hora"
            Case 3
             strTipo = "Memorando"
            Case 4
             strTipo = "Inteiro longo"
            Case 5
             strTipo = "Inteiro"
            Case 6
             strTipo = "Inteiro curto"
            Case 7
             strTipo = "Ponto flutuante"
            Case 8
             strTipo = "Real"
            Case 9
             strTipo = "Monetário"
           Case 10
             strTipo = "Lógico"
           Case 11
             strTipo = "Incremental"

          Case Else
             strTipo = "Desconhecido"

        End Select

          TreeView.AddItem strNome, ConfigItem(sstrSN, efQueryStrAction, efQSActionEditor, TreeView.Equal(efQueryStrWhat,conWhatCampo) & TreeView.Equal(efQueryStrFilter, Server.UrlEncode("camCodigo=" & sobjRSCam("camCodigo"))) & efQueryStrEditableStr ), sobjRSCam("camNome") & ", " & strTipo & "(" & sobjRSCam("camTamanho") & ")"

        End If

        sobjRSCam.MoveNext

      Loop

    End If
  End If

  sobjRSCam.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowTabelas
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

  Default.BeginHTML
  Response.Write "<TITLE>Menu do ZEUS III v1.1</TITLE>"

  Response.Write "<BODY>"
  Response.Write "<BASEFONT FACE=""Arial, Helvetica, Sans-Serif"">"

  Table.Style.Color2 = ""
  Table.BorderColor = "Black"

  Table.Style.Color1 = "#FFFFFF"
  Table.Style.Color2 = "#FFFFFF"

  Table.BeginTable "100%"                         'Início Tabela Principal

  Table.BeginRow 4
  Table.Cell "<B>ZEUS III v1.1</B>"
  Table.EndRow

  Table.BeginRow 2                                'Inicio Segunda linha
  Table.BeginColumn                               'Primeira Coluna (Menu)


  'Chama o Menu no Tree View

  ShowTreeViewMenu


  Table.EndCell                                   'Fim Primeira Coluna (Menu)
  Table.EndRow                                    'Fim Segunda linha
  Table.EndTable                                  'Fim Tabela Principal

  Response.Write "</BODY>"

  Default.EndHTML

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================

REM -------------------------------------------------------------------------
REM Fim do ZEUSIIIMenu.asp
REM =========================================================================
%>