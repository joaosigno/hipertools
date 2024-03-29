<%
REM =========================================================================
REM  /MakeSystem.inc
REM -------------------------------------------------------------------------
REM Descricao: C�digo de gera o c�digo "HiperTools" do sistema informado
REM            pelo usu�rio
REM Cria��o  : 09:43h 01/12/99
REM Local    : Bras�lia/DF
REM Autor    : Zevallos Tecnologia em Informacao
REM Versao   : 1.0.0
REM Copyright: 97-2000 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

  ' Declarando as vari�veis globais
  Private sstrResultFile, sstrSigla, sstrSistema, sstrArquivo
  Private intTotal

  ' Declarando os objetos
  Private sobjFileSystem, sobjOutputFile
  Private sobjConnection, sobjRecordSet

REM =========================================================================
REM Procedimento que prepara o ambiente
REM -------------------------------------------------------------------------
Private Sub Prepare

  ' Cria os objetos
  Set sobjFileSystem = CreateObject( "Scripting.FileSystemObject" )
  Set sobjConnection = Server.CreateObject("ADODB.Connection")

  ' Configura e abre o objeto de conex�o
  sobjConnection.ConnectionTimeout = Session("ConnectionTimeout")
  sobjConnection.CommandTimeout = Session("CommandTimeout")
  sobjConnection.Open Session("ConnectionString"), _
                      Session("RuntimeUserName"), _
                      Session("RuntimePassword")

  ' Configura o objeto de acesso a dados
  Set sobjRecordSet = Server.CreateObject("ADODB.RecordSet")
  sobjRecordSet.CacheSize = 150
  sobjRecordSet.CursorType = adOpenDynamic
  sobjRecordSet.LockType = adLockPessimistic

  ' Zerando as vari�veis
  intTotal = 0

End Sub
REM -------------------------------------------------------------------------
REM Fim do Prepare
REM =========================================================================

REM =========================================================================
REM Procedimento que escreve o cabe�alho do arquivo
REM -------------------------------------------------------------------------
Private Sub MakeHead

  ' Recupera os dados do sistema e da empresa
  OpenRecordSet( conSelectAll & "zeuSistemas INNER JOIN zeuEmpresas ON ( zeuSistemas.sisEmpresa = zeuEmpresas.empCodigo ) WHERE sisCodigo = " & sstrSistema )

  ' Se n�o existir uma pasta com o nome da empresa. Crio a pasta.
  sstrResultFile = Initializer.SiteRootDir & sobjRecordSet( "empNome" )
  If Not sobjFileSystem.FolderExists( sstrResultFile ) Then
     sobjFileSystem.CreateFolder( sstrResultFile )
  End If

  ' Se n�o existir uma pasta com a sigla do sistema. Crio a pasta.
  sstrResultFile = sstrResultFile & "\" & sobjRecordSet( "sisSigla" )
  If Not sobjFileSystem.FolderExists( sstrResultFile ) Then
     sobjFileSystem.CreateFolder( sstrResultFile )
  End If

  ' Configurando o nome do arquivo
  sstrArquivo = sobjRecordSet( "sisSigla" ) & ".asp"

  ' Configurando o nome completo do arquivo
  sstrResultFile = sstrResultFile & "\" & sstrArquivo

  ' Adiciono o cabe�alho padr�o
  Default.PageHeaderDefault "<H1>Criando o sistema " & sobjRecordSet( "sisNome" ) & "</H1>"

  ' Mostra ao usu�rio o que est� sendo feito agora
  ShowStatus "Gerando o cabe�alho"

  ' Armazena a sigla do sistema
  sstrSigla = Trim( sobjRecordSet( "sisSigla" ) )

  ' Come�a escrevendp o cabe�alho com os dados da tabela de "sistemas"
  AddLine "<" & "%@ LANGUAGE=""VBSCRIPT"" %" & ">"
  AddLine "<!--#INCLUDE VIRTUAL=""/HiperTools/HiperTools30.inc""-->"
  AddLine "<!--#INCLUDE VIRTUAL=""/Objects.inc""-->"
  AddLine "<" & "%"
  AddLine "REM ========================================================================="
  AddLine "REM /" & sstrArquivo
  AddLine "REM -------------------------------------------------------------------------"
  AddLine "REM Nome     : " & sobjRecordSet( "sisNome" )
  AddLine "REM Descricao: " & sobjRecordSet( "sisDescricao" )
  AddLine "REM Home     : " & sobjRecordSet( "sisHome" )
  AddLine "REM Criacao  : " & Now
  AddLine "REM Autor    : " & sobjRecordSet( "sisResponsavel" ) & " - " & sobjRecordSet( "sisTelefone" )
  AddLine "REM Versao   : " & sobjRecordSet( "sisVersao" )
  AddLine "REM Local    : " & sobjRecordSet( "empCidade" ) & " - " & sobjRecordSet( "empUF" )
  AddLine "REM Companhia: " & sobjRecordSet( "empNome" )
  AddLine "REM -------------------------------------------------------------------------"
  AddLine ""

End Sub
REM -------------------------------------------------------------------------
REM Fim do MakeHead
REM =========================================================================

REM =========================================================================
REM Procedimento que escreve a �rea global do sistema
REM -------------------------------------------------------------------------
Private Sub MakeGlobalArea

  ' Mostra ao usu�rio o que est� sendo feito agora
  ShowStatus "Gerando a �rea global do sistema"

  ' Recupera os dados do sistema e da empresa
  OpenRecordSet( conSelectAll & "zeuTabelas WHERE tabSistema = " & sstrSistema )

  ' Come�a escrevendp o cabe�alho com os dados da tabela de "sistemas"
  AddLine "  Const conScriptTimeout  = 15"
  AddLine "  Const conSessionTimeout = 300"
  AddLine ""
  Do While Not sobjRecordSet.EOF
     AddLine "  Const " & LeftStr( "conWhat" & sobjRecordSet( "tabNome" ), 32 ) & " = " & """" & sobjRecordSet( "tabCodigo" ) & """"
     sobjRecordSet.MoveNext
  Loop
  AddLine ""
  AddLine "  Main"
  AddLine ""
  AddLine "REM ========================================================================="
  AddLine "REM Rotina Principal do Sistema                                              "
  AddLine "REM -------------------------------------------------------------------------"
  AddLine "Private Sub Main"
  AddLine ""
  AddLine "  Server.ScriptTimeOut = conScriptTimeout"
  AddLine "  Session.TimeOut      = conSessionTimeout"
  AddLine ""
  AddLine "  MainBody"
  AddLine ""
  AddLine "End Sub"
  AddLine "REM -------------------------------------------------------------------------"
  AddLine "REM Final da Sub Main                                                        "
  AddLine "REM ========================================================================="
  AddLine ""

End Sub
REM -------------------------------------------------------------------------
REM Fim do MakeGlobalArea
REM =========================================================================

REM =========================================================================
REM Procedimento que escreve o c�digo de constru��o do ambiente do sistema
REM -------------------------------------------------------------------------
Private Sub MakeEnvironment

  ' Mostra ao usu�rio o que est� sendo feito agora
  ShowStatus "Gerando o ambiente do sistema"

  ' Escreve o c�digo que formata as tabelas do sistema
  AddLine "REM ========================================================================="
  AddLine "REM Altera o estilo do Objeto Table                                          "
  AddLine "REM -------------------------------------------------------------------------"
  AddLine "Private Sub FormatTable"
  AddLine ""
  AddLine "  Table.Style.BaseColor            = """""
  AddLine "  Table.Style.HeaderColor          = ""Orange"""
  AddLine "  Table.Style.FirstAltColor        = ""#ECECEC"""
  AddLine "  Table.Style.AlternateColor       = ""#ECECEC"""
  AddLine "  Table.Style.LastColor            = ""#ECECEC"""
  AddLine "  Table.Style.BorderColor          = ""Orange"""
  AddLine "  Table.Style.BorderFormat         = tbBdFormatOnlyLines"
  AddLine "  Table.Style.ColorFormat          = tbStFormatTitle"
  AddLine "  Table.Style.ExternalBorder.Width = 3"
  AddLine "  Table.Style.HeaderBorder.Width   = 2"
  AddLine "  Table.Style.InternalBorder.Width = 1"
  AddLine ""
  AddLine "  Set Edit.Style = Table.Style"
  AddLine ""
  AddLine "End Sub"
  AddLine "REM -------------------------------------------------------------------------"
  AddLine "REM Fim do FormatTable                                                       "
  AddLine "REM ========================================================================="
  AddLine ""

End Sub
REM -------------------------------------------------------------------------
REM Fim do MakeEnvironment
REM =========================================================================

REM =========================================================================
REM Procedimento que escreve o c�digo de constru��o da tela principal
REM -------------------------------------------------------------------------
Private Sub MakeFirstPage

  ' Mostra ao usu�rio o que est� sendo feito agora
  ShowStatus "Gerando o ambiente do sistema"

  ' Posicionando no in�cio da tabela de "tabelas", levo em considera��o que
  ' a �ltima tabela a ser aberta foi na procedure e "MakeGlobalArea"
  sobjRecordSet.MoveFirst

  ' Escreve o c�digo que formata a primeira p�gina do sistema
  AddLine "REM ========================================================================="
  AddLine "REM Procedimento que mostra a primeira p�gina                                "
  AddLine "REM -------------------------------------------------------------------------"
  AddLine "Private Sub ShowFirstPage"
  AddLine ""
  AddLine "  Table.Spacing = 0"
  AddLine ""
  AddLine "  URL.BeginURL Initializer.ScriptName"
  AddLine "  URL.Add efQueryStrAction, efQSActionEditor"
  AddLine "  URL.Add efQueryStrEditable, ""1"""
  AddLine ""
  AddLine "    Table.BeginTable ""45%"", ""Painel de Controle"", 2, True"
  AddLine "      Table.CellAlign = ""center"""
  Do While Not sobjRecordSet.EOF
     AddLine "      Table.Row URL.GetURL( """ & LeftStr( sobjRecordSet( "tabNome" ) & """, ", 28 ) & _
                                "efQueryStrWhat & ""="" & conWhat" & sobjRecordSet( "tabNome" ) & " )"
     sobjRecordSet.MoveNext
  Loop
  AddLine "    Table.EndTable"
  AddLine ""
  AddLine "  URL.EndURL"
  AddLine ""
  AddLine "End Sub"
  AddLine "REM -------------------------------------------------------------------------"
  AddLine "REM Fim do ShowFirstPage                                                     "
  AddLine "REM ========================================================================="
  AddLine ""

End Sub
REM -------------------------------------------------------------------------
REM Fim do MakeFirstPage
REM =========================================================================

REM =========================================================================
REM Procedimento que escreve o c�digo de constru��o da estrutura de dados do
REM sistema
REM -------------------------------------------------------------------------
Private Sub MakeDataDefining

  ' Mostra ao usu�rio o que est� sendo feito agora
  ShowStatus "Gerando a estrutura de dados do sistema"

  ' Declara as vari�veis locais
  Dim strAnterior, strPrimary, blnTeste, arrDataTypes( 11 ), arrRequireds( 1 )

  ' Define o tipo dos requerimento
  arrRequireds( 0 ) = "efNotRequired"
  arrRequireds( 1 ) = "efRequired"

  ' Define o tipo dos campos
  arrDataTypes(  0 ) = "efDataTypeChar"
  arrDataTypes(  1 ) = "efDataTypeVarChar"
  arrDataTypes(  2 ) = "efDataTypeDateTime"
  arrDataTypes(  3 ) = "efDataTypeText"
  arrDataTypes(  4 ) = "efDataTypeInt"
  arrDataTypes(  5 ) = "efDataTypeSmallInt"
  arrDataTypes(  6 ) = "efDataTypeTinyInt"
  arrDataTypes(  7 ) = "efDataTypeFloat"
  arrDataTypes(  8 ) = "efDataTypeReal"
  arrDataTypes(  9 ) = "efDataTypeMoney"
  arrDataTypes( 10 ) = "efDataTypeBit"
  arrDataTypes( 11 ) = "efDataTypeInt"

  ' Recupera os dados das tabelas e dos campos
  OpenRecordSet( conSelectAll & "zeuTabelas INNER JOIN zeuCampos ON " & _
                 "( zeuTabelas.tabCodigo = zeuCampos.camTabela ) WHERE " & _
                 "zeuTabelas.tabSistema = " & sstrSistema & " ORDER BY tabCodigo, camOrdem" )

  ' Percorre toda a fonte de dados e descreve todas as tabelas que formam o sistema
  Do While Not sobjRecordSet.EOF

     If ( strAnterior <> sobjRecordSet( "tabNome" ) ) Then
        AddLine "REM ========================================================================="
        AddLine "REM Procedimento que define a estrutura de dados tabela """ & sobjRecordSet( "tabNome" ) & """"
        AddLine "REM -------------------------------------------------------------------------"
        AddLine "Public Sub data" & sobjRecordSet( "tabNome" )
        AddLine ""
        AddLine "  Edit.DataTable """ & sstrSigla & sobjRecordSet( "tabNome" ) & """"
        AddLine ""
     End If

     AddLine "  Edit.DataAddField " & _
             """" & _
             sobjRecordSet( "tabSigla" ) & _
             LeftStr ( sobjRecordSet( "camNome" ) & """,", 25 ) & " " & _
             LeftStr ( arrDataTypes( sobjRecordSet( "camTipo" ) ) & "," , 18 )  & " " & _
             RigthStr( sobjRecordSet( "camTamanho" ) & ",", 5 ) & " " & _
             arrRequireds( sobjRecordSet( "camRequerido" ) )
     If ( sobjRecordSet( "camPrimario" ) > 0 ) Then
        strPrimary = "  Edit.DataAddPrimaryKey """ & _
                        sobjRecordSet( "tabSigla" ) & _
                        sobjRecordSet( "camNome" ) & """"
     End If

     ' Armazena o nome da tabela para ser utilizada com crit�rio de quebra
     strAnterior = sobjRecordSet( "tabNome" )

     ' Vai para o pr�ximo registro
     sobjRecordSet.MoveNext

     If sobjRecordSet.EOF Then
        blnTeste = True

     Else
        blnTeste = strAnterior <> sobjRecordSet( "tabNome" )

     End If


     If blnTeste Then

        AddLine ""

        ' Se foi definido um campo prim�rio
        If ( strPrimary > "" ) Then

           ' Adiciona o campo prim�rio � estrutura da tabela
           AddLine strPrimary

        ' Se n�o foi definido um campo prim�rio
        Else

           ' Avisa ao usu�rio que falta o campo prim�rio
           AddLine "  Show.Message ""N�o foi definido um campo prim�rio para esta tabela!"""

        End If

        AddLine ""
        AddLine "  If ( Not Edit.HasTable( """ & sstrSigla & strAnterior & """ ) ) Then"
        AddLine "      Edit.CreateTable """ & sstrSigla & strAnterior & """"
        AddLine "  End If"
        AddLine ""
        AddLine "End Sub"
        AddLine "REM -------------------------------------------------------------------------"
        AddLine "REM Fim do data" & strAnterior
        AddLine "REM ========================================================================="
        AddLine ""

        strPrimary = ""

     End If

  Loop

End Sub
REM -------------------------------------------------------------------------
REM Fim do MakeDataDefining
REM =========================================================================

REM =========================================================================
REM Procedimento que escreve o c�digo de constru��o dos formul�rios do
REM sistema
REM -------------------------------------------------------------------------
Private Sub MakeEditForms

  ' Mostra ao usu�rio o que est� sendo feito agora
  ShowStatus "Gerando os formul�rios do sistema"

  ' Declara as vari�veis locais
  Dim strAnterior, strUnit, strFind, strList, strFields
  Dim blnTeste
  Dim arrFldType( 15 ), arrValOpt( 11 )
  Dim strFieldsLine

  ' Define o tipo dos campos
  arrFldType(  1 ) = "efFldTypeText"
  arrFldType(  2 ) = "efFldTypeUF"
  arrFldType(  3 ) = "efFldTypeLookup"
  arrFldType(  4 ) = "efFldTypeCheck"
  arrFldType(  5 ) = "efFldTypeTextArea"
  arrFldType(  6 ) = "efFldTypeRadio"
  arrFldType(  7 ) = "efFldTypeSelect"
  arrFldType(  8 ) = "efFldTypePassword"
  arrFldType(  9 ) = "efFldTypeHTTP"
  arrFldType( 10 ) = "efFldTypeEMail"
  arrFldType( 11 ) = "efFldTypeSeparateDate"
  arrFldType( 12 ) = "efFldTypeImage"
  arrFldType( 13 ) = "efFldTypeFile"
  arrFldType( 14 ) = "efFldTypeAtualization"
  arrFldType( 15 ) = "efFldTypeColor"

  arrValOpt(  1 ) = "efValOptNone"
  arrValOpt(  2 ) = "efValOptCGC"
  arrValOpt(  3 ) = "efValOptCPF"
  arrValOpt(  4 ) = "efValOptDate"
  arrValOpt(  5 ) = "efValOptSepDate"
  arrValOpt(  6 ) = "efValOptDateMToday"
  arrValOpt(  7 ) = "efValOptSepDateMToday"
  arrValOpt(  8 ) = "efValOptTime"
  arrValOpt(  9 ) = "efValOptEmail"
  arrValOpt( 10 ) = "efValOptCompareDates"
  arrValOpt( 11 ) = "efValOptCEP"

  ' Posicionando no in�cio da consulta, levo em considera��o que a �ltima
  ' tabela a ser aberta foi na procedure e "MakeDataDefining"
  sobjRecordSet.MoveFirst

  ' Percorre toda a fonte de dados e descreve todas as tabelas que formam o sistema
  Do While Not sobjRecordSet.EOF

     If ( strAnterior <> sobjRecordSet( "tabNome" ) ) Then
        AddLine "REM ========================================================================="
        AddLine "REM Inicio do procedimento que desenha o formul�rio de edi��o da tabela"
        AddLine "REM """ & sobjRecordSet( "tabNome" ) & """"
        AddLine "REM -------------------------------------------------------------------------"
        AddLine "Public Sub form" & sobjRecordSet( "tabNome" )
        AddLine ""
        AddLine "  data"  & sobjRecordSet( "tabNome" )
        AddLine ""
        If sobjRecordSet( "tabDescricao" ) > "" Then
           AddLine "  Edit.FormBegin  """ & sstrSigla & sobjRecordSet( "tabNome" ) & """, """ & sobjRecordSet( "tabDescricao" ) &  """, 1, Edit.parWhat, efValLocClient"
        Else
           AddLine "  Edit.FormBegin  """ & sstrSigla & sobjRecordSet( "tabNome" ) & """, """ & sobjRecordSet( "tabNome" ) &  """, 1, Edit.parWhat, efValLocClient"
        End If
        AddLine ""

        ' Inicializa a defini��o do "FormUnit"
        strUnit = "    Edit.FormUnit """

        ' Inicializa a defini��o do "FormFind"
        strFind = "    Edit.FormFind """

        ' Inicializa a defini��o do "FormList"
        strList = "    Edit.FormList """

        ' Inicializa a defini��o dos "AddField"
        strFields = ""

     End If

     If ( sobjRecordSet( "camMostrar" ) <> 0 ) Then
        strUnit = strUnit & _
                  sobjRecordSet( "tabSigla" ) & _
                  sobjRecordSet( "camNome" ) & sobjRecordSet( "camDelimitador" )

        If sobjRecordSet( "camTexto" ) > "" Then
          strUnit = strUnit & Trim(sobjRecordSet( "camTexto" )) & " "

        End If

     End If

     If ( sobjRecordSet( "camLocalizacao" ) <> 0 ) Then
        strFind = strFind & _
                  sobjRecordSet( "tabSigla" ) & _
                  sobjRecordSet( "camNome" ) & ","
     End If

     If ( sobjRecordSet( "camListagem" ) <> 0 ) Then
        strList = strList & _
                  sobjRecordSet( "tabSigla" ) & _
                  sobjRecordSet( "camNome" ) & ","
     End If

     strFieldsLine = ""

     If Not IsNull(sobjRecordSet( "camMsgObrigatorio" )) Then
       strFieldsLine = """" & sobjRecordSet( "camMsgObrigatorio" ) & """"

     End If

     If strFieldsLine > "" Then
       strFieldsLine = ", " & strFieldsLine

     End If


     If Not IsNull(sobjRecordSet( "camPadrao" )) Then
       strFieldsLine = sobjRecordSet( "camPadrao" ) & strFieldsLine

     End If

     If strFieldsLine > "" Then
       strFieldsLine = ", , " & strFieldsLine

     End If

     If ( Not IsNull( sobjRecordSet( "camTipoValidacao" ) ) ) And ( sobjRecordSet( "camTipoValidacao" ) > 1 ) Then
       strFieldsLine = arrValOpt(sobjRecordSet( "camTipoValidacao" )) & strFieldsLine
     End If

     If strFieldsLine > "" Then
       strFieldsLine = ", " & strFieldsLine

     End If

     If ( Not IsNull( sobjRecordSet( "camTipoEdicao" ) ) ) And ( sobjRecordSet( "camTipoEdicao" ) > 1 ) Then
       strFieldsLine = arrFldType(sobjRecordSet( "camTipoEdicao" )) & strFieldsLine


     End If

     If strFieldsLine > "" Then
       strFieldsLine = ", " & strFieldsLine

     End If

     strFields = strFields & "      Edit.AddField """ & _
                             sobjRecordSet( "tabSigla" ) & _
                             sobjRecordSet( "camNome" ) & """, " & _
                             """" & sobjRecordSet( "camRotulo" ) & """" & _
                             strFieldsLine & vbCRLF


     If ( sobjRecordSet( "camMaximo" ) > 0 ) Then
        strFields = strFields & "        Edit.FieldShowSize """ & _
                              sobjRecordSet( "tabSigla" ) & _
                              sobjRecordSet( "camNome" ) & """, " & _
                              sobjRecordSet( "camMaximo" ) & vbCRLF
     End If


     ' Armazena o nome da tabela para ser utilizada com crit�rio de quebra
     strAnterior = sobjRecordSet( "tabNome" )

     ' Vai para o pr�ximo registro
     sobjRecordSet.MoveNext

     If sobjRecordSet.EOF Then
        blnTeste = True

     Else
        blnTeste = strAnterior <> sobjRecordSet( "tabNome" )

     End If

     If blnTeste Then
        AddLine Left(strUnit, Len(strUnit) - 1) & """"
        AddLine Left(strFind, Len(strFind) - 1) & """"
        AddLine Left(strList, Len(strList) - 1) & """"
        AddLine ""
        AddLine strFields
        AddLine "  Edit.FormEnd"
        AddLine ""
        AddLine "End Sub"
        AddLine "REM -------------------------------------------------------------------------"
        AddLine "REM Fim do """ & strAnterior
        AddLine "REM ========================================================================="
        AddLine ""
     End If

  Loop


End Sub
REM -------------------------------------------------------------------------
REM Fim do MakeEditForms
REM =========================================================================

REM =========================================================================
REM Procedimento que escreve o c�digo de constru��o do corpo principal do
REM sistema
REM -------------------------------------------------------------------------
Private Sub MakeMainBody

  ' Mostra ao usu�rio o que est� sendo feito agora
  ShowStatus "Gerando o corpo principal do sistema"

  ' Recupera os dados do sistema e da empresa
  OpenRecordSet( conSelectAll & "zeuTabelas WHERE tabSistema = " & sstrSistema )

  AddLine "REM ========================================================================="
  AddLine "REM Corpo Principal do sistema                                               "
  AddLine "REM -------------------------------------------------------------------------"
  AddLine "Private Sub MainBody"
  AddLine ""
  AddLine "  FormatTable"
  AddLine ""
  AddLine "  Edit.ConnectionString = Session( ""ConnectionString"" )"
  AddLine ""
  AddLine "  Edit.OpenConnection"
  AddLine ""
  AddLine "  Select Case Edit.parWhat"
  Do While Not sobjRecordSet.EOF
     AddLine ""
     AddLine "    Case conWhat" & sobjRecordSet( "tabNome" )
     AddLine "         form"  & sobjRecordSet( "tabNome" )
     sobjRecordSet.MoveNext
  Loop
  AddLine ""
  AddLine "    Case Else"
  AddLine ""
  AddLine "  End Select"
  AddLine ""
  AddLine "  Edit.RedirectActions"
  AddLine ""
  AddLine "  Default.BeginHTML"
  AddLine "  Default.BeginBody"
  AddLine "  Default.BeginBody"
  AddLine ""
  AddLine "	   If ( Edit.IsMyAction ) Then"
  AddLine "	      ShowFirstPage"
  AddLine "	   End If"
  AddLine ""
  AddLine "  Default.PageFooterDefault"
  AddLine "  Default.EndBody"
  AddLine "  Default.EndHTML"
  AddLine ""
  AddLine "End Sub"
  AddLine "REM -------------------------------------------------------------------------"
  AddLine "REM Final da Sub MainBody                                                    "
  AddLine "REM ========================================================================="
  AddLine ""

End Sub
REM -------------------------------------------------------------------------
REM Fim do MakeMainBody
REM =========================================================================

REM =========================================================================
REM Procedimento que escreve o final do arquivo
REM -------------------------------------------------------------------------
Private Sub MakeEndFile

  ' Mostra ao usu�rio o que est� sendo feito agora
  ShowStatus "Finalzando a escrita do arquivo"

  ' Escreve o c�digo do final do arquivo
  AddLine "REM -------------------------------------------------------------------------"
  AddLine "REM Fim do " & sstrArquivo
  AddLine "REM ========================================================================="
  AddLine "%" & ">"

End Sub
REM -------------------------------------------------------------------------
REM Fim do MakeEndFile
REM =========================================================================

REM =========================================================================
REM Procedimento que finaliza o ambiente
REM -------------------------------------------------------------------------
Private Sub EndPrepare

  ' Fecha os objetos abertos
  sobjOutputFile.Close
  sobjConnection.Close

  If ( sobjRecordSet.State = adStateOpen ) Then
     sobjRecordSet.Close
  End If

  ' Destr�i os objetos constru�dos
  Set sobjOutputFile = Nothing
	Set sobjFileSystem = Nothing
  Set sobjConnection = Nothing
  Set sobjRecordSet  = Nothing

End Sub
REM -------------------------------------------------------------------------
REM Fim do EndPrepare
REM =========================================================================

REM =========================================================================
REM Grava uma linha no arquivo aberto
REM -------------------------------------------------------------------------
Private Sub AddLine( ByVal strLine )

  ' Se o objeto de controle de arquivo n�o foi criado ainda
  if ( IsEmpty( sobjOutputFile ) ) then

     ' Crio o arquivo de sa�da
     Set sobjOutputFile = sobjFileSystem.CreateTextFile( sstrResultFile, True )

  end if

  ' Grave uma linha no arquivo aberto "sobjFile"
  sobjOutputFile.WriteLine strLine

End Sub
REM -------------------------------------------------------------------------
REM Fim do AddLine
REM =========================================================================

REM =========================================================================
REM Abre uma fonte de dados
REM -------------------------------------------------------------------------
Private Sub OpenRecordSet( strSQLCommand )

  ' Fecha a fonte de dados
  If ( sobjRecordSet.State = adStateOpen ) Then
     sobjRecordSet.Close
  End If

  ' Recurpera os dados da fonte de dados
  sobjRecordSet.Open strSQLCommand, sobjConnection ', adOpenKeySet, adLockReadOnly

End Sub
REM -------------------------------------------------------------------------
REM Fim do OpenRecordSet
REM =========================================================================

REM =========================================================================
REM Mostra uma linha de estado indicando em que est�gio estou na gera��o do
REM sistema
REM -------------------------------------------------------------------------
Private Sub ShowStatus( strShow )

  intTotal = intTotal + 1

  ' Mostra ao usu�rio o que est� sendo feito agora
  Show.HTML intTotal & conTotalBlocks & strShow & "...<BR>"

End Sub
REM -------------------------------------------------------------------------
REM Fim do ShowStatus
REM =========================================================================

REM =========================================================================
REM Retorna uma "String" alinhado � esquerda dentro do espa�o definido
REM -------------------------------------------------------------------------
Private Function LeftStr( strShow, intSize )

  If ( intSize - Len( strShow ) ) < 0 Then
     LeftStr = strShow
  Else
     LeftStr = strShow & Space( intSize - Len( strShow ) )
  End If

End Function
REM -------------------------------------------------------------------------
REM Fim do LeftStr
REM =========================================================================

REM =========================================================================
REM Retorna uma "String" alinhada � direita dentro do espa�o definido
REM -------------------------------------------------------------------------
Private Function RigthStr( strShow, intSize )

  If ( intSize - Len( strShow ) ) < 0 Then
     RigthStr = strShow
  Else
     RigthStr = Space( intSize - Len( strShow ) ) & strShow
  End If


End Function
REM -------------------------------------------------------------------------
REM Fim do RigthStr
REM =========================================================================

REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub DoMakeSystem( ByVal strSystem )

    sstrSistema = strSystem

    ' Prepara o caminho para a execu��o do c�digo
  	Prepare

  	' Constr�i o cabe�alho
  	MakeHead

  	' Constr�i a �rea global
  	MakeGlobalArea

  	' Constr�i o a tela principal
  	MakeFirstPage

  	' Constr�i a estrutura de dados do sistema
  	MakeDataDefining

  	' Constr�i os formul�rios do sistema
  	MakeEditForms

  	' Constr�i o ambiente do sistema
  	MakeEnvironment

  	' Constr�i o corpo principal do sistema
  	MakeMainBody

  	' Constr�i o final do arquivo
  	MakeEndFile

    ' Finaliza a execu��o do c�digo
  	EndPrepare

End Sub
REM -------------------------------------------------------------------------
REM Final do DoMakeSystem
REM =========================================================================

REM -------------------------------------------------------------------------
REM Fim do MakeSystem.inc
REM =========================================================================
%>


